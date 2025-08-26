import pandas as pd
import streamlit as st
from datetime import datetime, time, date
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File
import io
from notion_client import Client

st.title('Conciliacion Instant - Payouts')

#=========================================
# Accesos Sharepoint - Notion
#=========================================

site_url = "https://kashioinc.sharepoint.com/sites/Intranet2021"
username = "dussand.hurtado@kashio.net"
password = "Silvana1505$"

#=========================================
# Primera parte. Subida y lectura de archivo METABASE
#=========================================

st.header('Metabase')

# Inicializamos el session_state para mantener los datos entre ejecuciones
if 'ipayouts_data' not in st.session_state:
    st.session_state.ipayouts_data = None

if 'ipayouts_data_despues_corte' not in st.session_state:
    st.session_state.ipayouts_data_despues_corte = None

# AGREGAR: Inicializar df_pendientes en session_state
if 'df_pendientes' not in st.session_state:
    st.session_state.df_pendientes = None

# AGREGAR: Flag para controlar si ya se procesaron los pendientes
if 'pendientes_procesados' not in st.session_state:
    st.session_state.pendientes_procesados = False

#Subimos el excel de metabase 
file_uploader_metabase = st.file_uploader('Arrastra el archivo de metabase aquí: ', type=['xlsx'], accept_multiple_files=True)

if file_uploader_metabase:
    #elegir archivo de metabase y pendientes
    if isinstance(file_uploader_metabase, list):
        ipayouts_metabase_df = file_uploader_metabase[0]
        if len(file_uploader_metabase) > 1:
            st.session_state.df_pendientes = file_uploader_metabase[1]
        else:
            st.session_state.df_pendientes = None

    else:
        ipayouts_metabase_df = file_uploader_metabase
        st.session_state.df_pendientes = None

    ipayouts_metabase_df = pd.read_excel(ipayouts_metabase_df) # cargamos el excel

    columns_drop = [
        'descripcion',
        'referencia',
        'payout process',
        'ID cliente',
        'correo cliente',
        'motivo'
    ]

    ipayouts_metabase_df.drop(columns=columns_drop, inplace=True) #eliminamos las columnas innecesarias
    ipayouts_metabase_df['documento'] = ipayouts_metabase_df['documento'].astype(str) #convertimos el documento en un str

    ipayouts_metabase_df['fecha_creacion'] = ipayouts_metabase_df['fecha creacion'].dt.date

    alcance_bancos = [
        '(BCP) - Banco de Crédito del Perú',
        'Yape',
        '(BBVA) - BBVA Continental '
    ]

    ipayouts_metabase_df = ipayouts_metabase_df[ipayouts_metabase_df['banco'].isin(alcance_bancos)] #filtramos los bancos que vamos a usar

    ipayouts_metabase_df = ipayouts_metabase_df[ipayouts_metabase_df['estado'] == 'Pagado']

    def extraer_codigo(row): #definicimos una funcion para extraer el codigo operacion de la columna numero de operacion de metabase
        banco = row['banco']
        concepto = str(row['numero de operacion'])
        monto = str(row['monto'])

        if banco == '(BCP) - Banco de Crédito del Perú':
             codigo = concepto[18:27]
        elif banco == '(BBVA) - BBVA Continental ':
             codigo = concepto[:20]
        elif banco == 'Yape':
             codigo = concepto[-11:]
        else:
             None
        
        #limpiara y tomar los dos primeros digitos del monto
        monto_cuatro_digitos = monto.replace('.', '').replace(',','')[:4]

        return f'{codigo}{monto_cuatro_digitos}'
        
    ipayouts_metabase_df['codigo_operacion'] = ipayouts_metabase_df.apply(extraer_codigo, axis=1) #aplicamos la funcion

    if st.session_state.ipayouts_data is None:
        st.session_state.ipayouts_data = ipayouts_metabase_df.copy()
        st.session_state.pendientes_procesados = False
        #st.info("Datos del archivo cargados")

#======================================================
    # # Selector de fecha

    # Inicializamos una sola vez
    if "fecha_sel" not in st.session_state:
        st.session_state.fecha_sel = date.today() - pd.Timedelta(days=1)

    fecha_sel= st.date_input("SELECCIONAR FECHA DE DIA DE CONCILIACION: ", value=st.session_state.fecha_sel, key='fecha_sel') #seleccionar fecha del dia que se va a conciliar, o sea la de ayer

    # Detectar si la fecha cambió y resetear automáticamente
    if "ultima_fecha_sel" not in st.session_state:
        st.session_state.ultima_fecha_sel = fecha_sel

    if st.session_state.ultima_fecha_sel != fecha_sel:
        # La fecha cambió, resetear los datos relacionados con la fecha
        st.session_state.ayer_corte = (fecha_sel - pd.Timedelta(days=1))
        st.session_state.ipayouts_data_despues_corte = None
        st.session_state.ultima_fecha_sel = fecha_sel


    if 'ayer_corte' not in st.session_state:
        st.session_state.ayer_corte = (st.session_state.fecha_sel - pd.Timedelta(days=1))

    ayer_para_cortes = st.session_state.ayer_corte
    fecha_sel = st.session_state.fecha_sel
    
    #ayer = (fecha_sel - pd.Timedelta(days=1))

    año = fecha_sel.year
    mes_formateado = fecha_sel.strftime('%m_%B')
    nombre_archivo = f"Pendiente_Conciliar_{ayer_para_cortes}.xlsx"

    if st.session_state.df_pendientes is not None and not st.session_state.pendientes_procesados:
        try:
            df_pendientes = pd.read_excel(st.session_state.df_pendientes)
            df_pendientes['fecha_creacion'] = df_pendientes['fecha creacion'].dt.date
            df_pendientes['codigo_operacion'] = df_pendientes.apply(extraer_codigo, axis=1)

            st.session_state.ipayouts_data  = pd.concat([st.session_state.ipayouts_data , df_pendientes], ignore_index=True)
            # ipayouts_metabase_df = ipayouts_metabase_df.drop_duplicates(subset=['codigo_operacion'])

            st.session_state.pendientes_procesados = True
            st.success(f"Se agregaron los movimientos pendientes del {ayer_para_cortes}.")

        except Exception as e:
            st.warning(f"No se pudo cargar pendientes del {ayer_para_cortes}: {e}")


#=====================================================

    hora_corte_bbva = time(22, 00) # 22:00 pm
    hora_corte_bcp_yape = time(21, 15)
        
    # Combinar fecha de ayer con hora de corte
    dt_corte_bbva = datetime.combine(fecha_sel, hora_corte_bbva)
    dt_corte_bcp_yape = datetime.combine(fecha_sel, hora_corte_bcp_yape)

    # Crear columna con datetime de corte según banco
    st.session_state.ipayouts_data['corte_datetime'] = st.session_state.ipayouts_data['banco'].apply(
        lambda b: dt_corte_bbva if '(BBVA) - BBVA Continental ' in b else dt_corte_bcp_yape
    )

        # Crear columna 'estado_corte'
    st.session_state.ipayouts_data['estado_corte'] = st.session_state.ipayouts_data.apply(
        lambda row: 'Antes de corte'
        if row['fecha creacion'] < datetime.combine(
            fecha_sel,
            hora_corte_bbva if 'BBVA' in row['banco'] else hora_corte_bcp_yape
        )
        else 'Después de corte',
        axis=1
    )

    # Separar en dos DataFrames antes de aplicar el filtro
    if st.session_state.ipayouts_data_despues_corte is None:
        movimientos_pendientes = st.session_state.ipayouts_data[st.session_state.ipayouts_data['estado_corte'] == 'Después de corte'].copy()
        st.session_state.ipayouts_data_despues_corte = movimientos_pendientes.sort_values('fecha creacion' , ascending=True)


    # # Filtrar por operaciones ANTES de la hora de corte de ayer
    st.session_state.ipayouts_data = st.session_state.ipayouts_data[st.session_state.ipayouts_data['estado_corte'] == 'Antes de corte']
    st.session_state.ipayouts_data.sort_values('fecha creacion', ascending=True, inplace=True) #ordenamos la fecha creacion de menor a mayor

    montos_ipayouts = st.session_state.ipayouts_data.groupby(['banco'])['monto'].sum().reset_index() #armamos un pivot para revisar los montos 
    st.session_state.ipayouts_data  
    # st.session_state.ipayouts_data_despues_corte 
    st.dataframe(montos_ipayouts, use_container_width=True)

#================================================================
#  Definicion de funcion para guardado de registros pendientes
#================================================================


    def guardar_conciliacion(movimiento_pendientes):
        status_placeholder = st.empty()

        #with st.spinner():

        status_placeholder.info('Conectando a Sharepoint...')

        try:
            # Conectamos al sitio
            ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
            
            # Ruta relativa a la carpeta en SharePoint (CORREGIDA: debe empezar con /sites/...)
            folder_url = "/sites/Intranet2021/Shared Documents/Operaciones/PAYOUT/PAYOUTS VARIOS/Conciliaciones Instant Payout/Registros pendientes"

            # Verificamos que la carpeta exista
            folder = ctx.web.get_folder_by_server_relative_url(folder_url)
            ctx.load(folder)
            ctx.execute_query()

            # Si todo fue bien:
            #st.success(" Conectado correctamente a la carpeta 'Conciliaciones Payout'")

        except Exception as e:
            st.error(f"No se pudo conectar: {e}")
            return  # Salir si no se puede conectar

        # Obtener el año actual 
        año_actual = datetime.now().year #para la carpeta de año
        mes_actual = datetime.now().strftime('%m_%B') #para la carpeta de mes
        #archivo_nombre = ayer.strftime('Conciliacion_%Y_%m_%d.xlsx')
        archivo_nombre = f'Pendiente_Conciliar_{fecha_sel}.xlsx' #in case doesn't work, delete this

        
        # Rutas de las carpetas del año y mes (CORREGIDAS)
        nueva_carpeta_año = f'{folder_url}/{año_actual}'
        nueva_carpeta_mes = f'{nueva_carpeta_año}/{mes_actual}'

        status_placeholder.info(f'Verificando carpeta del año {año_actual}...')

        # Verificamos si existe la carpeta del año
        try:
            folder_año = ctx.web.get_folder_by_server_relative_url(nueva_carpeta_año)
            ctx.load(folder_año)
            ctx.execute_query()
            #st.info(f'La carpeta del año {año_actual} ya existe')
        except:
            try:
                folder_base = ctx.web.get_folder_by_server_relative_url(folder_url)
                folder_base.folders.add(str(año_actual))  # Convertir a string
                ctx.execute_query()
                #st.success(f'La carpeta del año {año_actual} creada exitosamente')
            except Exception as e:
                st.error(f'Error al crear la carpeta del año {año_actual}: {e}')
                return
        
        status_placeholder.info(f'Verificando carpeta del mes {mes_actual}...')

        # Verificamos si la carpeta del mes ya existe
        try:
            folder_mes = ctx.web.get_folder_by_server_relative_url(nueva_carpeta_mes)
            ctx.load(folder_mes)
            ctx.execute_query()
            #st.info(f"La carpeta del mes {mes_actual} ya existe.")
        except:
            try:
                folder_anio = ctx.web.get_folder_by_server_relative_url(nueva_carpeta_año)
                folder_anio.folders.add(mes_actual)
                ctx.execute_query()
                #st.success(f"Carpeta del mes {mes_actual} creada exitosamente.")
            except Exception as e:
                st.error(f"Error al crear la carpeta del mes {mes_actual}: {e}")
                return
            

        status_placeholder.info(f'Preparando archivo excel...')

        # Guardar archivo CSV con nombre del día de ayer
        try:
            # CORREGIDO: Ruta completa para el archivo
            ruta_archivo_completa = f"{nueva_carpeta_mes}/{archivo_nombre}"
            
            # Convertimos ambos DataFrames a Excel en memoria
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                # Guardar el primer DataFrame (Payouts Metabase)
                movimiento_pendientes.to_excel(writer, sheet_name=f'Pendientes_{fecha_sel}', index=False)
                #st.info(f"📊 Hoja 1: '{nombre_primera_hoja}' - {len(payouts_metabase_df)} registros")
                
                # # Guardar el segundo DataFrame (df_final)
                # df_final.to_excel(writer, sheet_name=nombre_segunda_hoja, index=False)
                # #st.info(f"📊 Hoja 2: '{nombre_segunda_hoja}' - {len(df_final)} registros")
                    
            excel_content = excel_buffer.getvalue()

            status_placeholder.info('Subiendo archivo a SharePoint...')

            #st.write("📂 Ruta final de guardado:", ruta_archivo_completa)
            
            # MÉTODO CORREGIDO: Usar upload_file en lugar de File.save_binary
            target_folder = ctx.web.get_folder_by_server_relative_url(nueva_carpeta_mes)
            target_folder.upload_file(archivo_nombre, excel_content).execute_query()

            status_placeholder.empty()
            
            st.success(f"Archivo '{archivo_nombre}' guardado correctamente en SharePoint.")
            
        except Exception as e:
            st.error(f"Error al guardar el archivo: {e}")

            status_placeholder.info('Intentando metodo alternativo...')
            
            # Método alternativo si el anterior falla
            try:
                #st.info("🔄 Intentando método alternativo...")
                
                # Método alternativo usando File.save_binary con ruta completa
                File.save_binary(ctx, ruta_archivo_completa, excel_content)

                #impiar el placeholder del estado
                status_placeholder.empty()

                #Mensaje de exito
                #st.success(f"Archivo '{archivo_nombre}' guardado con método alternativo (2 hojas).")
                
            except Exception as e2:
                status_placeholder.empty()
                st.error(f"Error también con método alternativo: {e2}")
                
                # Mostrar información de debug
                st.write("🔍 **Información de debug:**")
                st.write(f"- Ruta completa: {ruta_archivo_completa}")
                st.write(f"- Nombre archivo: {archivo_nombre}")
                st.write(f"- Carpeta mes: {nueva_carpeta_mes}")
                st.write(f"- Tamaño Excel: {len(excel_content)} bytes")

    def guardar_registros_pagados(codigos_encontrados):
        status_placeholder = st.empty()

        #with st.spinner():

        status_placeholder.info('Conectando a Sharepoint...')

        try:
            # Conectamos al sitio
            ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
            
            # Ruta relativa a la carpeta en SharePoint (CORREGIDA: debe empezar con /sites/...)
            folder_url = "/sites/Intranet2021/Shared Documents/Operaciones/PAYOUT/PAYOUTS VARIOS/Conciliaciones Instant Payout/Registros conciliados"

            # Verificamos que la carpeta exista
            folder = ctx.web.get_folder_by_server_relative_url(folder_url)
            ctx.load(folder)
            ctx.execute_query()

            # Si todo fue bien:
            #st.success(" Conectado correctamente a la carpeta 'Conciliaciones Payout'")

        except Exception as e:
            st.error(f"No se pudo conectar: {e}")
            return  # Salir si no se puede conectar

        # Obtener el año actual 
        año_actual = datetime.now().year #para la carpeta de año
        mes_actual = datetime.now().strftime('%m_%B') #para la carpeta de mes
        #archivo_nombre = ayer.strftime('Conciliacion_%Y_%m_%d.xlsx')
        archivo_nombre = f'OperacionesPagadas_{fecha_sel}.parquet' #in case doesn't work, delete this

        
        # Rutas de las carpetas del año y mes (CORREGIDAS)
        nueva_carpeta_año = f'{folder_url}/{año_actual}'
        nueva_carpeta_mes = f'{nueva_carpeta_año}/{mes_actual}'

        status_placeholder.info(f'Verificando carpeta del año {año_actual}...')

        # Verificamos si existe la carpeta del año
        try:
            folder_año = ctx.web.get_folder_by_server_relative_url(nueva_carpeta_año)
            ctx.load(folder_año)
            ctx.execute_query()
            #st.info(f'La carpeta del año {año_actual} ya existe')
        except:
            try:
                folder_base = ctx.web.get_folder_by_server_relative_url(folder_url)
                folder_base.folders.add(str(año_actual))  # Convertir a string
                ctx.execute_query()
                #st.success(f'La carpeta del año {año_actual} creada exitosamente')
            except Exception as e:
                st.error(f'Error al crear la carpeta del año {año_actual}: {e}')
                return
        
        status_placeholder.info(f'Verificando carpeta del mes {mes_actual}...')

        # Verificamos si la carpeta del mes ya existe
        try:
            folder_mes = ctx.web.get_folder_by_server_relative_url(nueva_carpeta_mes)
            ctx.load(folder_mes)
            ctx.execute_query()
            #st.info(f"La carpeta del mes {mes_actual} ya existe.")
        except:
            try:
                folder_anio = ctx.web.get_folder_by_server_relative_url(nueva_carpeta_año)
                folder_anio.folders.add(mes_actual)
                ctx.execute_query()
                #st.success(f"Carpeta del mes {mes_actual} creada exitosamente.")
            except Exception as e:
                st.error(f"Error al crear la carpeta del mes {mes_actual}: {e}")
                return
            

        status_placeholder.info(f'Preparando archivo parquet...')

        # Guardar archivo CSV con nombre del día de ayer
        try:
            # CORREGIDO: Ruta completa para el archivo
            ruta_archivo_completa = f"{nueva_carpeta_mes}/{archivo_nombre}"
            
            # csv_buffer = io.StringIO()
            # codigos_encontrados.to_csv(csv_buffer, index=False)
            # csv_content = csv_buffer.getvalue().encode('utf-8')

            # status_placeholder.info('Subiendo archivo a SharePoint...')

            # #st.write("📂 Ruta final de guardado:", ruta_archivo_completa)
            
            # # MÉTODO CORREGIDO: Usar upload_file en lugar de File.save_binary
            # target_folder = ctx.web.get_folder_by_server_relative_url(nueva_carpeta_mes)
            # target_folder.upload_file(archivo_nombre, csv_content).execute_query()
            # Asegurar que la columna documento sea string

            if 'documento' in codigos_encontrados.columns:
                codigos_encontrados['documento'] = codigos_encontrados['documento'].astype('string').fillna('')

            parquet_buffer = io.BytesIO()
            codigos_encontrados.to_parquet(parquet_buffer, index=False, engine='pyarrow')
            parquet_content = parquet_buffer.getvalue()

            target_folder = ctx.web.get_folder_by_server_relative_url(nueva_carpeta_mes)
            target_folder.upload_file(archivo_nombre, parquet_content).execute_query()

            status_placeholder.info('Subiendo archivo a SharePoint...')

            status_placeholder.empty()
            
            st.success(f"Archivo '{archivo_nombre}' guardado correctamente en SharePoint.")
            
        except Exception as e:
            st.error(f"Error al guardar el archivo: {e}")

            status_placeholder.info('Intentando metodo alternativo...')
            
            # Método alternativo si el anterior falla
            try:
                #st.info("🔄 Intentando método alternativo...")
                
                # Método alternativo usando File.save_binary con ruta completa
                File.save_binary(ctx, ruta_archivo_completa, parquet_content)

                #impiar el placeholder del estado
                status_placeholder.empty()

                #Mensaje de exito
                #st.success(f"Archivo '{archivo_nombre}' guardado con método alternativo (2 hojas).")
                
            except Exception as e2:
                status_placeholder.empty()
                st.error(f"Error también con método alternativo: {e2}")
                
                # Mostrar información de debug
                st.write("🔍 **Información de debug:**")
                st.write(f"- Ruta completa: {ruta_archivo_completa}")
                st.write(f"- Nombre archivo: {archivo_nombre}")
                st.write(f"- Carpeta mes: {nueva_carpeta_mes}")
                st.write(f"- Tamaño Excel: {len(parquet_content)} bytes")

    def registros_notion(merge_meta_banco):
        notion_token = "ntn_OV8209261688Lu7hdNom52kNGhWwBLLfUIBY3z30uMNetm"
        database_id =  "248030ee56d8808da559c748ab6c0ee0"

        notion = Client(auth=notion_token)

        status_placeholder = st.empty()
        progress_bar = st.progress(0)

        for idx, (_,rows) in enumerate(merge_meta_banco.iterrows()):
            try:
                notion.pages.create(
                    parent={'database_id': database_id},
                    properties={
                        'FechaTexto': {
                            'rich_text': [{'text': {'content': str(rows.get('FechaTexto', ''))}}]
                        },
                        'BANCO': {
                            'title': [{'text': {'content': str(rows.get('BANCO', ''))}}]
                        },
                        'Monto Banco': {
                            'number': round(float(rows.get('Monto Banco', 0)), 2)
                        },
                        'Monto Kashio': {
                            'number': round(float(rows.get('Monto Kashio', 0)), 2)
                        },
                        'Diferencia': {
                            'number': round(float(rows.get('Diferencia', 0)), 2)
                        }              
                    }
                )

                progress = (idx + 1) / len(merge_meta_banco)

                progress_bar.progress(min(progress,1.0))
                status_placeholder.success(f'Registro {idx + 1} guardado correctamente')

            except Exception as e:
                status_placeholder.error(f'Registro {idx + 1} falló: {e}')

  

#================================================================
# Segunda parte. Definicion de funciones para lecturas de eecc
#================================================================

    def procesar_bcp(estado_cuenta):
        """
        Procesa el estado de cuenta desde un archivo Excel.

        - Elimina columnas innecesarias.
        - Convierte la columna 'Operación - Número' a texto.
        - Crea una nueva columna 'codigo_operacion' basada en reglas de texto.

        Parámetros:
        estado_cuenta: archivo subido (por ejemplo, desde Streamlit file_uploader)

        Retorna:
        DataFrame procesado
        """
        # Leer Excel, omitiendo encabezados extras
        estado_cuenta_df = pd.read_excel(estado_cuenta, skiprows=4)

        # Columnas a eliminar
        columns_drop_eecc = [
            'Fecha valuta',
            'Saldo',
            'Sucursal - agencia',
            'Usuario',
            'UTC',
            'Referencia2'
        ]
        estado_cuenta_df.drop(columns=columns_drop_eecc, inplace=True)

        #renombramos columnas
        columnas_name = {'Fecha': 'fecha',
            'Descripción operación': 'descripcion_operacion',
            'Monto':'importe',
            'Operación - Número':'numero_operacion',
            'Operación - Hora':'hora_operacion'}
    

        estado_cuenta_df.rename(columns=columnas_name, inplace=True)

        # Asegurar que 'Operación - Número' es string
        estado_cuenta_df['numero_operacion'] = estado_cuenta_df['numero_operacion'].astype(str)

        def clasificacion_bancos(valor): #clasificamos bancos para filtrar filas innecesarias
            if valor.startswith('YPP'):
                return 'Yape'
            elif valor.startswith('A'):
                return "BCP"
            else:
                return "Otros"
            
        estado_cuenta_df['clasificacion_banco'] = estado_cuenta_df['descripcion_operacion'].apply(clasificacion_bancos)

        #filtramos solos las filas necesarias por la columna clasificacion bancos
        estado_cuenta_df = estado_cuenta_df[estado_cuenta_df['clasificacion_banco'] != 'Otros']

        # Crear columna 'codigo_operacion' con lógica condicional
        estado_cuenta_df['codigo_operacion'] = estado_cuenta_df.apply(
            lambda x: (
                str(x['descripcion_operacion'])[-11:] if str(x['descripcion_operacion'])[:3] == 'YPP'
                else str(x['numero_operacion']).zfill(8)
            ) + str(abs(x['importe']) * -1).replace('.', '').replace(',', '')[1:5],
            axis=1
        )


        #colocamos el nombre al banco para que muestre en el df final de bancos
        estado_cuenta_df['banco'] = estado_cuenta_df.apply(
            lambda x: 'Yape' if str(x['descripcion_operacion'])[:3] == 'YPP' else '(BCP) - Banco de Crédito del Perú',
            axis = 1
        )

        return estado_cuenta_df
    
    def procesar_bbva(estado_cuenta):
        """
        Procesa una variante del estado de cuenta desde un archivo Excel.

        - Omite las primeras 10 filas.
        - Elimina columnas innecesarias.
        - Convierte 'Nº. Doc.' a texto.
        - Extrae los últimos 20 caracteres de 'Concepto' como 'codigo_operacion'.

        Parámetros:
        ruta_excel: archivo subido (por ejemplo, desde Streamlit file_uploader)

        Retorna:
        DataFrame procesado
        """
        estado_cuenta_df = pd.read_excel(estado_cuenta, skiprows=10)
        
        
        # Eliminar columnas no requeridas
        columns_drop_eecc = [
            'F. Valor',
            'Código',
            'Oficina'
        ]
        estado_cuenta_df.drop(columns=columns_drop_eecc, inplace=True)

        #renombramos columnas
        columnas_name = {'F. Operación': 'fecha',
            'Nº. Doc.': 'descripcion_operacion',
            'Importe':'importe'}
        
        estado_cuenta_df.rename(columns=columnas_name, inplace=True)

        #eliminamos filas con valores nunlos en la columna fecha
        estado_cuenta_df = estado_cuenta_df[estado_cuenta_df['fecha'].notna()]

        #filtramos las filas necesarias, en este caso todas las que comiencen con *C/ PROV
        filtro = estado_cuenta_df['Concepto'].astype(str).str.startswith('*C/PROV')
        estado_cuenta_df = estado_cuenta_df[filtro] #aplicamos el filtro

        # Convertir a string para asegurar consistencia
        estado_cuenta_df['descripcion_operacion'] = estado_cuenta_df['descripcion_operacion'].astype(str)

        estado_cuenta_df['codigo_operacion'] = (
            estado_cuenta_df['Concepto'].astype(str).str[-20:] +
            estado_cuenta_df['importe'].apply(lambda x: str(abs(x) * -1)).str.replace('.', '').str.replace(',', '').str[1:5]
        )

        estado_cuenta_df['banco'] = '(BBVA) - BBVA Continental '
        return estado_cuenta_df


    #creamos el diccionario de funciones de cada banco
    procesadores_banck = {
        'bcp': procesar_bcp,
        'bbva': procesar_bbva
    }

#=============================================
# Tercera parte. Subida de estados de cuenta 
#=============================================

    st.header('Estados de cuenta')

    #creamos la seccion para subir el estado de cuenta del banco seleccionado
    estado_cuenta = st.file_uploader(f'Subir estados de cuenta', type=['xlsx', 'xls'], accept_multiple_files=True)


    df_consolidados = []

    if estado_cuenta:
        for archivo in estado_cuenta:
            nombre_archivo = archivo.name.lower()
            procesador = None
            #buscar funcion adecuada segun nombre de archivo
            for clave, funcion in procesadores_banck.items():
                if clave in nombre_archivo:
                    procesador = funcion
                    break

            if procesador:
                try:
                    df = procesador(archivo)
                    #st.dataframe(df)
                    df_consolidados.append(df)
                    st.success(f'Archivo procesado: {archivo.name}')
                except Exception as e:
                    st.error(f'Error al procesar {archivo.name}: {e}')
            else:
                st.warning(f'No se encontro una funcion para procesar: {archivo.name}')

    if st.session_state.ipayouts_data is not None and df_consolidados:

        df_final = pd.concat(df_consolidados, ignore_index=True) #Consolidamos todos los DF de los bancos BCP BBVA y Yape

        #mostramos solo las columnas necesarias
        df_final = df_final[['fecha', 'importe', 'codigo_operacion', 'banco']]
        df_final['fecha'] = pd.to_datetime(df_final['fecha']).dt.date

        #st.dataframe(df_final, use_container_width=True)

        #mostramos un pivot con los montos de los bancos
        montos_bancos_eecc = df_final.groupby(['fecha','banco'])['importe'].sum().abs().reset_index()
        st.dataframe(montos_bancos_eecc, use_container_width=True)

#============================================================
# Cuarta parte. Cruce de tablas para encontrar  diferencias
#============================================================

        st.header('Conciliacion')

        codigo_bancos_set = set(df_final['codigo_operacion']) # Crear un conjunto con los códigos de operación únicos del DataFrame df_final

        st.session_state.ipayouts_data['resultado_busqueda'] =  st.session_state.ipayouts_data['codigo_operacion'].apply(
            lambda x: x if x in codigo_bancos_set else 'No encontrado'
        ) #aplicamos una funcion que busca los codigos de operacion de metabase en el conjunto de codgios unicos y si lo encontra coloca el mismo y si no "no encontrado"

        st.subheader('Diferencias despues de cruce de numero de operacion')
        st.write(
            '''
            Esta tabla muestra las diferencias despues de el cruce de numeros de operacion entre el archivo metabase
            y los estados de cuenta subidos. Las diferencias encontradas vendrían a ser operaciones que se pagaron al día siguiente
            por lo que se deberá descargar y registrar los montos para poder conciliarlos el día de mañana. 

'''
        )
        if 'merge_realizado' not in st.session_state:
            st.session_state.merge_realizado = False

        if not st.session_state.merge_realizado:
        #hacemos un merge que me traiga el importe de los banccos respecto al codigo de operacion, desde el archivo de bancos
            st.session_state.ipayouts_data = st.session_state.ipayouts_data.merge(df_final[['codigo_operacion', 'importe']], left_on='codigo_operacion', right_on='codigo_operacion', how='left')
            st.session_state.merge_realizado = True

        #creamos una columna de saldo para revisar que no hayan operaciones con distintos importes. 
        st.session_state.ipayouts_data['saldo'] = (st.session_state.ipayouts_data['monto'] + st.session_state.ipayouts_data['importe']).fillna('No valor')
        #st.session_state.ipayouts_data

        #filtramos los codigos que se encontrarion 
        if 'codigos_encontrados_df' not in st.session_state:
            st.session_state.codigos_encontrados_df = None

        if st.session_state.codigos_encontrados_df is None:
            codigos_encontrados = st.session_state.ipayouts_data[st.session_state.ipayouts_data['resultado_busqueda'] != 'No encontrado']
        
            st.session_state.codigos_encontrados_df = codigos_encontrados
            
        # Alias local
        codigos_encontrados = st.session_state.codigos_encontrados_df  
        
        #creamos un pivot para mostrar los importes de los bancos por bancos
        codigos_encontrados_pivot = codigos_encontrados.groupby('banco')[['importe']].sum().reset_index()
        #unimos el df de metabase con lso bancos y montos y el df de los importe de lso bancos 
        merge_meta_banco = pd.merge(montos_ipayouts, codigos_encontrados_pivot, on='banco', how='inner')

        #creamos una columna de diferencias
        merge_meta_banco['Diferencia'] = merge_meta_banco['monto'] + merge_meta_banco['importe']

        rename_columns = {
            'fecha_creacion':'FechaTexto',
            'banco':'BANCO',
            'monto':'Monto Kashio',
            'importe':'Monto Banco'
        }
        
        merge_meta_banco = merge_meta_banco.rename(columns=rename_columns)
        merge_meta_banco.insert(0, 'FechaTexto', fecha_sel)
        st.dataframe(merge_meta_banco, use_container_width=True)

        if 'guardar_record_dif' not in st.session_state:
            st.session_state.guardar_record_dif = False

        registrar_diferencias_notion = st.button('REGISTRAR DIFERENCIAS', use_container_width=True)

        if not st.session_state.guardar_record_dif:
            if registrar_diferencias_notion:
                registros_notion(merge_meta_banco)
                st.session_state.guardar_record_dif = True

        # if registrar_diferencias_notion:

        with st.expander('Diferencias encontradas'):

            concicliacion_mañana_no_encontrado = st.session_state.ipayouts_data[st.session_state.ipayouts_data['resultado_busqueda'] == 'No encontrado'] #filtramos por los no encontrados

            bancos_unicos = ['Todos'] + sorted(st.session_state.ipayouts_data['banco'].unique()) #extraemos los valores unicos de los bancos

            bancos_unicos_sb = st.selectbox('Filtrar por banco', bancos_unicos) #creamos una lista desplegable con los bancos

            diferencias_filtro = merge_meta_banco[merge_meta_banco['BANCO'] == bancos_unicos_sb] 
            #aplicar el filtro segun seleccion

            if bancos_unicos_sb == 'Todos':
                concicliacion_mañana_filtrado = concicliacion_mañana_no_encontrado #filtramos por los bancos
                diferencias_filtro = merge_meta_banco
                
            else: 
                concicliacion_mañana_filtrado = concicliacion_mañana_no_encontrado[concicliacion_mañana_no_encontrado['banco'] == bancos_unicos_sb] #filtramos por el selectbox incluido


            #concicliacion_mañana_filtrado
            concicliacion_mañana_filtrado = concicliacion_mañana_filtrado[['empresa', 'fecha creacion','fecha operacion', 'inv public_id', 'po_public_id', 'Cliente'
                                                                                    , 'documento', 'numero de cuenta', 'CCI', 'monto', 'banco', 'numero de operacion', 'estado', 'codigo_operacion']] #mostramos las columnas necesarias
        

            st.dataframe(concicliacion_mañana_filtrado, use_container_width=True) 

            suma_monto =  round(concicliacion_mañana_filtrado['monto'].sum(), 2)

            suma_diferencias_filtro = round(diferencias_filtro['Diferencia'].sum(),2)
            
            diferencia_montos =round( suma_monto - suma_diferencias_filtro,2)

            cantidad_diferencias = len(concicliacion_mañana_filtrado) #numero de operacioens encontradas en la seccion de diferencias 


            if cantidad_diferencias == 0:
                st.success('Sin diferencias')    
            else:
                st.warning(f'{cantidad_diferencias} diferencias encontradas')
            
            
            if diferencia_montos == 0: #suma monto: suma de la columna monto del df de detalle de diferencias suma_diferencias: la suma de la columna Diferecnias del df de conciliacion
                st.success('Montos iguales')
            else:
                st.warning('Montos desiguales')

        
        # Inicializa el estado de guardado si no existe
        if 'guardad_registros_pendientes' not in st.session_state:
            st.session_state.guardad_registros_pendientes = False

        c1, c2 = st.columns(2)

        with c1:
            # cantidad_movimientos = len(st.session_state.ipayouts_data_despues_corte)
            # guardar_pospagos = st.button(f'GUARDAR {cantidad_movimientos} MOVIMIENTOS PENDIENTES', use_container_width=True)
            # if not st.session_state.guardad_registros_pendientes:
            #     if guardar_pospagos:
            #         guardar_conciliacion(st.session_state.ipayouts_data_despues_corte)
            #         st.session_state.guardad_registros_pendientes = True
            #         st.rerun()
            
            cantidad_movimientos = len(st.session_state.ipayouts_data_despues_corte)

            if cantidad_movimientos > 0:
                archivo_nombre = f'Pendiente_Conciliar_{fecha_sel}.xlsx'

                # Convertimos el DataFrame en Excel en memoria
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    st.session_state.ipayouts_data_despues_corte.to_excel(writer, index=False)

                excel_data = excel_buffer.getvalue()

                st.download_button(
                    label=f'DESCARGAR {cantidad_movimientos} MOVIMIENTOS PENDIENTES',
                    data=excel_data,
                    file_name=archivo_nombre,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
            else:
                st.info("No hay movimientos pendientes para descargar.")
                        
        with c2:
            # cantidad_movimientos_conciliados = len(codigos_encontrados)
            # if 'guardar_conciliacion' not in st.session_state:
            #     st.session_state.guardar_conciliacion = False

            # guardar_registros_conciliados = st.button(f'GUARDAR {cantidad_movimientos_conciliados} REGISTROS PAGADOS', use_container_width=True)

            # if not st.session_state.guardar_conciliacion:
            #     if guardar_registros_conciliados:
            #         guardar_registros_pagados(codigos_encontrados)
            #         st.session_state.guardar_conciliacion = True  
            cantidad_movimientos_conciliados = len(codigos_encontrados)

            if cantidad_movimientos_conciliados > 0:
                archivo_nombre_parquet = f'OperacionesPagadas_{fecha_sel}.parquet'

                if 'documento' in codigos_encontrados.columns:
                    codigos_encontrados['documento'] = codigos_encontrados['documento'].astype('string').fillna('')

                #convritmos el dataframe a parquet en memoria
                parquet_buffer = io.BytesIO()
                codigos_encontrados.to_parquet(parquet_buffer, index=False, engine='pyarrow')
                parquet_data = parquet_buffer.getvalue()

                st.download_button(
                    label=f'DESCARGAR {cantidad_movimientos_conciliados} REGISTROS PAGADOS',
                    data=parquet_data,
                    file_name=archivo_nombre_parquet,
                    mime='application/octet-stream',
                    use_container_width=True
                )
            else:
                st.info('No hay registros pagados para descargar')