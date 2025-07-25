import pandas as pd
import streamlit as st
from datetime import datetime
import re
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
import io
from notion_client import Client



# imagen = Image.open('kashio_cover.jpeg')   
# st.image(imagen, use_container_witdh=True)     

#=========================================
# Primera parte. Subida archivo METABASE
#=========================================

st.title('Conciliacion PAYOUTS dia anterior')
st.write('Herramienta para la conciliacion de los pagos del dia anterior')

#primero cargamos el archivo de los payouts del metabase 

payouts_metabase = st.file_uploader('Sube el archivo de payouts del metabase', type=['xlsx'])

if payouts_metabase is not None:
    payouts_metabase_df = pd.read_excel(payouts_metabase)
    #el tipo de datos para ope_psp
    #payouts_metabase_df['ope_psp'] = payouts_metabase_df['ope_psp'].astype('Int64').astype(str)

    payouts_metabase_df['ope_psp'] = (
    pd.to_numeric(payouts_metabase_df['ope_psp'], errors='coerce')  # convierte lo numérico, pone NaN al resto
    .astype('Int64')  # conserva los NaN
    .astype(str)  # lo pasas a string si lo necesitas para merge
)

    #creamos una columna con la fehca de proceso con solo la fecha
    payouts_metabase_df['fecha_proceso'] = pd.to_datetime(payouts_metabase_df['fecha pagado / rechazado']).dt.date
    #convertimos la columna de fecha_proceso a tipo fecha
    payouts_metabase_df['fecha_proceso'] = pd.to_datetime(payouts_metabase_df['fecha_proceso'])
    
    # #filtramos por fecha de hoy
    # hoy = pd.Timestamp.today().normalize()
    # ayer = pd.Timestamp.today().normalize() - pd.Timedelta(days=1)  # Ayer
    # payouts_metabase_df = payouts_metabase_df[payouts_metabase_df['fecha_proceso'] == ayer]

    #Extraemos la hora de creacion
    payouts_metabase_df['hora'] = payouts_metabase_df['fecha proceso'].dt.hour

    #Extraemos la fecha de proceso
    payouts_metabase_df['date'] = payouts_metabase_df['fecha proceso'].dt.date


    #para uso de nombres de archivos
    fecha = pd.to_datetime(payouts_metabase_df['fecha_proceso'].unique()[0]).strftime("%Y%m%d")

    #filtramos el estado de la operacion a pagado
    payouts_metabase_df = payouts_metabase_df[payouts_metabase_df['estado'] == 'Pagado']

    #filtramos por el tipo de moneda
    payouts_metabase_df = payouts_metabase_df[payouts_metabase_df['moneda'] == 'PEN']

    #filtramos todos los BANCOs menos scotiabank 
    payouts_metabase_df = payouts_metabase_df[payouts_metabase_df['name'] != '(Scotiabank)- Scotiabank ']    

    #creamos una tabla pivot con los montos de cada banco
    pivot_payouts = payouts_metabase_df.groupby(['fecha_proceso','name'])['monto total'].sum().reset_index()
    group_hour = payouts_metabase_df.groupby(['name', 'ope_psp']).agg({'monto total':'sum'}).reset_index()
    #group_hour['ope_psp'] = group_hour['ope_psp'].astype(str)
    columns_name = {
        'ope_psp':'Operación - Número'
    }
    
    group_hour = group_hour.rename(columns=columns_name)

    #st.dataframe(payouts_metabase_df, use_container_width=True)
    st.dataframe(pivot_payouts, use_container_width=True)


    #=========================================
    #Definicion de funciones 
    #=========================================

    #=========================================
    # Guardar conciliacion de la tabla consoolidada de los bancos y metabase
    #=========================================


    def guardar_conciliacion(payouts_metabase_df, df_final, nombre_primera_hoja="Payouts_Metabase", nombre_segunda_hoja="Operaciones Bancos"):
        site_url = "https://kashioinc.sharepoint.com/sites/Intranet2021"
        username = "dussand.hurtado@kashio.net"
        password = "Silvana1505$"
        #password = "Silvana1505$"
        status_placeholder = st.empty()

        #with st.spinner():

        status_placeholder.info('Conectando a Sharepoint...')

        try:
            # Conectamos al sitio
            ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
            
            # Ruta relativa a la carpeta en SharePoint (CORREGIDA: debe empezar con /sites/...)
            folder_url = "/sites/Intranet2021/Shared Documents/Operaciones/PAYOUT/PAYOUTS VARIOS/Conciliaciones Payout"

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
        archivo_nombre = f'Conciliacion_{fecha}.xlsx' #in case doesn't work, delete this

        
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
                payouts_metabase_df.to_excel(writer, sheet_name=nombre_primera_hoja, index=False)
                #st.info(f"📊 Hoja 1: '{nombre_primera_hoja}' - {len(payouts_metabase_df)} registros")
                
                # Guardar el segundo DataFrame (df_final)
                df_final.to_excel(writer, sheet_name=nombre_segunda_hoja, index=False)
                #st.info(f"📊 Hoja 2: '{nombre_segunda_hoja}' - {len(df_final)} registros")
                    
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

    #=========================================
    # Enviar registro de diferencias de los bancos a Notion
    #=========================================

    def registros_notion(conciliacion_payouts):
        notion_token = "ntn_Yk820926168aR213bRsLF9pqG3t88PU0YMqhUazW4ap2qE"
        database_id =  "21d030ee56d880de8976ce1b1fe6b8fc"

        notion = Client(auth=notion_token)

        status_placeholder = st.empty()
        progress_bar = st.progress(0)

        for idx, (_,rows) in enumerate(conciliacion_payouts.iterrows()):
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
                        },

                        'Estado': {
                            'select': {
                                'name': str(rows.get('Estado', ''))
                            }
                        }                                      
                    }
                )

                progress = (idx + 1) / len(conciliacion_payouts)

                progress_bar.progress(min(progress,1.0))
                status_placeholder.success(f'Registro {idx + 1} guardado correctamente')

            except Exception as e:
                status_placeholder.error(f'Registro {idx + 1} falló: {e}')

    #=========================================
    # Definicion de funciones de lectura de bancos
    #=========================================                


    #=========================================
    # BCP
    #=========================================  
    #definimos funciones para cada banco
    def procesar_bcp(archivo):
        bcp_eecc = pd.read_excel(archivo, skiprows=4)
        #cambiamos el tipo de dato del numero de operacion 
        bcp_eecc['Operación - Número'] = bcp_eecc['Operación - Número'].astype(str)
        #filtramos la columna Referencia 2 por los que contienen PAYOUT
        bcp_eecc = bcp_eecc[bcp_eecc['Referencia2'].str.contains('PAYOUT', case=False, na=False)]

        bcp_eecc['Hora'] = pd.to_datetime(bcp_eecc['Operación - Hora'], format = '%H:%M:%S', errors='coerce').dt.hour

        #eliminaremos columnas innecesarias 

        suma_monto_por_hora = bcp_eecc.groupby('Hora')['Monto'].sum().reset_index()
        
         #2. Obtenemos una fila representativa por hora, solo de pagos (montos negativos)
        pagos_negativos = bcp_eecc[bcp_eecc['Monto'] < 0]

        fila_negativa_por_hora  = pagos_negativos.sort_values('Hora').groupby('Hora').first().reset_index()

        bcp_consolidado = pd.merge(fila_negativa_por_hora, suma_monto_por_hora, on='Hora')

        bcp_consolidado = bcp_consolidado.drop(columns=['Fecha valuta','Descripción operación' ,'Saldo', 'Sucursal - agencia'
                                                        , 'Usuario', 'UTC', 'Hora', 'Operación - Hora'
                                                        , 'Monto_x'
                                                    ])
        bcp_consolidado = bcp_consolidado.rename(columns={'Monto_y':'Monto'})
        #creamos una columna con el nombre del banco
        bcp_consolidado['name'] = '(BCP) - Banco de Crédito del Perú'
        #total = bcp_eecc['Monto'].sum() * -1
        return bcp_consolidado
    
    
    #=========================================
    # INTERBANK
    #=========================================  
    
    def procesar_interbank(archivo):
        ibk_eecc = pd.read_excel(archivo, skiprows=13) #leemos el excel 
        # #eliminamos la primera columna
        ibk_eecc = ibk_eecc.drop(columns=['Unnamed: 0'])
        # # #eliminamos la fila sin valores
        # ibk_eecc = ibk_eecc.dropna(how='all')

        # # #cambiamos el nombre de las columnas
        columns_name = {
            'Fecha de Proc.': 'Fecha',
            'Cargos':'Monto',
            'Detalle': 'Referencia2',
            'Cod. de Operación': 'Operación - Número'
        }

        ibk_eecc = ibk_eecc.rename(columns=columns_name)

        # # #filtramos la columna 'Nombre de la solicitud' por los valores que contienen 
        ibk_eecc = ibk_eecc[ibk_eecc['Referencia2'].str.contains(r'\bPA(Y|YOU|YOUT|YO)?\b', case=False, na=False)]

        #cambiamos el numero de operacion a sin 0 inicial
        ibk_eecc['Operación - Número'] = ibk_eecc['Operación - Número'].astype(int).astype(str)

        # #limpiamos la columna 'Monto soles' y lo convertimo a float 
        # ibk_eecc['Monto'] = (
        #   ibk_eecc['Monto'].astype(str) #convertimos en string primero
        #   .str.replace('S/', '', regex=False) #reemplazamos S/ por nada para borrarlo
        #   .str.replace(',','', regex=False) #tambien la coma 
        #   .str.strip() #eliminamos espacios que existan
        #   .astype(float)  #y lo convertimos a decimal para poder sumarlo
        # )
        # # #total = ibk_eecc['Monto soles'].sum() #sumamos la columna monto soles

        # #creamos una columna con el nombre del banco
        ibk_eecc['name'] = '(Interbank) - Banco International del Perú'
        
        # # #eliminaremos columnas innecesarias 
        ibk_eecc = ibk_eecc.drop(columns=['Fecha de Op.', 'Movimiento'
                                           ,'Canal', 'Cod. de Ubicación', 'Abonos', 'Saldo contable'
                                           ])
        
        return ibk_eecc
    
    #=========================================
    # BBVA - OTROS BANCOS Y MANUALES
    #=========================================  

    def procesar_bbva_otros(archivo):
        bancos_bbva = pd.read_excel(archivo, skiprows=10)

        # Renombrar columnas
        columns_name = {
            'F. Operación': 'Fecha',
            'Concepto': 'Referencia2',
            'Importe': 'Monto',
            'Nº. Doc.':'Operación - Número'
        }
        bancos_bbva = bancos_bbva.rename(columns=columns_name)

        #filtrar los op del metabase
        valores_metabase = payouts_metabase_df[payouts_metabase_df['name'] == '(BBVA) - BBVA Continental ']['ope_psp'].dropna().astype(str).unique()

        #filtramos el metabase con los valores_metabase que son los numeros de op unicos del metabase
        df_bbva = bancos_bbva[
            bancos_bbva['Operación - Número'].astype(str).apply(
                lambda x: any(valor in x for valor in valores_metabase)
            )
        ].copy()

        df_bbva['Operación - Número'] = df_bbva['Operación - Número'].astype(int).astype(str)
        df_bbva['name'] = '(BBVA) - BBVA Continental '

        # DataFrame con filas que contienen "BXI"
        df_otros = bancos_bbva[
            bancos_bbva['Referencia2'].astype(str).str.contains('BXI', case=False, na=False)
        ].copy()

        #extraemos el numero de operacion de la columna Referencia2 y lo reemplazmos en la columna Operación - Número
        # df_otros['Operación - Número'] = df_otros['Referencia2'].astype(str).apply(
        #     lambda x: str(int(re.search(r'(\d{5,})$', x).group(1 if re.search(r'(\d{5,})$', x) else None)
        # )))

        df_otros['Operación - Número'] = df_otros['Referencia2'].astype(str).apply(
        lambda x: str(int(re.search(r'(\d{5,})$', x).group(1))) if re.search(r'(\d{5,})$', x) else None
        )

        #df_otros = df_otros[df_otros['Operación - Número'].notna()]
        
        df_otros['name'] = 'Otros bancos'

        # ==========
        # manuales
        # # ==========

        # df_manuales = bancos_bbva[bancos_bbva['Referencia2'].astype(str).str.contains('BXI CT', case=False, na=False)].copy()

        # df_manuales['name'] = 'Manuales'
        # df_manuales['Operación - Número'] = None

        bancos_bbva_filtrado = pd.concat([df_bbva, df_otros], ignore_index=True)

        #eliminamos columnas innecesaarias
        bancos_bbva_filtrado = bancos_bbva_filtrado.drop(
            columns=['F. Valor', 'Código', 'Oficina']
        )

        return bancos_bbva_filtrado

    #=========================================
    # DICCIONAARIO DE FUNCIONES POR BANCO
    #=========================================  

    #creamos el diccionario de funciones de cada banco
    procesadores_banck = {
        'bcp': procesar_bcp,
        'ibk': procesar_interbank,
        'bbva':procesar_bbva_otros
    }

    #=========================================
    # Lectura de los estados de cuenta de los bancos
    #=========================================  

    #creamos la seccion para subir el estado de cuenta del banco seleccionado
    estado_cuenta = st.file_uploader(f'Subir estados de cuenta', type=['xlsx', 'xls'], accept_multiple_files=True
                                     )
    
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

    if df_consolidados:
        df_final = pd.concat(df_consolidados, ignore_index=True)
        st.subheader("📊 Datos consolidados de todos los bancos")
        df_final_group = df_final.groupby(['name', 'Operación - Número']).agg({'Monto':'sum'}).reset_index() #informaciones de los bancos
        group_hour = payouts_metabase_df.groupby(['name', 'ope_psp']).agg({'monto total':'sum', 'hora':lambda x: x.unique()[0]}).reset_index() #informacion del metabase
        group_hour = group_hour.rename(columns={'ope_psp':'Operación - Número'})
        
        st.dataframe(df_final)

  
        merge_op = pd.merge(df_final_group, group_hour, on = 'Operación - Número', how='outer')
        merge_op['Diferencias'] = round((merge_op['monto total'] + merge_op['Monto']), 2)
        merge_op = merge_op[merge_op['Diferencias'] != 0]
        #st.dataframe(merge_op)

        #mostramos un pivot con los montos de los bancos 
        bancos_montos = df_final.groupby('name')['Monto'].sum().reset_index() #pivot de los datos consolidados de los bancos 
        bancos_montos['Monto'] = bancos_montos['Monto'].abs()
        #st.dataframe(bancos_montos, use_container_width=True)


   #=========================================
    # Registro de diferencias entre los bancos metabase y estados de cuenta
    #========================================= 


        st.subheader('Conciliacion de los montos de todos los bancos')
        st.write(''' En esta seccion podremos encontrar si hay diferencias
                  entre los montos de los bancos de los estados de cuenta y el metabase del core
                  de Kashio, para poder analizar los cortes de payouts regulares.''')
        #uniremos los df con los resultados finales
        conciliacion_payouts = pd.merge(pivot_payouts, bancos_montos, on='name', how='outer')
        #mostramos las diferencias
        conciliacion_payouts['Diferencia'] = round(conciliacion_payouts['monto total'] - conciliacion_payouts['Monto'], 2)

        #creamos una columna que nos arroja que banco tienen diferencias para pasar a analizaarlo
        conciliacion_payouts['Estado'] = conciliacion_payouts['Diferencia'].apply(lambda x: 'Conciliado' if x == 0 else 'Diferencias')
        
    
        columns_diferences = {
            'fecha_proceso': 'FechaTexto',
            'name':'BANCO',
            'monto total':'Monto Kashio',
            'Monto':'Monto Banco',
            'Diferencia':'Diferencia',
            'Estado':'Estado'
        }

        #sales

        conciliacion_payouts = conciliacion_payouts.rename(columns=columns_diferences)

        conciliacion_payouts['FechaTexto'] = conciliacion_payouts['FechaTexto'].fillna(conciliacion_payouts['FechaTexto'].values[0])

        st.dataframe(conciliacion_payouts, use_container_width=True)

        #hoy_str = hoy.strftime('%d/%m/%Y')
        #creamos una columna esstado por defecto a todo el df
        #payouts_metabase_df['Estado'] = f'Conci. {hoy_str}'
        payouts_metabase_df['Estado'] = f'Conciliacion_{fecha}' #en caso no funcione borrar

    #=========================================
    # Vista de diferencias encontradas
    #========================================= 

        # Inicializa el estado de guardado si no existe
        if 'guardado_metabase' not in st.session_state:
            st.session_state.guardado_metabase = False

        if 'guardar_record_dif' not in st.session_state:
            st.session_state.guardar_record_dif = False

        #mostramos un aviso si hay diferencias
        if 'Diferencias' in conciliacion_payouts['Estado'].values:
            st.warning('Se detectaron diferencias en la conciliación')
         
            if 'Banco metabes' not in merge_op.columns:
                columns_name = {
                    'name_x': 'Banco estados de cuenta',
                    'Operación - Número': 'Numero operacion banco',
                    'Monto': 'Monto bancos',
                    'name_y': 'Banco metabase',
                    'monto total': 'Monto metabase',

                }
                merge_op = merge_op.rename(columns=columns_name)
                
                #Mostrar solo detalle de diferencias para los bancos que tienen diferencias
                #creamos una columna con el banco final
                merge_op['Banco final'] = merge_op['Banco metabase'].combine_first(merge_op['Banco estados de cuenta'])
                # 1. Filtrar los bancos con diferencia mayor a 0
                bancos_con_diferencias = conciliacion_payouts[ (conciliacion_payouts['Diferencia'] > 0) | (conciliacion_payouts['Diferencia'] < 0) ]['BANCO'].unique()

                # 2. Filtrar merge_op solo para esos bancos
                merge_op_filtrado = merge_op[merge_op['Banco final'].isin(bancos_con_diferencias)]

                with st.expander('Detalle de diferencias'):
                    st.dataframe(merge_op_filtrado.iloc[:, :7], use_container_width=True)

                diferencias_ = payouts_metabase_df['ope_psp'].isin(merge_op['Numero operacion banco'])
                #payouts_metabase_df.loc[diferencias_, 'Estado'] = f'Conci. {hoy_str} - Diferencias' 
                payouts_metabase_df.loc[diferencias_, 'Estado'] = f'Conciliacion_{fecha} - Diferencias' #in case doesn't work, delete this
                metabase_filter_dife = payouts_metabase_df[diferencias_].copy()
                #st.dataframe(metabase_filter_dife)
                #boton para guardar  

                c1, c2 = st.columns(2)      
                with c1:          
                    if not st.session_state.guardado_metabase:
                        if st.button('Guardar conciliación en SharePoint', use_container_width=True):
                            guardar_conciliacion(payouts_metabase_df, df_final)
                            st.session_state.guardado_metabase = True
                            #st.rerun()
                with c2:
                    if not st.session_state.guardar_record_dif:
                        if st.button('Registrar diferencias en Notion', use_container_width=True):
                            registros_notion(conciliacion_payouts)
                            st.session_state.guardar_record_dif = True
                            st.rerun()

                # st.divider()
                # st.title('Busqueda de diferencias')

                # archivo_diferencias = st.file_uploader('Sube el lote de la hora identificada con diferencias.', type=['xls', 'xlsx'])

                # if archivo_diferencias is not None: #cambiar las variables una vez completado el codigo
                #     diferencias_ibk = pd.read_excel(archivo_diferencias, skiprows=24)
                #     diferencias_ibk = diferencias_ibk.drop(columns=[ #Eliminamos las columnas que no nos sirve
                #     'Unnamed: 0', 'Beneficiario',
                #     'Unnamed: 3', 'Unnamed: 4',
                #     'Unnamed: 5', 'Tipo de abono',
                #     'Unnamed: 8', 'Unnamed: 9',
                #     'Cuenta', 'Unnamed: 12',
                #     'Unnamed: 13', 'Unnamed: 14',
                #     'Documento', 'Unnamed: 17',
                #     'Unnamed: 18', 'Vencimiento',
                #     'Unnamed: 21', 'Unnamed: 22',
                #     'Unnamed: 23', 'Monto',
                #     'Unnamed: 26', 'Unnamed: 27',
                #     'Estado', 'Unnamed: 30',
                #     'Unnamed: 31', 'Unnamed: 32',
                #     'Observación'
                # ])
                #     diferencias_ibk = diferencias_ibk.dropna(how='all') #eliminamos las filas con valores nulos
                    

                #     columns_diferencias_name = {
                #         'Unnamed: 1': 'Beneficiario',
                #         'Unnamed: 6': 'Tipo de abono',
                #         'Unnamed: 10':'Cuenta',
                #         'Unnamed: 15':'Documento',
                #         'Unnamed: 19':'Vencimiento',
                #         'Unnamed: 24':'Monto',
                #         'Unnamed: 28':'Estado',

                #     }
                #     diferencias_ibk = diferencias_ibk.rename(columns=columns_diferencias_name) #cambiamos el nombre de las columnas
                    

                #    creamos un codigo unico con cada nombre y monto para encontrar la diferencia
                #     diferencias_ibk['Beneficiario'] = diferencias_ibk['Beneficiario']\
                #         .astype(str)\
                #         .str.replace(r'DNI.*', '', regex=True)\
                #      .apply(lambda x: x.replace(' ', ''))
                    

                    
                #     diferencias_ibk['Monto'] = diferencias_ibk['Monto'].astype(str)
                #     diferencias_ibk['Codigo'] = (
                #         diferencias_ibk['Beneficiario'].astype(str).str[:2] +
                #         diferencias_ibk['Beneficiario'].astype(str).str[-3:] 
                #     )
                
                #        diferencias_ibk['Monto'].astype(float).astype(int).astype(str).str[:2]

                #         diferencias_ibk['Codigo'] = diferencias_ibk['Codigo'].str.replace(' ', '').str.lower()
                #     st.write('excel diferencias ibk')
                #     diferencias_ibk[['Beneficiario', 'Codigo']]
                #     unicos_dif = diferencias_ibk['Codigo'].unique()
                #     st.write(unicos_dif, use_container_witdh = True)


                #     creamos el codigo para el df de diferencias de toda la hora identificada 
                #     metabase_filter_dife['cliente'] = metabase_filter_dife['cliente'].astype(str).str.upper()
                #     metabase_filter_dife['cliente'] = metabase_filter_dife['cliente'].astype(str).apply(lambda x: x.replace(' ', ''))
                #     metabase_filter_dife['Codigo'] = metabase_filter_dife['cliente'].astype(str).str[:2] + metabase_filter_dife['cliente'].astype(str).str[-2:]
                #     + " " + metabase_filter_dife['monto total'].astype(str).str[:2]
                #     unicos_filt = metabase_filter_dife['Codigo'].unique()
                #     st.write(unicos_filt, use_container_witdh = True)
                #     st.write('excel metabase ibk')
                #     metabase_filter_dife[['cliente', 'Codigo']]

                #     cruzamos para encontrar la diferencia
                #     diferencia_encontrada = pd.merge(diferencias_ibk, metabase_filter_dife, on='Codigo', how='left')
                #     diferencia_encontrada
                    
                #     comparacion = diferencias_ibk['Codigo'].reset_index(drop=True) == metabase_filter_dife['Codigo'].reset_index(drop=True)
                #     comparacion

        else:
            st.success('No se encontraron diferencias en la conciliación')

            with st.container():

                k1, k2 = st.columns(2)
                with k1:
                    if not st.session_state.guardado_metabase:
                        if st.button('Guardar conciliación en SharePoint', use_container_width=True):
                            #payouts_metabase_df['Estado'] = f'Conci. {hoy_str}'
                            payouts_metabase_df['Estado'] = f'Conciliacion_{fecha}' #en caso no funcione borrar
                            guardar_conciliacion(payouts_metabase_df, df_final)
                with k2:
                    if not st.session_state.guardar_record_dif:
                        if st.button('Registrar conciliacion en Notion', use_container_width=True):
                            registros_notion(conciliacion_payouts)
                            st.session_state.guardar_record_dif = True
                            # st.rerun()

                # c1, c2 = st.columns(2)      
                # with c1:          
                #     if not st.session_state.guardado_metabase:
                #         if st.button('Guardar conciliación en SharePoint', use_container_width=True):
                #             guardar_conciliacion(payouts_metabase_df, df_final)
                #             st.session_state.guardado_metabase = True
                #             st.rerun()
                # with c2:
                #     if not st.session_state.guardar_record_dif:
                #         if st.button('Registrar diferencias en Notion', use_container_width=True):
                #             registros_notion(conciliacion_payouts)
                #             st.session_state.guardar_record_dif = True
                #             st.rerun()
