import streamlit as st
import pandas as pd

st.title('Conciliacion Instant - Payouts')


#=========================================
# Primera parte. Subida y lectura de archivo METABASE
#=========================================

#Subimos el excel de metabase 
file_uploader_metabase = st.file_uploader('Arrastra el archivo de metabase aquí: ', type=['xlsx'])

if file_uploader_metabase is not None:
    ipayouts_metabase_df = pd.read_excel(file_uploader_metabase) # cargamos el excel

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

    alcance_bancos = [
        '(BCP) - Banco de Crédito del Perú',
        'Yape',
        '(BBVA) - BBVA Continental '
    ]

    ipayouts_metabase_df = ipayouts_metabase_df[ipayouts_metabase_df['banco'].isin(alcance_bancos)] #filtramos los bancos que vamos a usar

    ipayouts_metabase_df = ipayouts_metabase_df[ipayouts_metabase_df['estado'] == 'Pagado']


    montos_ipayouts = ipayouts_metabase_df.groupby('banco')['monto'].sum().reset_index()
    st.dataframe(montos_ipayouts, use_container_width=True)


#=========================================
# Segunda parte. Definicion de funciones para lecturas de eecc
#=========================================

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

        
        columnas_name = {'Fecha': 'fecha',
            'Descripción operación': 'descripcion_operacion',
            'Monto':'importe',
            'Operación - Número':'numero_operacion',
            'Operación - Hora':'hora_operacion'}
    

        estado_cuenta_df.rename(columns=columnas_name, inplace=True)
        # Asegurar que 'Operación - Número' es string
        estado_cuenta_df['numero_operacion'] = estado_cuenta_df['numero_operacion'].astype(str)

        # Crear columna 'codigo_operacion' con lógica condicional
        estado_cuenta_df['codigo_operacion'] = estado_cuenta_df.apply(
            lambda x: x['descripcion_operacion'][-11:] if str(x['descripcion_operacion'])[:3] == 'YPP' else x['numero_operacion'],
            axis=1
        )
        estado_cuenta_df['banco'] = estado_cuenta_df.apply(
            lambda x: 'Yape' if str(x['descripcion_operacion'])[:3] == 'YPP' else '(BCP) - Banco de Crédito del Perú',
            axis = 1
        )

        return estado_cuenta_df
    
    import pandas as pd

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

        columnas_name = {'F. Operación': 'fecha',
            'Nº. Doc.': 'descripcion_operacion',
            'Importe':'importe'}
        

        estado_cuenta_df.rename(columns=columnas_name, inplace=True)

        # Convertir a string para asegurar consistencia
        estado_cuenta_df['descripcion_operacion'] = estado_cuenta_df['descripcion_operacion'].astype(str)

        # Extraer código operación desde 'Concepto'
        estado_cuenta_df['codigo_operacion'] = estado_cuenta_df['Concepto'].astype(str).str[-20:]

        estado_cuenta_df['banco'] = '(BBVA) - Banco Continental '
        return estado_cuenta_df


    #creamos el diccionario de funciones de cada banco
    procesadores_banck = {
        'bcp': procesar_bcp,
        'bbva': procesar_bbva
    }


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

    if df_consolidados:
        df_final = pd.concat(df_consolidados, ignore_index=True)
        df_final

        def extraer_codigo(row):
            banco = row['banco']
            concepto = str(row['numero de operacion'])

            if banco == '(BCP) - Banco de Crédito del Perú':
                return concepto[18:27]
            elif banco == '(BBVA) - BBVA Continental ':
                return concepto[:20]
            elif banco == 'Yape':
                return concepto[-11:]
            else:
                return None
        
        ipayouts_metabase_df['codigo_operacion'] = ipayouts_metabase_df.apply(extraer_codigo, axis=1)
        ipayouts_metabase_df

        codigo_bancos_set = set(df_final['codigo_operacion'])

        ipayouts_metabase_df['resultado_busqueda'] =  ipayouts_metabase_df['codigo_operacion'].apply(
            lambda x: x if x in codigo_bancos_set else 'No encontrado'
        )

        conciliar_mañana = ipayouts_metabase_df[ipayouts_metabase_df['resultado_busqueda'] == 'No encontrado']
        ipayouts_metabase_df
        st.write(len(conciliar_mañana))