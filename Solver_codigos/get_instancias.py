import pandas as pd
import streamlit as st
import base64

# descarga de archivos
def xldownload(excel, name):
    data = open(excel, 'rb').read()
    b64 = base64.b64encode(data).decode('UTF-8')
    href = f'<a href="data:file/xls;base64,{b64}" download={name}>Download {name}</a>'
    return href

def get_col_widths(dataframe):
    # First we find the maximum length of the index column
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

def CrearInstancias(xls):
    #st.write('Comenzando función: ')
    df  = pd.read_excel(xls, 'Dotaciones').dropna()
    df2 = pd.read_excel(xls, 'Parametros',index_col=0)
    df3 = pd.read_excel(xls, 'Feriados').dropna()

    puestos = df.index

    dia_inicial = df2.loc['FECHA INICIO','VALOR'].date() #strftime("%Y-%m-%d")
    dia_final = df2.loc['FECHA TERMINO','VALOR'].date()  #strftime("%Y-%m-%d")
    dotacion_externa = 4
    duracion_turno = dict()

    horas_semanales_contrato = int(df2.loc['MAXIMO HORAS POR SEMANA','VALOR'])
    UTM = float(df2.loc['UTM','VALOR'])

    falta_de_trabajadores = int(df2.loc['FALTA DE TRABAJADORES','VALOR'])
    exceso_de_trabajadores = int(df2.loc['EXCESO DE TRABAJADORES','VALOR'])
    descanso = float(df2.loc['DESCANSO','VALOR'])
    mantener_el_mismo_turno = float(df2.loc['MANTENER EL MISMO TURNO','VALOR'])
    maximo_de_minutos_trabajados = float(df2.loc['MAXIMO DE MINUTOS TRABAJADOS','VALOR'])
    maximo_de_turnos_seguidos = float(df2.loc['MAXIMO DE DIAS SEGUIDOS','VALOR'])
    maximo_horas_extra = int(df2.loc['MAXIMO HORAS EXTRA POR SEMANA','VALOR'])


    # if sys.platform == 'linux':
    #     this_folder = os.popen("pwd").read()[:-1]
    # elif platform.system() == 'Windows':
    #     this_folder = os.popen("echo %cd%").read()[:-1]
    #
    #
    # try:
    #     os.mkdir('Instancias')
    # except FileExistsError:
    #     pass  ## qué hacer en caso de que exista
    # out_folder = os.path.join(this_folder,'Instancias')

    horas_requeridas = dict()
    dias_requeridos = dict()

    nro_a_mes = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',7:'Julio',
                    8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}
    # mes_a_nro = {'Enero':1,'Febrero':2,'Marzo':3,'Abril':4,'Mayo':5,'Junio':6,'Julio':7,
    #                 'Agosto':8,'Septiembre':9,'Octubre':10,'Noviembre':11,'Diciembre':12}

    dict_de_excels = {}
    total = len(puestos)*2
    contador_avance = 0
    with st.spinner('Cargando cargos y turnos...'):
        barra_progreso_crearinstancias = st.progress(0)
        for turno_corto in [True, False]:
            if turno_corto:
                tipo_turno = '_turno_corto'
            else:
                tipo_turno = '_turno_largo'

            for index in puestos:
                horas_requeridas[index] = dict()
                dias_requeridos[index]  = dict()
                for i in range(12):
                    horas_requeridas[index][i+1], dias_requeridos[index][i+1] = map(int,df[nro_a_mes[i+1]].loc[index].split("/"))
                    if horas_requeridas[index][i+1] == 24:
                        if not turno_corto:
                            duracion_turno[i+1] = 12
                        else:
                            duracion_turno[i+1] = 8
                    elif horas_requeridas[index][i+1] == 16:
                            duracion_turno[i+1] = 8
                    elif horas_requeridas[index][i+1] == 12:
                        if not turno_corto:
                            duracion_turno[i+1] = 12
                        else:
                            duracion_turno[i+1] = 6
                    else:
                        duracion_turno[i+1] = horas_requeridas[index][i+1]

            #Instancia Parametros
            for index in puestos:
                df_param = pd.DataFrame(columns=['PARAMETROS','VALOR', 'UNIDAD'])
                dotacion = int(df['Dotacion'].loc[index])
                dotacion_turno = int(df['Dotacion turno'].loc[index])
                nombre_instancia = '_'.join(df['Cargo'].loc[index].split())+'_contratados_'+str(dotacion)+tipo_turno+'.xlsx'
                df_param = df_param.append({'PARAMETROS':'FECHA INICIO', 'VALOR':dia_inicial, 'UNIDAD':''}, ignore_index=True) #1r lunes del mes
                df_param = df_param.append({'PARAMETROS':'FECHA TERMINO', 'VALOR':dia_final, 'UNIDAD':''}, ignore_index=True)
                df_param = df_param.append({'PARAMETROS':'MAXIMO HORAS POR SEMANA', 'VALOR':horas_semanales_contrato, 'UNIDAD':'HORAS'}, ignore_index=True)
                df_param = df_param.append({'PARAMETROS':'MAXIMO HORAS EXTRA POR SEMANA', 'VALOR':maximo_horas_extra, 'UNIDAD':'HORAS'}, ignore_index=True)
                df_param = df_param.append({'PARAMETROS':'TRABAJADORES CONTRATADOS EN EL PUESTO', 'VALOR':dotacion, 'UNIDAD':'PERSONAS'}, ignore_index=True)
                df_param = df_param.append({'PARAMETROS':'TRABAJADORES EXTERNOS DISPONIBLES', 'VALOR':dotacion_externa, 'UNIDAD':'PERSONAS'}, ignore_index=True)
                df_param = df_param.append({'PARAMETROS':'TRABAJADORES NECESARIOS EN EL PUESTO', 'VALOR':dotacion_turno, 'UNIDAD':'PERSONAS'}, ignore_index=True)

                #Costos
                factor = 1
                costo_por_hora_contratado = factor*float(df['Sueldo Mes'].loc[index])/(4*horas_semanales_contrato*UTM)
                costo_trabajador_externo = costo_por_hora_contratado * 1.5
                costo_extra_50  = factor*float(df['HE al 50%'].loc[index])/UTM
                costo_extra_100 = factor*float(df['HE al 100%'].loc[index])/UTM
                costo_extra_200 = factor*float(df['HE al 200%*'].loc[index])/UTM
                costo_extra_externo = costo_trabajador_externo * 1.5

                df_costos = pd.DataFrame(columns=['PESO','VALOR','UNIDAD'])
                df_costos = df_costos.append({'PESO':'FALTA DE TRABAJADORES', 'VALOR':falta_de_trabajadores,'UNIDAD':'UTM POR TRABAJADOR (RECOMENDADO 10 VECES EL MAXIMO DE LOS DEMAS)'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'EXCESO DE TRABAJADORES', 'VALOR':exceso_de_trabajadores,'UNIDAD':'UTM POR TRABAJADOR (FICTICIO)'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'DESCANSO', 'VALOR':descanso,'UNIDAD':'UTM POR DIA NO RESPETADO'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'MANTENER EL MISMO TURNO', 'VALOR':mantener_el_mismo_turno,'UNIDAD':'UTM POR SEMANA NO RESPETADA'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'MAXIMO DE MINUTOS TRABAJADOS', 'VALOR':maximo_de_minutos_trabajados,'UNIDAD':'UTM POR SEMANA NO RESPETADA'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'MAXIMO DE DIAS SEGUIDOS', 'VALOR':maximo_de_turnos_seguidos,'UNIDAD':'UTM POR SEMANA NO RESPETADA'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'COSTO TRABAJADOR CONTRATADO', 'VALOR':costo_por_hora_contratado,'UNIDAD':'UTM POR HORA (SALARIO MENSUAL EN UTM/180)'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'COSTO TRABAJADOR EXTERNO', 'VALOR':costo_trabajador_externo,'UNIDAD':'UTM POR HORA'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'COSTO HORA EXTRA 50%', 'VALOR':costo_extra_50,'UNIDAD':'UTM POR HORA'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'COSTO HORA EXTRA 100%', 'VALOR':costo_extra_100,'UNIDAD':'UTM POR HORA'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'COSTO HORA EXTRA 200%', 'VALOR':costo_extra_200,'UNIDAD':'UTM POR HORA'}, ignore_index=True)
                df_costos = df_costos.append({'PESO':'COSTO HORA DOMINGO EXTERNO', 'VALOR':costo_extra_externo,'UNIDAD':'UTM POR HORA'}, ignore_index=True)

                df_req = pd.DataFrame(columns=['MES','HORAS/DIAS','DURACION TURNO (HORAS)'])
                df_req = df_req.append({'MES':'Enero','HORAS/DIAS':str(horas_requeridas[index][1])+'/'+str(dias_requeridos[index][1]),'DURACION TURNO (HORAS)':duracion_turno[1]},ignore_index=True)
                df_req = df_req.append({'MES':'Febrero','HORAS/DIAS':str(horas_requeridas[index][2])+'/'+str(dias_requeridos[index][2]),'DURACION TURNO (HORAS)':duracion_turno[2]},ignore_index=True)
                df_req = df_req.append({'MES':'Marzo','HORAS/DIAS':str(horas_requeridas[index][3])+'/'+str(dias_requeridos[index][3]),'DURACION TURNO (HORAS)':duracion_turno[3]},ignore_index=True)
                df_req = df_req.append({'MES':'Abril','HORAS/DIAS':str(horas_requeridas[index][4])+'/'+str(dias_requeridos[index][4]),'DURACION TURNO (HORAS)':duracion_turno[4]},ignore_index=True)
                df_req = df_req.append({'MES':'Mayo','HORAS/DIAS':str(horas_requeridas[index][5])+'/'+str(dias_requeridos[index][5]),'DURACION TURNO (HORAS)':duracion_turno[5]},ignore_index=True)
                df_req = df_req.append({'MES':'Junio','HORAS/DIAS':str(horas_requeridas[index][6])+'/'+str(dias_requeridos[index][6]),'DURACION TURNO (HORAS)':duracion_turno[6]},ignore_index=True)
                df_req = df_req.append({'MES':'Julio','HORAS/DIAS':str(horas_requeridas[index][7])+'/'+str(dias_requeridos[index][7]),'DURACION TURNO (HORAS)':duracion_turno[7]},ignore_index=True)
                df_req = df_req.append({'MES':'Agosto','HORAS/DIAS':str(horas_requeridas[index][8])+'/'+str(dias_requeridos[index][8]),'DURACION TURNO (HORAS)':duracion_turno[8]},ignore_index=True)
                df_req = df_req.append({'MES':'Septiembre','HORAS/DIAS':str(horas_requeridas[index][9])+'/'+str(dias_requeridos[index][9]),'DURACION TURNO (HORAS)':duracion_turno[9]},ignore_index=True)
                df_req = df_req.append({'MES':'Octubre','HORAS/DIAS':str(horas_requeridas[index][10])+'/'+str(dias_requeridos[index][10]),'DURACION TURNO (HORAS)':duracion_turno[10]},ignore_index=True)
                df_req = df_req.append({'MES':'Noviembre','HORAS/DIAS':str(horas_requeridas[index][11])+'/'+str(dias_requeridos[index][11]),'DURACION TURNO (HORAS)':duracion_turno[11]},ignore_index=True)
                df_req = df_req.append({'MES':'Diciembre','HORAS/DIAS':str(horas_requeridas[index][12])+'/'+str(dias_requeridos[index][12]),'DURACION TURNO (HORAS)':duracion_turno[12]},ignore_index=True)


                # path_instancia = os.path.join(out_folder,nombre_instancia)
                with pd.ExcelWriter(f'Instancias/{nombre_instancia}', datetime_format='yyyy-mm-dd', engine='xlsxwriter') as writer:
                    df_param.to_excel(writer, sheet_name='Parametros', index=False)
                    df_costos.to_excel(writer, sheet_name='Costos', index=False)
                    df_req.to_excel(writer, sheet_name='Requerimientos', index=False)
                    df3.to_excel(writer, sheet_name='Feriados', index=False)
                    for i, width in enumerate(get_col_widths(df_param)):
                        j = max(0,i-1)
                        writer.sheets['Parametros'].set_column(j, j, width+2)

                    for i, width in enumerate(get_col_widths(df_costos)):
                        j = max(0,i-1)
                        writer.sheets['Costos'].set_column(j, j, width+2)

                    for i, width in enumerate(get_col_widths(df_req)):
                        j = max(0,i-1)
                        writer.sheets['Requerimientos'].set_column(j, j, width+2)

                    for i, width in enumerate(get_col_widths(df3)):
                        j = max(0,i-1)
                        writer.sheets['Feriados'].set_column(j, j, width+2)

                #print(df_param)
                #print(df_costos)
                #print(nombre_instancia)
                contador_avance += 1
                barra_progreso_crearinstancias.progress(contador_avance/total)
                #dict_de_excels = dict(dict_de_excels, **{nombre_instancia:xldownload(writer, nombre_instancia)}) # ([xldownload(writer, nombre_instancia), nombre_instancia])
                dict_de_excels = dict(dict_de_excels, **{nombre_instancia: writer})  # ([xldownload(writer, nombre_instancia), nombre_instancia])
        barra_progreso_crearinstancias.empty()
    return dict_de_excels