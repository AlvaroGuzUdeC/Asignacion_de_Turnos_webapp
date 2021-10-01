import streamlit as st
import pandas as pd
import base64

from Solver_codigos.get_instancias import CrearInstancias
from Solver_codigos.ExtendedWeek import solution_by_week
from Solver_codigos.ComputeSolutionCost import *


# Parámetros/variables/funciones que se usarán:

st.set_page_config(layout="wide")
metodos = ['Resolver Dotaciones', 'Comparación de resultados']  #Nombre de los métodos con los que contará la app
paginas_navegacion = ['Home', 'Instrucciones'] + metodos + ['About Us'] #Nombre de las posibles páginas de navegación

def xldownload(excel_writer, name):
    ## Función para descargar la clase "writer" como excel.
    data = open(excel_writer, 'rb').read()
    b64 = base64.b64encode(data).decode('UTF-8')
    href = f'<a href="data:file/xls;base64,{b64}" download="{name}.xlsx">Descargar {name}</a>'
    return href
# ------------------------------------------------------------------------------------------------------------------


st.image('Imagenes_web/logo.png')

# Título

_, center, _ = st.columns((1, 2, 1))
center.markdown('# Asignación de Turnos.')

# Sidebar

with st.sidebar:
    st.image('Imagenes_web/logo.png')
    st.text('')
    st.text('')
    pag_navegacion_actual = st.radio('Navegar', paginas_navegacion)

## Página de bienvenida.
if pag_navegacion_actual == paginas_navegacion[0]:
    '''
    ## Bienvenido/a.

    Se encuentra en la interfaz para la ejecución del programa de __asignación de turnos__. Esta aplicación fue desarrollada
    en Surpoint Analytics SpA con el fin de facilitar el uso del _solver_ encargado de resolver el problema asignación
    de turnos.

    A la izquierda está la barra de navegación. En _Instrucciones_ se encuentran los tutoriales para el uso de las
    diferentes herramientas de esta interfaz.
    '''

## Página de instrucciones:
if pag_navegacion_actual == paginas_navegacion[1]:
    st.header('Instrucciones.')
    '''
    Primero, frente a cualquier problema con la interfaz o querer reiniciar el proceso, se recomiendo limpiar reinicar la página.
    Esto borrará el caché y volverá a correr la interfaz desde cero.
    '''
    '''
    Indique qué método desea revisar:
    '''
    metodo_instrucciones = st.selectbox('Seleccione método.', metodos)

    if metodo_instrucciones == metodos[0]:  # Instrucciones para el primer método
        '''
        Este método consiste de dos etapas. Primero recibe un archvo de dotaciones en formato excel y,
        a partir de este archivo, crea una instancia (archivo en formato excel) por cada ocupación y turno señalado
        en el archivo de dotaciones. Luego, por cada una de estas instancias, crea un archivo _solución_, que contiene
        la calendarización propuesta por el __solver__.

        __Entradas__: Sólo hay una entrada: el excel de dotaciones que se sube al presionar el _uploader_.
        '''

        st.file_uploader('Ejemplo de uploader:')

        '''
        Este excel de dotaciones debe tener tres hojas: Dotaciones, Parametros y Feriados. Cada una debe estar estructurada
        de la siguiente forma:

        __Dotaciones__:
        '''

        st.image('Imagenes_web/DotacionesEjemplo.jpg')

        '''
        __Parametros__:
        '''

        st.image('Imagenes_web/DotacionesEjemplo2.jpg')

        '''
        __Feriados__:
        '''

        st.image('Imagenes_web/DotacionesEjemplo3.jpg')

        '''
        Puede revisar el siguiente excel de dotaciones como ejemplo:
        '''

        st.markdown(xldownload('elementos_web/Dotaciones.xlsx', 'Dotaciones ejemplo'), unsafe_allow_html=True)
        '''
        Una vez cargado este archivo dotaciones, deberá presionar el botón "Cargar instancias" y, si subió bien el archivo
        y su contenido está correcto, se cargarán todas las instancias de los distintos cargos. Podrá comprobar
        que todos los cargos se cargaron si en el __Multiselect__ se despliegan como opciones.         
        '''
        '''
        _Ejemplo de __Multiselec__ con algunos cargos:_
        '''
        st.multiselect('Multiselect', ['Supervisor_Caldera_contratados_1_turno_corto.xlsx', 'Operador_Pozos_contratados_2_turno_largo.xlsx'])
        '''
        En este _Multiselect_ deberá escoger los cargos que quieres resolver. Una vez haya elegido todos los cargos a resolver,
        debe presionar el botón "Ejecutar Solver". La barra de progreso que se desplegará corresponde a la cantidad de cargos
        resueltos en un instante dado. Tener en cuenta que el solver tarda aproximadamente 800 segundos (13 minutos y medio aprox)
        por cargo.
        '''
        '''
        Finalmente, cuando se termine el proceso aparecerá un cuadro como el siguiente:
        '''
        st.success('Se ha completado la calendarización')
        '''
        Y además aparecerá una espacio expandible como el siguiente:
        '''
        with st.expander('Haga click aquí'):
            st.write('Lista de descargas.')
        '''
        Desde el cual se desplegará una lista con todos los cargos y turnos resueltos. Basta hacer click para descargar
        los excels resultantes. Notar que existirán dos archivos por cargo, uno llamado "Resultado" y otro llamado "Instancia".
        El "Resultado" corresponde a la solución asociada que calculó el solver, y el archivo "Instancia" es el archivo que el
        solver ocupara para resolver dicho cargo. Existen estas dos opciones porque serán útiles para el siguiente método
        que permitirá crear una calendarización manual y comparar los costos totales asociados. Si sólo desea conocer la calendarización,
        puede ignorar el archivo "Instancia".
        
        El botón 'Reset All' permite reinicar el proceso desde cero.  
        '''

    elif metodo_instrucciones == metodos[1]:  # Instrucciones para el segundo metodo
        '''
        Este método consiste de una única etapa. Recibe dos archivos asocaidos a un cargo: 'Instancia' y 'Resultado modificado'.
        El archivo de 'Instancia' corresponde al archivo que puede descargarse al ejecutar el solver (primer método disponible),
        y el archivo "Resultado modificado" corresponde a una modificación del archivo que se puede descargar al 
        ejecutar el solver (primer método disponible). La idea es que usted puede modificar el archivo "Resultado" entregado
        por el solver (sólo la primera hoja del excel) para así crear una calendarización del cargo de forma manual y comparar
        los costos de esta calendarización manual con los costos de la calendarización propuesta por el solver.

        __Entradas__: Hay dos entradas: un excel de "Instancia" y otro de "Resultado modificado". Ambos deben subirse a los 
        __uploaders__ corrspondientes.
        '''

        '''
        Una vez subidos los archivos, deberá presionar el botón "Computar Costos". Si todo está en orden, se desplegarán distintos
        graficos que mostrarán las comparaciones de los distintos tipos de costos. Notar que "OG" se refiere a la calendarización
        original (propuesta por el solver) y, "MD", a la modificación manual que fue entregada.
        
        Puede probar, a modo de ejemplo, con los siguientes archivos:      
        '''
        st.markdown(xldownload('elementos_web/Instancia_Operador_Pozos_contratados_2_turno_corto.xlsx', 'Instancia ejemplo'), unsafe_allow_html=True)
        st.markdown(xldownload('elementos_web/Resultado_Operador_Pozos_contratados_2_turno_corto.xlsx', 'Resultado modificado ejemplo'), unsafe_allow_html=True)

if pag_navegacion_actual == metodos[0]:
    st.header('Crear soluciones desde archivo de Dotaciones.')

    #Uploader del archivo excel de dotaciones
    if 'dotaciones_uploader_key' not in st.session_state:
        st.session_state.dotaciones_uploader_key = 0
    excel_metodo1_uploader = st.file_uploader('Subir excel de dotaciones.', type="xlsx", key=st.session_state.dotaciones_uploader_key)

    #Obtener los excel instancias con get_instancias
    boton_cargar_instancias = st.button('Cargar instancias')
    if boton_cargar_instancias:
        if excel_metodo1_uploader is None:
            st.error('No se ha subido ningún archivo. Por favor suba el excel al uploadre e intente de nuevo.')
        else:
            excel_metodo1 = pd.ExcelFile(excel_metodo1_uploader)
            st.session_state.dict_excels = CrearInstancias(excel_metodo1)

    #Comprobamos si se crearon las instancias y luego se eligen las que se quiere resolver/calanderizar
    if 'dict_excels' in st.session_state:
        excels_names = list(st.session_state.dict_excels.keys())
        container_multiselect = st.container()
        check_all = st.checkbox('Seleccionar todos')
        if check_all:
            instancias_seleccionadas = container_multiselect.multiselect('Multiselect', excels_names, excels_names)
        else:
            instancias_seleccionadas = container_multiselect.multiselect('Multiselect', excels_names)
        #Una vez escogidos los cargos a resolver, se ocupa ExtendedWeek.solution_by_week() para obtener los excels resultados
        if st.button('Ejecutar solver'):
            st.session_state.dict_excels_resultados = {}
            with st.spinner('Cargando calendarizaciones...'):
                total_soluciones = len(instancias_seleccionadas)
                barra_progreso_soluciones = st.progress(0)
                contador_soluciones = 0
                for xls_name in instancias_seleccionadas:
                    xls_writer = st.session_state.dict_excels[xls_name]
                    xls = pd.ExcelFile(xls_writer)
                    resultado_name, resultado_writer = solution_by_week(xls)
                    st.session_state.dict_excels_resultados = dict(st.session_state.dict_excels_resultados, **{resultado_name: resultado_writer})
                    contador_soluciones += 1
                    barra_progreso_soluciones.progress(contador_soluciones/total_soluciones)

    # Si se cargaron las soluciones, se despliegan los links de descarga
    if 'dict_excels_resultados' in st.session_state:
        st.success('Se ha completado la calendarización de todo los cargos')
        resultados_expander = st.expander("Click aquí para descargar resultados", expanded=True)
        with resultados_expander:
            st.markdown('***')
            for name in st.session_state.dict_excels_resultados.keys():
                display_name = name[10:].replace('_', ' ')
                st.markdown(f'__{display_name}__')

                #Subiendo archivo de resultado
                st.markdown(xldownload(st.session_state.dict_excels_resultados[name], name), unsafe_allow_html=True)
                # Subiendo archivo de instancia
                instancia_name = name[10:]
                st.markdown(xldownload(st.session_state.dict_excels[instancia_name+'.xlsx'], 'Instancia_'+instancia_name), unsafe_allow_html=True)

                st.markdown('***')




    if st.button('Reiniciar todo'):
        st.session_state.dotaciones_uploader_key += 1
        try:
            del (st.session_state.dict_excels)
        except:
            pass
        try:
            del (st.session_state.dict_excels_resultados)
        except:
            pass
        st.legacy_caching.clear_cache()
        st.experimental_rerun()

if pag_navegacion_actual == metodos[1]:
    st.header('Comparaciones de costes en base a un archivo modificado manualmente')

    # Uploader del archivo excel de instancia y de resultado manual
    if 'instancia_uploader_key' not in st.session_state:
        st.session_state.instancia_uploader_key = 0
    instancia_metodo2_uploader = st.file_uploader('Subir excel de isntancia.', type="xlsx",
                                              key=st.session_state.instancia_uploader_key)
    if 'resultado_uploader_key' not in st.session_state:
        st.session_state.resultado_uploader_key = 0
    resultado_metodo2_uploader = st.file_uploader('Subir excel de resultados modificado.', type="xlsx",
                                              key=st.session_state.resultado_uploader_key)

    #Se ejecuta la función ComputeSolutionCost
    if st.button('Computar Costos'):
        if (instancia_metodo2_uploader is None) or (resultado_metodo2_uploader is None):
            st.error('Revise que ha subido ambos archivos.')
        else:
            #Costos asociados al resultado original
            costos1, costos2, costos3 = get_costos_from_resultadoxls(resultado_metodo2_uploader)

            #Costos asociados al resultado modificado
            costo_cuant_mod, cost_interno_mod, costo_total_mod = ComputeSolutionCosts(instancia_metodo2_uploader, resultado_metodo2_uploader)

            #Gráficos
            col1, col2 = st.columns(2)
            with col1:
                st.subheader('Costos Cuantificables')
                chart_costo_cuant = pd.DataFrame([[costos1, 0], [0, costo_cuant_mod]], columns=['Original', 'Modificado'], index=['OG', 'MD'])
                st.bar_chart(chart_costo_cuant)
            with col2:
                st.subheader('Costos Internos')
                chart_costo_int = pd.DataFrame([[costos2, 0], [0, cost_interno_mod]], columns=['Original', 'Modificado'] ,index=['OG', 'MD'])
                st.bar_chart(chart_costo_int)
            st.subheader('Costos Totales')
            chart_costo_tot = pd.DataFrame([[costos3, 0], [0, costo_total_mod]], columns=['Original', 'Modificado'] ,index=['OG', 'MD'])
            st.bar_chart(chart_costo_tot)

            st.markdown('__ La diferencia entre el costo total original y el modificado es de: __')
            total_display = round(costos3-costo_total_mod, 3)
            st.write('OG-MD = '+str(total_display))
            if total_display > 0:
                porcentaje = round((1 - (costo_total_mod)/costos3)*100,3)
                st.markdown(f'__ Es decir, el costo modificado es {porcentaje}% menos que el original.__')
            elif total_display-1 < 0:
                porcentaje = -round((1 - (costo_total_mod+100) / costos3) * 100, 3)
                st.markdown(f'__ Es decir, el costo modificado es {porcentaje}% más que el original.__')


    pass

if pag_navegacion_actual == 'About Us':
    '''
    __SURPOINT__, Es una empresa global de Tecnologías de la Información, con sede matriz en la ciudad de Concepción,
     Región del Bio Bío, Chile. Contamos con más de 12 años de experiencia y más de 100 proyectos realizados.
    '''
    '''
    Nuestro portafolio cuenta con 3 áreas de especialidad y servicios enfocados en:

 Proyectos de Desarrollo de Software, aplicaciones empresariales, aplicaciones móviles, Business Intelligence, desarrollo en SAP | Abap.
 Proyectos de Tecnología, Fibra Óptica, Soporte en Conectividad e Infraestructura, Cableado, Redes Wifi, Certificación, Fusión FO.
 Proyectos Analytics, Redes Neuronales, Predicción de Demanda, Modelamiento de Procesos.
Entregamos un servicio personalizado para cada uno de los requerimientos y necesidades de nuestros clientes, focalizado en mantener estándares de excelencia en la industria de servicios en tecnologías de la información.
    '''
    '''
    __SURPOINT ANALYTICS SE RESERVA TODOS LOS DERECHOS DE USO, PRIVACIDAD Y DIVULGACIÓN DE ESTA APLICACIÓN WEB.__ 
    '''