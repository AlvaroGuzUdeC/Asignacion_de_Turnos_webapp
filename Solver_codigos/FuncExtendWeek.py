class Parametros:
    #Clase con los parámetros del problema. Algunos leidos desde archivo otros modificables aquí
    def __init__(self, default=True, MinTotalMinutes = None):
        from datetime import date
        self.Horizon = 0               #Tiempo de modelación semanal
        self.NumberOfWeeks  = 0        #Semanas totales a simular
        self.LengthOfShifts = 0        #Duración diaria del turno
        self.TotalHoursPerDay = 0      #Requerimiento de horas diarias
        self.ShiftsPerDay = 0          #Número de turnos posibles en un dia
        self.minConsecutiveShifts = 1  #Numero minimo de turnos por semana
        self.maxConsecutiveShifts = 0  #Numero maximo de turnos por semana
        self.minConsecutiveDaysOff = 0 #Numero minimo de turnos por semana
        self.maxWeekends = 2           #Numero máximo de fines de semana a trabajar (inutil pegando semana a semana)
        self.weightForUnder = 0        #Costo de falta de dotación
        self.weightForOver = 0         #Costo de dotación más de la necesaria
        self.MaxTotalMinutes = 0       #Minutos máximos a trabajar
        self.MinTotalMinutes = 0       #Minutos minimos a trabajar (sólo trabajadores contratados)
        self.TotalContract = 0         #Número de trabajadores contratados
        self.TotalExternal = 0         #Número de trabajadores extras
        self.Requirement = 0           #Número de trabajadores necesarios
        self.ContratadosObligados = True  #Usar o no la condición de al menos un contratado en cada turno
        self.MaximoHorasExtra = 0      #Máximo de horas extras
        self.workerWeight = dict()     #Costo de cada trabajador (interno)
        self.domingosPorMes = dict()        #Máximo de domingos a trabajar al mes
        self.GlobaldomingosPorMes = dict()  #
        self.TurnosDeNoche = dict()
        self.GlobalTurnosDeNoche = dict()
        self.costoContratado = 0            # Factor costo trabajador contratado
        self.costoExterno = 0               # Factor costo trabajador externo
        self.CostoExtra50 = 0           # Costo horas extra
        self.CostoExtra100 = 0          # Costo horas extra
        self.CostoExtra200 = 0          # Costo horas extra
        self.CostoExtraExterno = 0
        self.prohibitNextWeight = 0
        self.maxShiftsWeight = 0
        self.minTotalMinutesWeight = 0
        self.maxTotalMinutesWeight = 0
        self.minConsecutiveShiftsWeight = 0
        self.maxConsecutiveShiftsWeight = 0
        self.minConsecutiveDaysOffWeight = 0
        self.weekendsWeight = 0
        self.externalWeight = 0    
        self.LargoDelTurno = 0
        self.ConteoHorasExtra = dict()
        self.dia_inicio = date.today()         #Fecha de inicio de la modelación (date)
        self.dia_fin = date.today()         #Fecha de inicio de la modelación (date)
        self.horas_requeridas = dict()
        self.dias_requeridos  = dict()
        self.duracion_turnos = dict()
        self.feriados = dict()
        self.IdShifts1Turno  = ['T. Día']    #Id de cada turno
        self.IdShifts2Turnos = ['T. Día','T. Noche']    #Id de cada turno
        self.IdShifts3Turnos = ['T. Mañana','T. Tarde','T. Noche']
        self.IdShifts = list()
        self.IdStaffContratado = ["Trabajador "+chr(x) for x in range(ord('A'), ord('Z') + 1)]+["Trabajador "+str(x) for x in range(1, 100 + 1)] #Id de los trabajadores
        self.IdStaffEventual = ["Eventual "+chr(x) for x in range(ord('A'), ord('Z') + 1)]+["Eventual "+str(x) for x in range(1, 100 + 1)] #Id de los trabajadores
        self.IdStaff = list()
        #Rellenar defaults
        if default == True:
            self.fill_default()
        if MinTotalMinutes != None:
            self.MinTotalMinutes = MinTotalMinutes

    def fill_default(self):
        self.MinTotalMinutes = 35*60
        self.prohibitNextWeight = 1
        self.minConsecutiveShiftsWeight = 1
        self.minConsecutiveDaysOffWeight = 1
        self.weekendsWeight = 1
        self.externalWeight = 1    

def lunes_de_la_semana(fecha):
    import datetime
    date = datetime.datetime(fecha.year,fecha.month,fecha.day).date()
    # for dia in range(1,8):
    #     if date.weekday() == 0:
    if fecha.day-date.weekday()>=0:
        return datetime.datetime(fecha.year,fecha.month,fecha.day-date.weekday())
    else:
        return datetime.datetime(fecha.year,fecha.month,7-date.weekday()+fecha.day)

def funcionBloqueo(index,item):
    #Función necesaria para bloquear dias en caso de 3 turnos por dia
    a = index%2
    if item < 2:
        b = 1
    else:
        b = 0
    return (a+b)%2

def CheckIndexColumnExists(df,nombre):
    if nombre == 'Parametros':
        indexes=['FECHA INICIO',
                    'FECHA TERMINO',
                    # 'DURACION DEL TURNO',
                    #'REQUERIMIENTO DE TIEMPO POR DIA',
                    #'REQUERIMIENTO DE DIAS POR SEMANA',
                    'MAXIMO HORAS POR SEMANA',
                    'MAXIMO HORAS EXTRA POR SEMANA',
                    'TRABAJADORES CONTRATADOS EN EL PUESTO',
                    'TRABAJADORES EXTERNOS DISPONIBLES',
                    'TRABAJADORES NECESARIOS EN EL PUESTO',
                    #'USAR CONDICION CONTRATADO OBLIGATORIO'
                    ]
        columnas=['VALOR']
    elif nombre == 'Costos':
        indexes=['FALTA DE TRABAJADORES',
                    'EXCESO DE TRABAJADORES',
                    'DESCANSO',
                    'MANTENER EL MISMO TURNO',
                    # 'MINIMO DE MINUTOS TRABAJADOS',
                    'MAXIMO DE MINUTOS TRABAJADOS',
                    'MAXIMO DE DIAS SEGUIDOS',
                    # 'MINIMO DE TURNOS DESCANSADOS',
                    'COSTO TRABAJADOR CONTRATADO',
                    'COSTO TRABAJADOR EXTERNO',
                    'COSTO HORA EXTRA 50%',
                    'COSTO HORA EXTRA 100%',
                    'COSTO HORA EXTRA 200%',
                    'COSTO HORA DOMINGO EXTERNO']
        columnas=['VALOR']
    elif nombre == 'Requerimientos':
        indexes=['Enero',
                    'Febrero',
                    'Marzo',
                    'Abril',
                    'Mayo',
                    'Junio',
                    'Julio',
                    'Agosto',
                    'Septiembre',
                    'Octubre',
                    'Noviembre',
                    'Diciembre']
        columnas=['HORAS/DIAS']
    else:
        raise Exception("CheckIndexColumnExists error")

    for index in indexes:
        if index not in df.index:
            print(index)
            print(df.index)
            raise Exception(nombre+" no encontrado en Instancia.xlsx")

    for columna in columnas:
        if columna not in df.columns:
            print(columna)
            print(df.columns)
            raise Exception("Nombre de columna mal escrito en "+nombre+" de Instancia.xlsx")

def UpdateShiftRequirement(prm,mes):
    import sys
    prm.Horizon = min(7,int(prm.dias_requeridos[mes]))
    prm.TotalHoursPerDay = int(prm.horas_requeridas[mes])
    prm.LargoDelTurno = int(prm.duracion_turnos[mes])

    if prm.TotalHoursPerDay == 24 and prm.LargoDelTurno == 12:
        prm.LengthOfShifts = 12*60
        prm.ShiftsPerDay = 2
        prm.maxConsecutiveShifts = min(prm.Horizon,7)
        prm.IdShifts = prm.IdShifts2Turnos
    elif prm.TotalHoursPerDay == 24 and prm.LargoDelTurno == 8:
        prm.LengthOfShifts = 8*60
        prm.ShiftsPerDay = 3
        prm.maxConsecutiveShifts = min(prm.Horizon,7)
        prm.IdShifts = prm.IdShifts3Turnos
    elif prm.TotalHoursPerDay == 20:
        prm.LengthOfShifts = 10*60
        prm.ShiftsPerDay = 2
        prm.maxConsecutiveShifts = min(prm.Horizon,7)
        prm.IdShifts = prm.IdShifts2Turnos
    elif prm.TotalHoursPerDay == 16:
        prm.LengthOfShifts = 8*60
        prm.ShiftsPerDay = 2
        prm.maxConsecutiveShifts = min(prm.Horizon,7)
        prm.IdShifts = prm.IdShifts2Turnos
    elif prm.TotalHoursPerDay == 12:
        prm.LengthOfShifts = 12*60
        prm.ShiftsPerDay = 1
        prm.maxConsecutiveShifts = min(prm.Horizon,7)
        prm.IdShifts = prm.IdShifts1Turno
    elif prm.TotalHoursPerDay == 12 and prm.LargoDelTurno == 6:
        prm.LengthOfShifts = 6*60
        prm.ShiftsPerDay = 2
        prm.maxConsecutiveShifts = min(prm.Horizon,7)
        prm.IdShifts = prm.IdShifts2Turnos
    elif prm.TotalHoursPerDay == 10:
        prm.LengthOfShifts = 10*60
        prm.ShiftsPerDay = 1
        prm.maxConsecutiveShifts = min(prm.Horizon,7)
        prm.IdShifts = prm.IdShifts1Turno
    else:
        print(prm.TotalHoursPerDay, prm.LengthOfShifts)
        sys.exit("Error: ShiftRequirement, Horas de trabajo por dia")

    return prm

def ReadFromExcel(Instancia,prm, DEBUG=False, fecha = None):
    #Lectura del archivo Excel de entrada y construcciốn del formato texto leido por el solver
    from Solver_codigos.roster_parser import ParseRoster
    import pandas as pd

    #Lectura de hojas de archivo excel
    xls = pd.ExcelFile(Instancia)
    df  = pd.read_excel(xls, 'Parametros',index_col=0)
    df2 = pd.read_excel(xls, 'Costos',index_col=0)
    df3 = pd.read_excel(xls, 'Requerimientos',index_col=0)
    df4 = pd.read_excel(xls, 'Feriados')

    df.index = df.index.str.strip()
    df.columns = df.columns.str.strip()
    df2.index = df2.index.str.strip()
    df2.columns = df2.columns.str.strip()
    df3.index = df3.index.str.strip()
    df3.columns = df3.columns.str.strip()

    CheckIndexColumnExists(df,'Parametros')
    CheckIndexColumnExists(df2,'Costos')
    CheckIndexColumnExists(df3,'Requerimientos')
    #Lectura de parámetros

    import datetime
    if type(df.loc['FECHA INICIO','VALOR']) is type(datetime.datetime.now()):
        prm.dia_inicio = lunes_de_la_semana(df.loc['FECHA INICIO','VALOR'])
    else:
        raise Exception('FECHA INICIO mal ingresado en Parametros Instancia.xlsx')

    nro_a_mes = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',7:'Julio',
            8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}
    try:
        for i in range(12):
            prm.horas_requeridas[i+1], prm.dias_requeridos[i+1] =  map(int,df3.loc[nro_a_mes[i+1],'HORAS/DIAS'].split("/"))
    except:
        raise Exception('HORAS/DIAS mal ingresado en Requerimientos Instancia.xlsx')

    try:
        for i in range(12):
            prm.duracion_turnos[i+1] =  int(df3.loc[nro_a_mes[i+1],'DURACION TURNO (HORAS)'])
    except:
        raise Exception('DURACION TURNO (HORAS) mal ingresado en Requerimientos Instancia.xlsx')

    # try:
    # except:
    #     raise Exception('REQUERIMIENTO DE DIAS POR SEMANA mal ingresado en Parametros Instancia.xlsx')

    # try:
    # except:
    #     raise Exception('REQUERIMIENTO DE TIEMPO POR DIA mal ingresado en Parametros Instancia.xlsx')

    if type(df.loc['FECHA TERMINO','VALOR']) is type(datetime.datetime.now()):
        prm.dia_fin = df.loc['FECHA TERMINO','VALOR']
    else:
        raise Exception('FECHA TERMINO mal ingresado en Parametros Instancia.xlsx')

    prm.NumberOfWeeks = int(round((prm.dia_fin-prm.dia_inicio).days/7))

    if prm.NumberOfWeeks <= 0:
        raise Exception('Problemas entre la fecha de inicio y término. Revisar Parametros Instancia.xlsx') 

    if prm.NumberOfWeeks > 100:
        raise Exception('Maximo 100 semanas de calendarización. Revisar fechas en Parametros Instancia.xlsx') 
    # try:
    #     prm.TurnoLargo = str(df.loc['DURACION DEL TURNO','VALOR'])
    # except:
    #     raise Exception('DURACION DEL TURNO mal ingresado en Parametros Instancia.xlsx')


    #Construcción de parámetros compatibles con horas de turno y requerimientos
    #Duración de turnos y cantidad
    #Maximo de turnos consecutivos
    if fecha is None:
        prm = UpdateShiftRequirement(prm,prm.dia_inicio.month)
    else:
        prm = UpdateShiftRequirement(prm,fecha.month)

    #Minutos por semana
    try:
        prm.MaxTotalMinutes = int(df.loc['MAXIMO HORAS POR SEMANA','VALOR'])*60
    except:
        raise Exception('MAXIMO HORAS POR SEMANA mal ingresado en Parametros Instancia.xlsx')

    # Maximo horas extra por semana
    try:
        prm.MaximoHorasExtra = int(df.loc['MAXIMO HORAS EXTRA POR SEMANA','VALOR'])
    except:
        raise Exception('MAXIMO HORAS EXTRA POR SEMANA mal ingresado en Parametros Instancia.xlsx')

    #Trabajadores totales
    try:
        prm.TotalContract = int(df.loc['TRABAJADORES CONTRATADOS EN EL PUESTO','VALOR'])
    except:
        raise Exception('TRABAJADORES CONTRATADOS EN EL PUESTO mal ingresado en Parametros Instancia.xlsx')

    try:
        prm.TotalExternal = min(int(df.loc['TRABAJADORES EXTERNOS DISPONIBLES','VALOR']),prm.ShiftsPerDay)
    except:
        raise Exception('TRABAJADORES EXTERNOS DISPONIBLES mal ingresado en Parametros Instancia.xlsx')

    #Trabajadores necesarios
    try:
        prm.Requirement = int(df.loc['TRABAJADORES NECESARIOS EN EL PUESTO','VALOR'])
    except:
        raise Exception('TRABAJADORES NECESARIOS EN EL PUESTO mal ingresado en Parametros Instancia.xlsx')

    #Usar condicion contratados obligatorios
    # try:
    #     if int(df.loc['USAR CONDICION CONTRATADO OBLIGATORIO','VALOR']) == 1:
    #         prm.ContratadosObligados = True
    # except:
    #     pass


    #Costo distintos trabajadores
    try:
        prm.costoContratado = float(df2.loc['COSTO TRABAJADOR CONTRATADO','VALOR'])
    except:
        raise Exception('COSTO TRABAJADOR CONTRATADO mal ingresado en Costos Instancia.xlsx')

    try:
        prm.costoExterno = float(df2.loc['COSTO TRABAJADOR EXTERNO','VALOR'])
    except:
        raise Exception('COSTO TRABAJADOR EXTERNO mal ingresado en Costos Instancia.xlsx')

    for i in range(0,prm.TotalContract):
        prm.workerWeight[i] = prm.costoContratado
        prm.IdStaff.append(prm.IdStaffContratado[i])
    for i in range(prm.TotalContract,prm.TotalContract+prm.TotalExternal):
        prm.workerWeight[i] = prm.costoExterno
        prm.IdStaff.append(prm.IdStaffEventual[i-prm.TotalContract])

    #Costo de falta o exceso de trabajadores
    try:
        prm.weightForUnder = int(df2.loc['FALTA DE TRABAJADORES','VALOR'])
    except:
        raise Exception('FALTA DE TRABAJADORES mal ingresado en Costos Instancia.xlsx')

    try:
        prm.weightForOver = int(df2.loc['EXCESO DE TRABAJADORES','VALOR'])
    except:
        raise Exception('EXCESO DE TRABAJADORES mal ingresado en Costos Instancia.xlsx')

    #Costo horas extra
    try:
        prm.CostoExtra50 = float(df2.loc['COSTO HORA EXTRA 50%','VALOR'])
    except:
        raise Exception('COSTO HORA EXTRA 50% mal ingresado en Costos Instancia.xlsx')

    try:
        prm.CostoExtra100 = float(df2.loc['COSTO HORA EXTRA 100%','VALOR'])
    except:
        raise Exception('COSTO HORA EXTRA 100% mal ingresado en Costos Instancia.xlsx')

    try:
        prm.CostoExtra200 = float(df2.loc['COSTO HORA EXTRA 200%','VALOR'])
    except:
        raise Exception('COSTO HORA EXTRA 200% mal ingresado en Costos Instancia.xlsx')

    try:
        prm.CostoExtraExterno = float(df2.loc['COSTO HORA DOMINGO EXTERNO','VALOR'])
    except:
        raise Exception('COSTO HORA DOMINGO EXTERNO mal ingresado en Costos Instancia.xlsx')

    try:
        prm.daysOffWeight = float(df2.loc['DESCANSO','VALOR'])
    except:
        raise Exception('DESCANSO mal ingresado en Costos Instancia.xlsx')

    try:
        prm.maxShiftsWeight = float(df2.loc['MANTENER EL MISMO TURNO','VALOR'])
    except:
        raise Exception('MANTENER EL MISMO TURNO mal ingresado en Costos Instancia.xlsx')

    try:
        prm.maxTotalMinutesWeight = float(df2.loc['MAXIMO DE MINUTOS TRABAJADOS','VALOR'])
    except:
        raise Exception('MAXIMO DE MINUTOS TRABAJADOS mal ingresado en Costos Instancia.xlsx')

    try:
        prm.maxConsecutiveShiftsWeight = float(df2.loc['MAXIMO DE DIAS SEGUIDOS','VALOR'])
    except:
        raise Exception('MAXIMO DE TURNOS SEGUIDOS mal ingresado en Costos Instancia.xlsx')

    
    for index,row in df4.iterrows():
        prm.feriados[index] = row['DIAS'] #.date()

    #Build Text to parse
    tmp_instance = ''
    tmp_instance += "# This is a comment. Comments start with #\n"
    tmp_instance += "SECTION_HORIZON\n"
    tmp_instance += "# All instances start on a Monday\n"
    tmp_instance += "# The horizon length in days:\n"
    tmp_instance += str(prm.Horizon)+"\n\n"

    tmp_instance += "SECTION_SHIFTS\n"
    tmp_instance += "# ShiftID, Length in mins, Shifts which cannot follow this shift | separated\n"
    for index in range(0,prm.ShiftsPerDay):
        tmp_instance += str(prm.IdShifts[index])
        tmp_instance += ","
        tmp_instance += str(prm.LengthOfShifts)
        tmp_instance += ","
        # tmp_instance += str(prm.IdShifts[(index+1)%prm.ShiftsPerDay]) #VOY AQUI
        tmp_instance += "\n"
    tmp_instance += "\n"

    tmp_instance += "SECTION_STAFF\n"
    tmp_instance += "# ID, MaxShifts, MaxTotalMinutes, MinTotalMinutes, MaxConsecutiveShifts, MinConsecutiveShifts, MinConsecutiveDaysOff, MaxWeekends, Costo, Contratado, CostoExtra50, CostoExtra100, CostoExtra200, CostoExtraExterno\n"
    for index in range(0,prm.TotalContract+prm.TotalExternal):
        tmp_instance += prm.IdStaff[index]
        prm.domingosPorMes[prm.IdStaff[index]] = 0
        prm.GlobaldomingosPorMes[prm.IdStaff[index]] = 0
        prm.TurnosDeNoche[prm.IdStaff[index]] = 0
        prm.GlobalTurnosDeNoche[prm.IdStaff[index]] = 0
        tmp_instance += ","
        if index < prm.TotalContract:
            for item in range(0,prm.ShiftsPerDay):
                    # tmp_instance += IdShifts[item]+"="+str(Horizon)
                #HECHO A MANO, FUNCIONA CON HASTA 3 TURNOS
                if prm.ShiftsPerDay not in [1,2,3]:
                    sys.exit("Not implemented:ReadFromExcel")
                if prm.ShiftsPerDay == 1:
                    tmp_instance += prm.IdShifts[item]+"="+str(prm.Horizon) 
                if prm.ShiftsPerDay == 2:
                    tmp_instance += prm.IdShifts[item]+"="+str(prm.Horizon)
                    # tmp_instance += prm.IdShifts[item]+"="+str(prm.Horizon*((index+item)%2))
                if prm.ShiftsPerDay == 3:
                    tmp_instance += prm.IdShifts[item]+"="+str(prm.Horizon)
                    # tmp_instance += prm.IdShifts[item]+"="+str(prm.Horizon*(funcionBloqueo(index,item)))
                if item != prm.ShiftsPerDay - 1:
                    tmp_instance += "|"
        else:
            for item in range(0,prm.ShiftsPerDay):
                tmp_instance += prm.IdShifts[item]+"="+str(prm.Horizon)
                if item != prm.ShiftsPerDay - 1:
                   tmp_instance += "|"
        #Max minutes por semana
        tmp_instance += ","
        if index < prm.TotalContract:
            tmp_instance += str(prm.MaxTotalMinutes)
        else:
            tmp_instance += str(7*24*60)

        #Min minutes por semana
        tmp_instance += ","
        if index < prm.TotalContract:
            tmp_instance += str(prm.MinTotalMinutes)
        else:
            tmp_instance += str(0)
        tmp_instance += ","
        #MaxConsecutiveShifts
        if index < prm.TotalContract:
            # tmp_instance += str(prm.maxConsecutiveShifts)
            tmp_instance += str(prm.Horizon)
        else:
            tmp_instance += str(prm.Horizon)
        tmp_instance += ","
        #minConsecutiveShifts
        if index < prm.TotalContract:
            tmp_instance += str(prm.minConsecutiveShifts)
        else:
            tmp_instance += str(0)
        tmp_instance += ","
        #minConsecutiveDaysOff
        tmp_instance += str(prm.minConsecutiveDaysOff)
        tmp_instance += ","
        #maxWeekends
        tmp_instance += str(prm.maxWeekends)
        tmp_instance += ","
        #workerWeight
        tmp_instance += str(prm.workerWeight[index])
        tmp_instance += ","
        #Contratado o No
        if index < prm.TotalContract:
            tmp_instance += str(1)
        else:
            tmp_instance += str(0)
        tmp_instance += ","
        #CostoExtra50
        if index < prm.TotalContract:
            tmp_instance += str(prm.CostoExtra50)
        else:
            tmp_instance += str(0)
        tmp_instance += ","
        #CostoExtra100
        if index < prm.TotalContract:
            tmp_instance += str(prm.CostoExtra100)
        else:
            tmp_instance += str(0)
        tmp_instance += ","
        #CostoExtra200
        if index < prm.TotalContract:
            tmp_instance += str(prm.CostoExtra200)
        else:
            tmp_instance += str(0)
        tmp_instance += ","
        #CostoExtraExterno
        if index < prm.TotalContract:
            tmp_instance += str(0)
        else:
            tmp_instance += str(prm.CostoExtraExterno)
        tmp_instance += "\n"
    tmp_instance += "\n"

    tmp_instance += "SECTION_DAYS_OFF\n"
    tmp_instance += "# EmployeeID, DayIndexes (start at zero)\n"
    tmp_instance += "\n"

    tmp_instance += "SECTION_SHIFT_ON_REQUESTS\n"
    tmp_instance += "# EmployeeID, Day, ShiftID, Weight\n"
    tmp_instance += "\n"

    tmp_instance += "SECTION_SHIFT_OFF_REQUESTS\n"
    tmp_instance += "# EmployeeID, Day, ShiftID, Weight\n"
    tmp_instance += "\n"

    tmp_instance += "SECTION_COVER\n"
    tmp_instance += "# Day, ShiftID, Requirement, Weight for under, Weight for over\n"
    #Requerimientos por turnos y por dias
    for dia in range(0,prm.Horizon):
        for index in range(0,prm.ShiftsPerDay):
            tmp_instance += str(dia)
            tmp_instance += ","
            tmp_instance += prm.IdShifts[index]
            tmp_instance += ","
            tmp_instance += str(prm.Requirement)
            tmp_instance += ","
            tmp_instance += str(prm.weightForUnder)
            tmp_instance += ","
            tmp_instance += str(prm.weightForOver)
            tmp_instance += "\n"

    tmp_instance += "\n"

    tmp_instance += "SECTION_COSTOS\n"
    tmp_instance += "# daysOffWeight, prohibitNextWeight, maxShiftsWeight, minTotalMinutesWeight, maxTotalMinutesWeight, minConsecutiveShiftsWeight, maxConsecutiveShiftsWeight, minConsecutiveDaysOffWeight, weekendsWeight, externalWeight\n"
    tmp_instance += str(prm.daysOffWeight)+","      #daysOffWeight
    tmp_instance += str(prm.prohibitNextWeight)+","      #prohibitNextWeight
    tmp_instance += str(prm.maxShiftsWeight)+","      #maxShiftsWeight
    tmp_instance += str(prm.minTotalMinutesWeight)+","      #minTotalMinutesWeight
    tmp_instance += str(prm.maxTotalMinutesWeight)+","      #maxTotalMinutesWeight
    tmp_instance += str(prm.minConsecutiveShiftsWeight)+","      #minConsecutiveShiftsWeight
    tmp_instance += str(prm.maxConsecutiveShiftsWeight)+","      #maxConsecutiveShiftsWeight
    tmp_instance += str(prm.minConsecutiveDaysOffWeight)+","      #minConsecutiveDaysOffWeight
    tmp_instance += str(prm.weekendsWeight)+","      #weekendsWeight
    tmp_instance += str(prm.externalWeight)          #externalWeight

    tmp_instance += "\n"

    if DEBUG:
        print(tmp_instance)
    
    if fecha is None:
        dia_inicio = prm.dia_inicio
    else:
        dia_inicio = fecha

    return ParseRoster(contents_from_excel=tmp_instance,MinutosPorSemana=prm.MaxTotalMinutes,
                        contratados_obligatorios=prm.ContratadosObligados,MaximoHorasExtra=prm.MaximoHorasExtra,
                        feriados=prm.feriados,dia_inicio=dia_inicio), prm

def UpdateConditions(problem,solution, debug = None, prm = None, fecha = None, week = None):
    import sys
    import numpy as np
    import Solver_codigos.instance
    import datetime

    #Borra restricciones antiguas
    # for staffId, schedule in solution.schedule.items():
    for staffId in problem.staff.keys():
        problem.staff[staffId].daysOff = set()
        problem.staff[staffId].maxShifts = dict()

    #No aplicar restricciones a internos
    Externals = [prm.IdStaff[i] for i in range(prm.TotalContract,prm.TotalContract+prm.TotalExternal)]
    items = [item for item in problem.shifts.keys()]

    #Cuenta domingos trabajados
    for staffId, schedule in solution.schedule.items():
        if staffId in Externals:
            continue
        turno_final = list(schedule)[-1]
        solution.DomingoTrabajado[staffId] = 0
        if turno_final.strip() != '' and problem.horizon == 7:
            solution.DomingoTrabajado[staffId] = 1
            prm.domingosPorMes[staffId] += 1
            prm.GlobaldomingosPorMes[staffId] += 1

    #Cuenta turnos de noche
    tmpTrabajoDeNoche =  dict()
    for staffId, schedule in solution.schedule.items():
        if staffId in Externals:
            continue
        shiftsTaken = dict()
        tmpTrabajoDeNoche[staffId] = 0
        for idx, shift in enumerate(solution.schedule[staffId]):
            shiftsTaken[shift] = shiftsTaken.get(shift, 0) + 1
        for shift, count in shiftsTaken.items():
            # print(shift,count,len(items))
            if shift == items[-1] and count > 1 and len(items)>1:
                prm.TurnosDeNoche[staffId] += 1
                prm.GlobalTurnosDeNoche[staffId] += 1
                tmpTrabajoDeNoche[staffId] = 1
        if tmpTrabajoDeNoche[staffId] == 0:
            prm.TurnosDeNoche[staffId] = 0
        # print(staffId,schedule)
        # print(staffId,shiftsTaken)

    #Bloquea domingos
    # for staffId, schedule in solution.schedule.items():
    for staffId in problem.staff.keys():
        if prm.domingosPorMes[staffId] > 1 and staffId not in Externals:
            problem.staff[staffId].daysOff.add(6)
            if debug is not None:
                "bloquearle domingo a {}\n".format(staffId)

    #Bloquea turnos de noche
    for id in problem.staff.keys():
        pb = problem.staff[id]
        if id in Externals:
            continue

        print(id,prm.TurnosDeNoche[id])
        if prm.TurnosDeNoche[id] > 1:
            pb.maxShifts[items[-1]] = 0

    #Resetea el proximo mes
    prox_domingo = fecha+datetime.timedelta(days=6)
    print(prox_domingo)
    if int(prox_domingo.day/7)==0 or prox_domingo.day==7:
        # for staffId, schedule in solution.schedule.items():
        for staffId in problem.staff.keys():
            if staffId in Externals:
                continue
            if tmpTrabajoDeNoche[staffId] != 1:
                problem.staff[staffId].maxShifts = dict() #Desbloquea si no trabajó la ultima semana de noche
            problem.staff[staffId].daysOff = set() #Resetea el bloqueo del domingo
            prm.domingosPorMes[staffId] = 0
            prm.TurnosDeNoche[staffId] = 0

    #DEBUG
    if debug is not None:
        for id in problem.staff.keys():
            pb = problem.staff[id]
            print("despues",pb.maxShifts)
    #DEBUG

    #Bloquea primer dia del nuevo periodo si trabajo un turno de noche
    if int(prox_domingo.day/7)==0 or prox_domingo.day==7:
        for staffId, schedule in solution.schedule.items():
            turno_final = list(schedule)[-1]
            if turno_final.strip() == items[-1] and staffId not in Externals and problem.horizon == 7 and len(items)>1:
                problem.staff[staffId].daysOff.add(0)

    # for staffId, schedule in solution.schedule.items():
    for staffId in problem.staff.keys():
        print(problem.staff[staffId].daysOff)

    #Cuenta dias trabajados
    dias_trabajados = dict()
    for staffId, schedule in solution.schedule.items():
        for dias in list(schedule):
            if dias != ' ':
                dias_trabajados.setdefault(staffId, []).append(1)

    #Bloquea dias si trabajo la semana anterior completa
    # for staffId, schedule in solution.schedule.items():
    for staffId in problem.staff.keys():
        if staffId in Externals:
            continue
        if sum(dias_trabajados[staffId]) > 6:
            problem.staff[staffId].maxConsecutiveShifts = 6
        else:
            problem.staff[staffId].maxConsecutiveShifts = 7

    #Actualiza feriados de la semana
    problem.feriados = list()
    tmp_dia = fecha
    for dia in range(7):
        if tmp_dia in prm.feriados.values():
            problem.feriados.append(tmp_dia.weekday())
        tmp_dia += datetime.timedelta(days=1)

    print("Feriados",problem.feriados)
    return problem, prm

def OLDUpdateConditions(problem,solution, debug = None, prm = None, week = None):
    import sys
    import numpy as np
    import Solver_codigos.instance
    #return problem

    #Si son menos de 7 dias a la semana, se asume domingo de descanso por lo que no hay problemas de restricciones
    if problem.horizon < 7:
        return problem

    #Borra restricciones antiguas
    for staffId, schedule in solution.schedule.items():
        problem.staff[staffId].daysOff = set()

    #No aplicar restricciones a internos
    Externals = [prm.IdStaff[i] for i in range(prm.TotalContract,prm.TotalContract+prm.TotalExternal)]
    items = [item for item in problem.shifts.keys()]

    #Cuenta domingos trabajados
    for staffId, schedule in solution.schedule.items():
        turno_final = list(schedule)[-1]
        if turno_final.strip() != '':
            prm.domingosPorMes[staffId] += 1
            prm.GlobaldomingosPorMes[staffId] += 1

    #Bloquea domingos
    for staffId, schedule in solution.schedule.items():
    #if prm.domingosPorMes[staffId]%2 == 0 and prm.domingosPorMes[staffId] > 0 and staffId not in Externals:
        if prm.domingosPorMes[staffId] > 1 and staffId not in Externals:
            problem.staff[staffId].daysOff.add(6)
    #prm.domingosPorMes[staffId] = 0
            if debug is not None:
                "bloquearle domingo a {}\n".format(staffId)
    #print(problem.staff[staffId].shiftOnRequests[6])
    #sys.exit()

    #Resetea el proximo mes
    if (week+1)%4==0:
        for staffId, schedule in solution.schedule.items():
            problem.staff[staffId].daysOff = set()
            prm.domingosPorMes[staffId] = 0

    #Cambia turnos cada 2 semanas
    if (week+1)%2==0:
        for id in problem.staff.keys():
            pb = problem.staff[id]
            #DEBUG
            if debug is not None:
                print("antes",pb.maxShifts)
            #DEBUG
            if id in Externals:
                continue

            if prm.ShiftsPerDay == 2:
                try:
                    pb.maxShifts[items[0]] = pb.maxShifts.pop(items[1])
                    turno_habil_semana_siguiente = items[1]
                except:
                    pb.maxShifts[items[1]] = pb.maxShifts.pop(items[0])
                    turno_habil_semana_siguiente = items[0]

            #Para el caso de 3 turnos por dias, se consideran 2 turnos de dia como intercambiables
            if prm.ShiftsPerDay == 3:
                try:
                    pb.maxShifts[items[0]] = pb.maxShifts.pop(items[2])
                    pb.maxShifts[items[1]] = 0
                    turno_habil_semana_siguiente = items[2]
                except:
                    pb.maxShifts[items[2]] = pb.maxShifts.pop(items[0])
                    del pb.maxShifts[items[1]]
                    turno_habil_semana_siguiente = items[np.random.randint(2)]

        #DEBUG
        if debug is not None:
            for id in problem.staff.keys():
                pb = problem.staff[id]
                print("despues",pb.maxShifts)
        #DEBUG

        #Bloquea primer dia del nuevo periodo si trabajo un turno de noche
        for staffId, schedule in solution.schedule.items():
            turno_final = list(schedule)[-1]
            if turno_final.strip() == items[-1] and staffId not in Externals:
                problem.staff[staffId].daysOff.add(0)

    """
    #Recomienda no trabajar los domingos despues de haber trabajado uno
    for id in problem.staff.keys():
        pb = problem.staff[id]
        if id in Externals:
            continue
        if prm.ShiftsPerDay == 2:
            try:
                turno_habil_semana_siguiente = items[1]
            except:
                turno_habil_semana_siguiente = items[0]
        if prm.ShiftsPerDay == 3:
            try:
                turno_habil_semana_siguiente = items[2]
            except:
                turno_habil_semana_siguiente = items[np.random.randint(2)]

    #DEBUG
    if debug is not None:
        for staffId, schedule in solution.schedule.items():
            print(staffId,problem.staff[staffId].shiftOffRequests)
        for staffId, schedule in solution.schedule.items():
            print(problem.staff[staffId].daysOff)
    """

    return problem

def create_month_days(prm,NumberOfWeeks):
    from datetime import timedelta
    column_names = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
    tmp_list = list(column_names)
    columnas = []
    dia = prm.dia_inicio
    prm.NumberOfWeeks = NumberOfWeeks
    # prm.NumberOfWeeks = int(round((prm.dia_fin-prm.dia_inicio).days/7))
    for i in range(0,prm.NumberOfWeeks):
        for j,item in enumerate(tmp_list):
            if dia.year%prm.dia_inicio.year == 0:
                columnas.append(tmp_list[dia.weekday()]+"-"+str(dia.day)+"-"+str(dia.month))
            else:
                columnas.append(tmp_list[dia.weekday()]+"-"+str(dia.day)+"-"+str(dia.month)+"-"+str(dia.year))
            dia += timedelta(days=1)

    return columnas

def SaveOriInstance(problem,Instancia):
    f = open(Instancia,"w")
    f.write("# This is a comment. Comments start with #\n")
    f.write("SECTION_HORIZON\n")
    f.write("# All instances start on a Monday\n")
    f.write("# The horizon length in days:\n")
    f.write(str(problem.horizon)+"\n\n")

    f.write("SECTION_SHIFTS\n")
    f.write("# ShiftID, Length in mins, Shifts which cannot follow this shift | separated\n")
    for id in problem.shifts.keys():
        f.write(str(id))
        f.write(",")
        f.write(str(problem.shifts[id].length))
        f.write(",")
        f.write(str(problem.shifts[id].prohibitNext.pop()))
        f.write("\n")
    f.write("\n")

    f.write("SECTION_STAFF\n")
    f.write("# ID, MaxShifts, MaxTotalMinutes, MinTotalMinutes, MaxConsecutiveShifts, MinConsecutiveShifts, MinConsecutiveDaysOff, MaxWeekends\n")
    for id in problem.staff.keys():
        pb = problem.staff[id]
        f.write(str(id))
        f.write(",")
        for item in problem.shifts.keys():
            try:
                f.write(str(item)+"="+str(pb.maxShifts[item]))
            except:
                f.write(str(item)+"="+str(problem.horizon))
            if item != list(problem.shifts.keys())[-1]:
                f.write("|")
        f.write(",")
        f.write(str(pb.maxTotalMinutes))
        f.write(",")
        f.write(str(pb.minTotalMinutes))
        f.write(",")
        f.write(str(pb.maxConsecutiveShifts))
        f.write(",")
        f.write(str(pb.minConsecutiveShifts))
        f.write(",")
        f.write(str(pb.minConsecutiveDaysOff))
        f.write(",")
        f.write(str(pb.maxWeekends))
        f.write("\n")
    f.write("\n")

    f.write("SECTION_DAYS_OFF\n")
    f.write("# EmployeeID, DayIndexes (start at zero)\n")
    for id in problem.staff.keys():
        if len(problem.staff[id].daysOff)==0:
            continue
        f.write(str(id))
        f.write(",")
        f.write(str(problem.staff[id].daysOff.pop()))
        f.write("\n")
    f.write("\n")

    f.write("SECTION_SHIFT_ON_REQUESTS\n")
    f.write("# EmployeeID, Day, ShiftID, Weight\n")
    for id in problem.staff.keys():
        for dia in problem.staff[id].shiftOnRequests.keys():
            f.write(str(id))
            f.write(",")
            f.write(str(problem.staff[id].shiftOnRequests[dia].day))
            f.write(",")
            f.write(str(problem.staff[id].shiftOnRequests[dia].id))
            f.write(",")
            f.write(str(problem.staff[id].shiftOnRequests[dia].weight))
            f.write("\n")
    f.write("\n")

    f.write("SECTION_SHIFT_OFF_REQUESTS\n")
    f.write("# EmployeeID, Day, ShiftID, Weight\n")
    for id in problem.staff.keys():
        for dia in problem.staff[id].shiftOffRequests.keys():
            f.write(str(id))
            f.write(",")
            f.write(str(problem.staff[id].shiftOffRequests[dia].day))
            f.write(",")
            f.write(str(problem.staff[id].shiftOffRequests[dia].id))
            f.write(",")
            f.write(str(problem.staff[id].shiftOffRequests[dia].weight))
            f.write("\n")
    f.write("\n")

    f.write("SECTION_COVER\n")
    f.write("# Day, ShiftID, Requirement, Weight for under, Weight for over\n")


    for dia in range(0,problem.horizon):
        for ShiftId in problem.cover[dia].keys():
            pb = problem.cover[dia][ShiftId]
            f.write(str(pb.day))
            f.write(",")
            f.write(str(pb.shiftId))
            f.write(",")
            f.write(str(pb.requirement))
            f.write(",")
            f.write(str(pb.weightForUnder))
            f.write(",")
            f.write(str(pb.weightForOver))
            f.write("\n")

    f.write("\n")
    f.close()
    #return problem
