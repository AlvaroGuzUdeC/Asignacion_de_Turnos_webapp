import pandas as pd
import numpy as np
from Solver_codigos.FuncExtendWeek import *
from Solver_codigos.WriteOutFormat import *
import Solver_codigos.validator as validator
from Solver_codigos.solver import GenerateInitialConfiguration, CreateEmptySolution
import sys
import copy
import datetime
import os
import traceback

def read_solution(df,solution,week,NroTrabajadores,horizon):
    for staff in df.columns[1:NroTrabajadores+1]:
        list_staff = []
        valid_index = list(range((week)*7,(week+1)*7))
        dia = 1
        for i in valid_index:
            list_staff.append(df.loc[i,staff])
            dia += 1
            if dia > horizon:
                break
        solution.schedule[staff] = list_staff
    return solution

def Main(instance_name=None,solution_name=None,Debug=True):
    #Read Solution
    df = pd.read_excel(solution_name)
    df.replace('', ' ', inplace=True)
    df.replace(np.nan, ' ', inplace=True)

    #Read Instance
    prm = Parametros()
    #Lectura del archivo de entrada y construcción de la instancia problem
    problem, prm = ReadFromExcel(instance_name, prm, DEBUG=Debug)
    solution = GenerateInitialConfiguration(problem)

    #Diccionario que guarda las instancias de soluciones semanales
    weeklySolution = dict()
    weeksWorked = dict()

    #Resolver mes a mes
    Week0 = 0
    fecha = prm.dia_inicio
    numberofweeks = prm.NumberOfWeeks
    for week in range(numberofweeks):
        prevWeek = Week0
        prevFecha = fecha
        #Leer solucion
        print(fecha,problem.horizon)
        # raise
        solution = read_solution(df,solution,week,len(problem.staff.keys()),problem.horizon)
        #Calcular Costo
        validator.CalculatePenalty(solution, problem)
        #Fecha
        fecha += datetime.timedelta(weeks=1)
        #Se actualizan las condiciones del problema para considerar otras restricciones
        problem, prm = UpdateConditions(problem, solution, debug = Debug, prm = prm, fecha = fecha, week = week)
        #Escribir mejor solución
        weeklySolution[week] = copy.deepcopy(solution)
        #Prev Problem
        problem_ = copy.deepcopy(problem)
        Week0 += 1
        prm_ = copy.deepcopy(prm)

        for staffId, schedule in solution.schedule.items():
            weeksWorked.setdefault(staffId, []).append(week)
        #Leer requerimientos del excel
        if prevFecha.month != fecha.month:
            prm = Parametros()
            problem, prm = ReadFromExcel(instance_name, prm, DEBUG=Debug, fecha = fecha)

            for staffId, schedule in solution.schedule.items():
                # weeksWorked.setdefault(staffId, []).append(week)
                try:
                    problem.staff[staffId].daysOff = set(problem_.staff[staffId].daysOff)
                    problem.staff[staffId].maxShifts = dict(problem_.staff[staffId].maxShifts)
                    problem.staff[staffId].maxConsecutiveShifts = problem_.staff[staffId].maxConsecutiveShifts
                    prm.domingosPorMes[staffId] = prm_.domingosPorMes[staffId]
                    prm.GlobaldomingosPorMes[staffId] = prm_.GlobaldomingosPorMes[staffId]
                    prm.TurnosDeNoche[staffId] = prm_.TurnosDeNoche[staffId]
                    prm.GlobalTurnosDeNoche[staffId] = prm_.GlobalTurnosDeNoche[staffId]
                except Exception:
                    print(traceback.format_exc())
                    pass
            solution = GenerateInitialConfiguration(problem)

    GlobalPlanification2 = dict()
    for i in range(numberofweeks):
        workedthisweek = list()
        for staffId, schedule in weeklySolution[i].schedule.items():
            workedthisweek.append(staffId)
        for FullstaffId in weeksWorked.keys():#Todos los que trabajaron
            if FullstaffId in workedthisweek:#Trabajaron en la semana
                daysworked = 0
                for staffId, schedule in weeklySolution[i].schedule.items():
                    if FullstaffId == staffId:
                        for item in schedule:
                            daysworked += 1
                            GlobalPlanification2.setdefault(staffId,[]).append(item)
                        if daysworked < 7:
                            for j in range(0,7-daysworked):
                                GlobalPlanification2.setdefault(staffId,[]).append(' ')
            else:
                for j in range(7):
                    GlobalPlanification2.setdefault(FullstaffId,[]).append(' ')
            

    #Calendario con fechas reales
    dias_y_fechas = create_month_days(prm,numberofweeks)

    #Mostrar resultado final
    if Debug:
        print(GlobalPlanification2)
        print(dias_y_fechas)

    #Dataframes para exportar a excel
    df1 = pd.DataFrame(GlobalPlanification2,index=dias_y_fechas)
    df2 = df1.T

    #Archivo de Salida
    output_name = solution_name.name[:-5]+"_Corregido.xlsx"

    return WriteOutFormatandCosts(f'Resultados/{output_name}',df1,df2,prm,weeklySolution)


def get_costos_from_resultadoxls(xls_upladed):
    xls = pd.ExcelFile(xls_upladed)
    calendario_sheet = pd.read_excel(xls, 'Calendario')
    return list(calendario_sheet.iloc[0, 7:10])

def ComputeSolutionCosts(uploadre_instancia, uploadre_resultado_modificado):
    """
    Funcion principal para llamar al solver
    """

    instancia = uploadre_instancia
    resultado = uploadre_resultado_modificado

    return Main(instance_name=instancia, solution_name=resultado)