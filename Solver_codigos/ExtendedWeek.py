import os
import argparse
from Solver_codigos.FuncExtendWeek import *
from Solver_codigos.WriteOutFormat import *
from joblib import Parallel, delayed
import multiprocessing as mp
import datetime
import Solver_codigos.solver
import copy
import pylab
import pandas as pd

def SolveMonthly(fecha=None, Week0=None, problem=None, Debug=None,
                 debug_folder=None, IterationTime=None, T=None, prm=None):
    # Cálculo del número de semanas a procesar
    tmp_fecha = fecha  # + datetime.timedelta(weeks=1)
    # if tmp_fecha > prm.dia_fin:
    #     numberofweeks = 0
    # else:
    #     numberofweeks = 1
    numberofweeks = 0
    while fecha.month == tmp_fecha.month:
        tmp_fecha += datetime.timedelta(weeks=1)
        numberofweeks += 1
        if tmp_fecha > prm.dia_fin:
            break

    prm.NumberOfWeeks = numberofweeks
    # Diccionarios que guarda los resultados semanales
    GlobalPlanification = dict()
    tmp_weeklySolution = dict()
    # Loop de soluciones semanales
    for week in range(Week0, Week0 + numberofweeks):
        # Archivo de depuración con información de soluciones por semanales y por proceso paralelo
        if Debug:
            debug_file = open(
                os.path.join(debug_folder, os.uname().nodename + "_output_" + str(week) + "_" + str(fecha) + ".txt"),
                "w")
        else:
            debug_file = None

        # Funcion principal, se llama al solver y se obtiene una planificación semanal
        solution, graphData = Solver_codigos.solver.Anneal(problem=problem, maxTime=IterationTime, Temperature=T,
                                            debug=debug_file)

        # Fecha
        fecha += datetime.timedelta(weeks=1)
        # Se actualizan las condiciones del problema para considerar otras restricciones
        problem, prm = UpdateConditions(problem, solution, debug=debug_file, prm=prm, fecha=fecha, week=week)

        # Se copia la solución semanal
        tmp_weeklySolution[week] = copy.deepcopy(solution)

        # Se guerdan la planificación semanal
        for staffId, schedule in solution.schedule.items():
            for item in schedule:
                GlobalPlanification.setdefault(staffId, []).append(item)
            if prm.Horizon < 7:
                for i in range(0, 7 - prm.Horizon):
                    GlobalPlanification.setdefault(staffId, []).append(' ')

        if Debug:
            # Se escribe la solución en el archivo de depuración
            solution.SaveDebug(debug_file, solution)

        if Debug:
            # Se generan gráficos de convergencia del solver para la semana
            for idx, _ in enumerate(graphData):
                x = [k[0] for k in graphData[idx]]
                y = [k[1] for k in graphData[idx]]
                y2 = [k[2] for k in graphData[idx]]
                pylab.plot(x, y, 'r')
                pylab.plot(x, y2, 'r--')
                pylab.yscale('log')
                pylab.legend(['Annealing', 'Annealing valid', 'Hill climb', 'Hill climb valid'])
                pylab.savefig(os.path.join(debug_folder, os.uname().nodename + "_output_" + str(week) + "_" + str(
                    fecha - datetime.timedelta(weeks=1)) + ".png"))

            # Se crea una planificación semanal
            Planification = []
            for schedule in solution.schedule.items():
                Planification.append(schedule[1])

            # Se escribe la solución semanal en el archivo de depuración
            debug_file.write(pd.DataFrame(Planification, index=list(problem.staff.keys())).to_string())
            debug_file.write('\nScore: {}\n'.format(solution.score))
            debug_file.write('Hard violations: {}\n'.format(solution.hardViolations))
            debug_file.write('Soft violations: {}\n'.format(solution.softViolations))
            debug_file.write('External violations: {}\n'.format(solution.externalViolations))
            debug_file.write('DaysOff violations: {}\n'.format(solution.daysOffViolations))

            print('Score:', solution.score)
            print('Hard violations:', solution.hardViolations)
            print('MinConsecutiveShifts violations:', solution.minConsecutiveShiftsViolations)
            print('MaxConsecutiveShifts violations:', solution.maxConsecutiveShiftsViolations)
            print('MinConsecutiveDaysoff violations:', solution.minConsecutiveDaysOffViolations)
            print('Weekends violations:', solution.weekendsViolations)
            print('Max Shifts violations:', solution.maxShiftsViolations)
            print('Max Total Minutes violations:', solution.maxTotalMinutesViolations)
            print('Min Total Minutes violations:', solution.minTotalMinutesViolations)
            print('Days Off violations:', solution.daysOffViolations)
            print('Soft violations:', solution.softViolations)
            print('On Request violations:', solution.onRequestViolations)
            print('Off Request violations:', solution.offRequestViolations)
            print('Cover requeriment violations:', solution.requirementViolations)
            print('Horas contrato trabajadas', solution.HorasContratoSemanales)
            print('Horas totales trabajadas', solution.HorasExtraSemanales)
            print('Domingos trabajados:', prm.GlobaldomingosPorMes)
            debug_file.close()

    return (sum([tmp_weeklySolution[week].score for week in range(Week0, Week0 + numberofweeks)]),
            tmp_weeklySolution, fecha, problem, GlobalPlanification, Week0 + numberofweeks, prm)


def Main(instancia=None, out_folder=None, IterationTime=None,
         T=None, Debug=False):
    #
    # solucion usando el esquema simulated annealing
    #

    # Revisa que se ingreso al menos un archivo de entrada
    if instancia == None or out_folder == None or IterationTime == None or T == None:
        sys.exit("Argumento de entrada faltante")

    # Se crea si es neceario, la carpeta de salida
    try:
        # Create target Directory
        os.mkdir(out_folder)
    except FileExistsError:
        pass

    # Se crea si es neceario, la carpeta de debug
    if Debug == True:
        debug_folder = os.path.join(out_folder, "Debug")
        try:
            os.mkdir(debug_folder)
        except FileExistsError:
            pass
    else:
        debug_folder = None

    # FOR PARALLEL COMPMUTATION
    NPROC = min(mp.cpu_count(), 4)

    # Solver parameters empty class
    prm = Parametros()

    # Lectura del archivo de entrada y construcción de la instancia problem
    problem, prm = ReadFromExcel(instancia, prm, DEBUG=Debug)

    # Diccionario que guarda los resultados semanales
    GlobalPlanification = dict()

    # Diccionario que guarda las instancias de soluciones semanales
    weeklySolution = dict()
    weeksWorked = dict()

    # Resolver mes a mes
    Week0 = 0
    fecha = prm.dia_inicio
    while fecha < prm.dia_fin:
        prevFecha = fecha
        prevWeek = Week0
        costosSoluciones = Parallel(n_jobs=NPROC)(delayed(SolveMonthly)(fecha=fecha, Week0=Week0, problem=problem,
                                                                        Debug=Debug, debug_folder=debug_folder,
                                                                        IterationTime=IterationTime, T=T, prm=prm) for i
                                                  in range(NPROC))
        # print(prevWeek,minimo[1][prevWeek].schedule)

        minimo = min(costosSoluciones, key=lambda t: t[0])
        # Escribir mejor solución
        weeklySolution.update(minimo[1])
        fecha = minimo[2]
        problem_ = copy.deepcopy(minimo[3])
        for key, value in minimo[4].items():
            GlobalPlanification.setdefault(key, []).extend(value)
        Week0 = minimo[5]
        prm_ = copy.deepcopy(minimo[6])

        # print(prevFecha, fecha, prevWeek)
        if fecha == prevFecha:
            break
        # Leer requerimientos del excel
        prm = Parametros()
        problem, prm = ReadFromExcel(instancia, prm, DEBUG=Debug, fecha=fecha)

        for staffId, schedule in minimo[1][prevWeek].schedule.items():
            for weeks in range(prevWeek, Week0):
                weeksWorked.setdefault(staffId, []).append(weeks)
            try:
                problem.staff[staffId].daysOff = set(problem_.staff[staffId].daysOff)
                problem.staff[staffId].maxShifts = dict(problem_.staff[staffId].maxShifts)
                problem.staff[staffId].maxConsecutiveShifts = problem_.staff[staffId].maxConsecutiveShifts
                prm.domingosPorMes[staffId] = prm_.domingosPorMes[staffId]
                prm.GlobaldomingosPorMes[staffId] = prm_.GlobaldomingosPorMes[staffId]
                prm.TurnosDeNoche[staffId] = prm_.TurnosDeNoche[staffId]
                prm.GlobalTurnosDeNoche[staffId] = prm_.GlobalTurnosDeNoche[staffId]
            except:
                pass

        # print(prm.GlobaldomingosPorMes,prm_.GlobaldomingosPorMes)
        print('in first while')
    GlobalPlanification2 = dict()
    for i in range(Week0):
        workedthisweek = list()
        for staffId, schedule in weeklySolution[i].schedule.items():
            workedthisweek.append(staffId)
        for FullstaffId in weeksWorked.keys():  # Todos los que trabajaron
            if FullstaffId in workedthisweek:  # Trabajaron en la semana
                daysworked = 0
                for staffId, schedule in weeklySolution[i].schedule.items():
                    if FullstaffId == staffId:
                        for item in schedule:
                            daysworked += 1
                            GlobalPlanification2.setdefault(staffId, []).append(item)
                        if daysworked < 7:
                            for j in range(0, 7 - daysworked):
                                GlobalPlanification2.setdefault(staffId, []).append(' ')
            else:
                for j in range(7):
                    GlobalPlanification2.setdefault(FullstaffId, []).append(' ')
        print('in for')
    # Calendario con fechas reales
    dias_y_fechas = create_month_days(prm, Week0)

    # Mostrar resultado final
    if Debug:
        print(GlobalPlanification2)
        print(dias_y_fechas)

    # Dataframes para exportar a excel
    df1 = pd.DataFrame(GlobalPlanification2, index=dias_y_fechas)
    df2 = df1.T

    # Archivo de Salida
    timestamp = str(datetime.datetime.now().strftime("%d-%m-%y_%H-%M-%S"))
    try:
        # Create target Directory
        os.mkdir(os.path.join(out_folder, timestamp))
    except FileExistsError:
        pass

    def path_leaf(path):
        import ntpath
        head, tail = ntpath.split(path)
        return tail or ntpath.basename(head)

    instancia_name = os.path.splitext(path_leaf(instancia))[0]

    from shutil import copy2
    copy2(instancia, os.path.join(os.path.join(out_folder, timestamp), instancia_name + ".xlsx"))
    out_name = "Resultado_" + instancia_name
    output_name = os.path.join(os.path.join(out_folder, timestamp), out_name + ".xlsx")

    # Escribe el archivo de salida
    return out_name, WriteOutFormat(output_name, df1, df2, prm, weeklySolution)


def solution_by_week(xls):
    mp.freeze_support()
    """
    Funcion principal para llamar al solver
    """
    parser = argparse.ArgumentParser()
    #    parser.add_argument("input_file", help="Ingrese el nombre del archivo Excel de entrada")
    parser.add_argument("-n", "--out_name", help="Ingrese el prefijo del archivo de salida", default="Salida")
    parser.add_argument("-f", "--out_folder", help="Ingrese el nombre de la carpeta de salida", default="resultados")
    parser.add_argument("-t", "--iterationtime", help="Ingrese el tiempo de iteración por semana", type=float,
                        default=200.0)
    parser.add_argument("-T", "--temperature", help="Ingrese la Temperatura inicial del solver", type=float,
                        default=40.0)
    parser.add_argument("-D", "--debug", help="Ingrese esta opcion si quiere depurar el programa", action="store_true")
    args = parser.parse_args()

    #print(args)
    # Call Main function
    return Main(instancia=xls, out_folder=args.out_folder, IterationTime=args.iterationtime, T=args.temperature, Debug=args.debug)
