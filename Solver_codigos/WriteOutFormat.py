import pandas as pd
import xlsxwriter
import sys
import base64


def xldownload(excel, name):
    data = open(excel, 'rb').read()
    b64 = base64.b64encode(data).decode('UTF-8')
    href = f'<a href="data:file/xls;base64,{b64}" download={name}>Download {name}</a>'
    return href

def write_excel_with_column_size(df,sheetname,writer):
    df.to_excel(writer, sheet_name=sheetname)  # send df to writer
    worksheet = writer.sheets[sheetname]  # pull worksheet object
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 4  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
#        print(col,idx,max_len)
    worksheet.set_column(idx+1, idx + 1, max_len)  # set column width
    worksheet.set_column(0, 0, 15)  # set column width
    #writer.save()

# def get_col_widths(dataframe):
#     # First we find the maximum length of the index column   
#     idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
#     # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
#     return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

def get_col_widths(dict):
    # First we find the maximum length of the index column   
    # idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [max([len(str(s)) for s in dict[col]] + [len(col)]) for col in dict.keys()]

def get_col_widths(list):
    # First we find the maximum length of the index column   
    # idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    return [max([len(str(s)) for s in dict[col]] + [len(elem)]) for elem in list]

def write_violation(row,col,worksheet,text_wrap,bold,weeklySolutions,Eventuales,violation,name):
    #Detalle de cada violación
    if sum([getattr(solution, violation)[worker] 
            for solution in weeklySolutions 
                for worker in getattr(solution, violation).keys()]) > 0:
        for i,solution in enumerate(weeklySolutions):
            worksheet.write(row , col+i+1, "Semana "+str(i+1),bold)

        worksheet.write(row , col+0, name, text_wrap)
        worksheet.set_row(row, 40)
        total_lines = 0
        for week,solution in enumerate(weeklySolutions):
            worker_line = 0
            for worker in getattr(solution, violation).keys():
                if worker in Eventuales:
                    continue
                worker_line += 1
                worksheet.write(row+worker_line, col+0, worker,bold)
                if getattr(solution, violation)[worker] > 0:
                    worksheet.write(row+worker_line, col+1 + week, str(getattr(solution, violation)[worker]))
                if worker_line > total_lines:
                    total_lines = worker_line
        row += 2 + total_lines

    return row

def write_monthly_schedule(weeklySolutions,dias_semana,worksheet,bold):
    #WRITE SOLUTION SCHEDULE
    row = 0
    for i, solution in enumerate(weeklySolutions):
        worksheet.write(row , 0, "Semana "+str(i+1))
        worksheet.write_row(row, 1, dias_semana[i],bold)
        row += 1        
        for key in solution.schedule.keys():
            worksheet.write(row, 0, key, bold)
            worksheet.write_row(row, 1, solution.schedule[key])
            row += 1        
        row += 1

    #SET COLUMN WIDTH
    widths = dict()
    for i, solution in enumerate(weeklySolutions):
        for j, dias in enumerate(dias_semana[i]):
            widths.setdefault(j+1,[]).append(len(str(dias)) + 2) 
        for key in solution.schedule.keys():
            widths.setdefault(0,[]).append(len(str(key)) + 2) 
            for j,dias in enumerate(solution.schedule[key]):
                widths.setdefault(j+1,[]).append(len(str(dias)) + 2) 

    max_width = dict()
    for key in widths.keys():
        max_width[key] = max(widths[key])

    for key in max_width.keys():
        worksheet.set_column(key, key, max_width[key])

    return row

def write_otros(row,col,worksheet,weeklySolutions,text_wrap,bold,Eventuales,conteo,name):
    worksheet.write(row , col+0, name, text_wrap)
    worksheet.set_row(row, 30)
    cantidad_por_persona = dict()
    for solution in weeklySolutions:
        for worker in getattr(solution, conteo).keys():
            cantidad_por_persona.setdefault(worker,[]).append(getattr(solution, conteo)[worker])

    worksheet.write(row,col+1,"Cantidad",bold)
    row += 1
    for worker in cantidad_por_persona.keys():
        if worker in Eventuales:
            continue
        worksheet.write(row,col,worker,bold)
        worksheet.write(row,col+1,sum(cantidad_por_persona[worker]))
        row += 1

    row +=1
    return row

def write_monthly_violations(row,col,worksheet,weeklySolutions,text_wrap,bold,Eventuales):

    #Costos
    bold_wrap = bold
    bold_wrap.set_text_wrap()
    worksheet.write(row, col+1 ,"Cuantificable",bold_wrap)
    worksheet.write(row, col+2 ,"Interno",bold_wrap)
    worksheet.write(row, col+3 ,"Total",bold_wrap)
    row += 1
    worksheet.write(row, col+0 ,"Costo mensual",bold_wrap)

    costo_cuanti_mensual  = sum([solution.scoreCuantificable for solution in weeklySolutions])
    costo_interno_mensual = sum([solution.scoreInterno for solution in weeklySolutions])
    costo_mensual = sum([solution.score for solution in weeklySolutions])
    worksheet.write(row,col+1,costo_cuanti_mensual)
    worksheet.write(row,col+2,costo_interno_mensual)
    worksheet.write(row,col+3,costo_mensual)
    row += 2

    #Número de violaciones
    worksheet.write(row, col+1 ,"Obligatorias",bold_wrap)
    worksheet.write(row, col+2 ,"Otras",bold_wrap)
    row += 1
    worksheet.write(row, col+0 ,"Número de violaciones",bold_wrap)
    worksheet.set_row(row, 30)

    dias_sobra_gente = 0
    for solution in weeklySolutions:
        for dia in solution.requirementViolationsSobran.keys():
            if solution.requirementViolationsSobran[dia] > 0:
                dias_sobra_gente += solution.requirementViolationsSobran[dia]

    violaciones_fuertes = sum([solution.softViolations for solution in weeklySolutions]) - dias_sobra_gente
    violaciones_debiles = sum([solution.hardViolations for solution in weeklySolutions])

    worksheet.write(row, col+1 ,violaciones_fuertes)
    worksheet.write(row, col+2 ,violaciones_debiles)
    row += 2

    piv_col = col
    #Horas Contrato trabajadas
    write_otros(row,col,worksheet,weeklySolutions,text_wrap,bold,Eventuales,
                            'HorasContratoSemanales',"Horas Contrato")
    col += 2
    #Horas Extra trabajadas
    write_otros(row,col,worksheet,weeklySolutions,text_wrap,bold,Eventuales,
                            'HorasExtraSemanales',"Horas Extra")
    col += 2
    #Horas Totales trabajadas
    write_otros(row,col,worksheet,weeklySolutions,text_wrap,bold,Eventuales,
                            'TotalHorasTrabajadas',"Total Horas")
    col += 2
    #Domingos trabajados
    row = write_otros(row,col,worksheet,weeklySolutions,text_wrap,bold,Eventuales,
                            'DomingoTrabajado',"Domingos trabajados")

    col = piv_col
    #Requirement
    row = write_violation(row,col,worksheet,text_wrap,bold,weeklySolutions,Eventuales,
                            'requirementViolationsFaltan',"Falta de trabajadores")
    row = write_violation(row,col,worksheet,text_wrap,bold,weeklySolutions,Eventuales,
                            'requirementViolationsSobran',"Exceso de trabajadores")

    #Gente contratada
    row = write_violation(row,col,worksheet,text_wrap,bold,weeklySolutions,Eventuales,
                            'ViolacionContratadosPorDia',"Contratados en el turno")

    #Max consecutive shifts
    row = write_violation(row,col,worksheet,text_wrap,bold,weeklySolutions,Eventuales,
                            'maxConsecutiveShiftsViolations',"Máximo de turnos consecutivos")
    #Max Shift
    row = write_violation(row,col,worksheet,text_wrap,bold,weeklySolutions,Eventuales,
                            'maxShiftsViolations',"Mantener turnos")
    #Min hours
    row = write_violation(row,col,worksheet,text_wrap,bold,weeklySolutions,Eventuales,
                            'minTotalMinutesViolations',"Mínimo de horas trabajadas")
    #Max hours
    row = write_violation(row,col,worksheet,text_wrap,bold,weeklySolutions,Eventuales,
                            'maxTotalMinutesViolations',"Horas extras")
    #Days Off
    row = write_violation(row,col,worksheet,text_wrap,bold,weeklySolutions,Eventuales,
                            'daysOffViolations',"Domingo o lunes de descanso")

    return costo_cuanti_mensual, costo_interno_mensual, costo_mensual

def write_total_schedule_per_month(row,worksheet,weeklySolutions,dias,bold):
    # Some data we want to write to the worksheet.
    for solution in weeklySolutions:
        IdTurnos = [turnos for turnos in solution.IdTurnos]
        break
    NroTurnos = len(IdTurnos)
    Requirement = max([solution.Requerimientos for solution in weeklySolutions])
    # Workers = set([key for solution in weeklySolutions for key in solution.schedule.keys()])

    GlobalPlanification = dict()
    for solution in weeklySolutions:
        for staffId, schedule in solution.schedule.items():
            for item in schedule:
                GlobalPlanification.setdefault(staffId,[]).append(item)
            if len(schedule) < 7:
                for i in range(0,7-len(schedule)):
                    GlobalPlanification.setdefault(staffId,[]).append(' ')


    calendar = [{} for i in range(0,NroTurnos)]
    tot_dias = [dia for i in dias for dia in i]

    for i,day in enumerate(tot_dias):
        worksheet.write(row+(i)*Requirement+1,0,day,bold)

    # print(GlobalPlanification)
    for key in GlobalPlanification.keys():
        for day, nombre in enumerate(tot_dias):
            for i in range(NroTurnos):
                if GlobalPlanification[key][day].strip() == IdTurnos[i]:
                    calendar[i].setdefault(day,[]).append(key)
                    
    for turno in range(NroTurnos):
        for day in sorted(calendar[turno]):
            for i,worker in enumerate(calendar[turno][day]):
                worksheet.write(row+(day)*Requirement+1+i,2+turno,worker)

    row += (day)*Requirement+1+i
    return row


def WriteOutFormat(output_name, df1, df, prm, weeklySolution):
    import pandas as pd
    import datetime

    writer = pd.ExcelWriter(output_name, engine='xlsxwriter')
    # Opciones básicas
    write_excel_with_column_size(df1, 'Solución', writer)

    workbook = writer.book
    worksheet1 = workbook.add_worksheet("Calendario")
    bold = workbook.add_format({'bold': True})
    text_wrap = workbook.add_format()
    text_wrap.set_text_wrap()

    # Columnas calendario Global
    NroTurnos = min(len(pd.unique(df.values.ravel('K'))) - 1, 3)
    for i in range(NroTurnos):
        worksheet1.write(0, i + 2, "Turno " + str(i + 1), bold)

    nro_a_mes = {1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio', 7: 'Julio',
                 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'}

    # Escritura mes a mes
    weekspermonth = dict()
    dias_semana = dict()
    nombres_semana = list(df.columns.values)
    fecha = prm.dia_inicio
    for week in range(0, prm.NumberOfWeeks):
        if fecha.year % prm.dia_inicio.year == 0:
            weekspermonth.setdefault(fecha.month, []).append(week)
        else:
            weekspermonth.setdefault(str(fecha.year) + '-' + str(fecha.month), []).append(week)
            nro_a_mes[str(fecha.year) + '-' + str(fecha.month)] = nro_a_mes[fecha.month] + '-' + str(fecha.year)
        dias_semana[week] = nombres_semana[week * 7:(week + 1) * 7]
        fecha += datetime.timedelta(weeks=1)

    col = 0
    tot_sche_row = 0
    for month in weekspermonth.keys():
        tmp_worksheet = workbook.add_worksheet(nro_a_mes[month])
        month_solutions = [weeklySolution[i] for i in weekspermonth[month]]
        month_days = [dias_semana[i] for i in weekspermonth[month]]
        # write monthly solution
        next_row = write_monthly_schedule(month_solutions, month_days,
                                          tmp_worksheet, bold)

        # write monthly violations and costs
        write_monthly_violations(next_row, col, tmp_worksheet, month_solutions, text_wrap, bold, prm.IdStaffEventual)

        # write_total_schedule
        tot_sche_row = write_total_schedule_per_month(tot_sche_row, worksheet1, month_solutions, month_days, bold)

    # Write complete year
    dias_semana = dict()
    nombres_semana = list(df.columns.values)
    fecha = prm.dia_inicio
    for week in range(0, prm.NumberOfWeeks):
        dias_semana[week] = nombres_semana[week * 7:(week + 1) * 7]
        fecha += datetime.timedelta(weeks=1)
    solutions = [weeklySolution[i] for i in weeklySolution.keys()]
    dias = [dias_semana[i] for i in weeklySolution.keys()]

    # Write total violations plus costs plus dias trabajados
    next_row = 0
    col = 6
    write_monthly_violations(next_row, col, worksheet1, solutions, text_wrap, bold, prm.IdStaffEventual)
    worksheet1.set_column(0, 0, 15)
    worksheet1.set_column(1, 1, 3)
    for i in range(0, 8):
        worksheet1.set_column(6 + i, 6 + i, 15)
    for i in range(0, 3):
        worksheet1.set_column(2 + i, 2 + i, 13)

    workbook.close()
    return writer
    #writer.save()
    #writer.close()


def WriteOutFormatandCosts(output_name, df1, df, prm, weeklySolution):
    import pandas as pd
    import datetime

    writer = pd.ExcelWriter(output_name, engine='xlsxwriter')
    # Opciones básicas
    write_excel_with_column_size(df1, 'Solución', writer)

    workbook = writer.book
    worksheet1 = workbook.add_worksheet("Calendario")
    bold = workbook.add_format({'bold': True})
    text_wrap = workbook.add_format()
    text_wrap.set_text_wrap()

    # Columnas calendario Global
    NroTurnos = min(len(pd.unique(df.values.ravel('K'))) - 1, 3)
    for i in range(NroTurnos):
        worksheet1.write(0, i + 2, "Turno " + str(i + 1), bold)

    nro_a_mes = {1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio', 7: 'Julio',
                 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'}

    # Escritura mes a mes
    weekspermonth = dict()
    dias_semana = dict()
    nombres_semana = list(df.columns.values)
    fecha = prm.dia_inicio
    for week in range(0, prm.NumberOfWeeks):
        if fecha.year % prm.dia_inicio.year == 0:
            weekspermonth.setdefault(fecha.month, []).append(week)
        else:
            weekspermonth.setdefault(str(fecha.year) + '-' + str(fecha.month), []).append(week)
            nro_a_mes[str(fecha.year) + '-' + str(fecha.month)] = nro_a_mes[fecha.month] + '-' + str(fecha.year)
        dias_semana[week] = nombres_semana[week * 7:(week + 1) * 7]
        fecha += datetime.timedelta(weeks=1)

    col = 0
    tot_sche_row = 0
    for month in weekspermonth.keys():
        tmp_worksheet = workbook.add_worksheet(nro_a_mes[month])
        month_solutions = [weeklySolution[i] for i in weekspermonth[month]]
        month_days = [dias_semana[i] for i in weekspermonth[month]]
        # write monthly solution
        next_row = write_monthly_schedule(month_solutions, month_days,
                                          tmp_worksheet, bold)

        # write monthly violations and costs
        write_monthly_violations(next_row, col, tmp_worksheet, month_solutions, text_wrap, bold, prm.IdStaffEventual)

        # write_total_schedule
        tot_sche_row = write_total_schedule_per_month(tot_sche_row, worksheet1, month_solutions, month_days, bold)

    # Write complete year
    dias_semana = dict()
    nombres_semana = list(df.columns.values)
    fecha = prm.dia_inicio
    for week in range(0, prm.NumberOfWeeks):
        dias_semana[week] = nombres_semana[week * 7:(week + 1) * 7]
        fecha += datetime.timedelta(weeks=1)
    solutions = [weeklySolution[i] for i in weeklySolution.keys()]
    dias = [dias_semana[i] for i in weeklySolution.keys()]

    # Write total violations plus costs plus dias trabajados
    next_row = 0
    col = 6
    costo_cuant, cost_interno, costo_total = write_monthly_violations(next_row, col, worksheet1, solutions, text_wrap, bold, prm.IdStaffEventual)
    worksheet1.set_column(0, 0, 15)
    worksheet1.set_column(1, 1, 3)
    for i in range(0, 8):
        worksheet1.set_column(6 + i, 6 + i, 15)
    for i in range(0, 3):
        worksheet1.set_column(2 + i, 2 + i, 13)

    workbook.close()
    return costo_cuant, cost_interno, costo_total





