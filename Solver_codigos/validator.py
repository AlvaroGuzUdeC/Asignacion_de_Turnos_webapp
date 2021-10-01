import math

# start class
class StaffMemberResult:
    
    def __init__(self):
        self.id = ''
        self.totalMinutes = 0
        self.minConsecutiveShifts = float('inf')
        self.maxConsecutiveShifts = 0
        self.minConsecutiveDaysOff = float('inf')
        self.weekends = 0
        self.offRequestPenalty = 0
        self.onRequestPenalty = 0
        self.hardViolations = 0
        self.minConsecutiveShiftsViolations = dict() #aqui
        self.maxConsecutiveShiftsViolations = dict() #aqui
        self.minConsecutiveDaysOffViolations = dict() #aqui
        self.weekendsViolations = dict() #aqui
        self.maxShiftsViolations = dict() #aqui
        self.maxTotalMinutesViolations = dict() #aqui
        self.minTotalMinutesViolations = dict() #aqui
        self.daysOffViolations = dict() #aqui
        self.softViolations = 0 #aqui
        self.onRequestViolations = dict() #aqui
        self.offRequestViolations = dict() #aqui
        self.externalViolations = 0 #aqui
        self.requirementViolations = dict() #aqui   
        
    def BuildInfo(self, solution, problem, staffId):
        self.id = staffId
        self.minConsecutiveShiftsViolations[staffId] = 0  #aqui
        self.maxConsecutiveShiftsViolations[staffId] = 0 #aqui
        self.minConsecutiveDaysOffViolations[staffId] = 0 #aqui
        self.weekendsViolations[staffId] = 0 #aqui
        self.maxShiftsViolations[staffId] = 0 #aqui
        self.minTotalMinutesViolations[staffId] = 0 #aqui
        self.maxTotalMinutesViolations[staffId] = 0 #aqui
        self.daysOffViolations[staffId] = 0 #aqui
        self.onRequestViolations[staffId] = 0 #aqui
        self.offRequestViolations[staffId] = 0 #aqui
        solution.CostoParcial[staffId] = 0
        solution.scoreCuantificableParcial[staffId] = 0
        solution.HorasContratoSemanales[staffId] = 0
        solution.HorasExtraSemanales[staffId] = 0
        solution.TotalHorasTrabajadas[staffId] = 0
        staffMember = problem.staff[staffId]
        memberSchedule = solution.schedule[staffId]
        lastShift = ''
        shiftsTaken = dict()
        for idx, shift in enumerate(memberSchedule):
            shiftsTaken[shift] = shiftsTaken.get(shift, 0) + 1

            if shift != ' ':
                self.totalMinutes += problem.shifts[shift].length

                if idx in staffMember.daysOff: #aqui
                    self.hardViolations += 1 #aqui
                    self.daysOffViolations[staffId] += 1 #aqui
                    solution.CostoParcial[staffId] += problem.costos.daysOffWeight
                    solution.scoreCuantificableParcial[staffId] += problem.costos.daysOffWeight
                    
                
                if lastShift != '' and lastShift != ' ' and shift in problem.shifts[lastShift].prohibitNext:
                    self.hardViolations += 1
                    solution.CostoParcial[staffId] += problem.costos.prohibitNextWeight

            if idx in staffMember.shiftOnRequests and staffMember.shiftOnRequests[idx].id != shift:
#                self.softViolations += 1 #aqui
                self.onRequestPenalty += staffMember.shiftOnRequests[idx].weight
                self.onRequestViolations[staffId] += 1 #aqui
                solution.CostoParcial[staffId] += staffMember.shiftOnRequests[idx].weight

            if idx in staffMember.shiftOffRequests and staffMember.shiftOffRequests[idx].id == shift:
#                self.softViolations += 1 #aqui
                self.offRequestPenalty += staffMember.shiftOffRequests[idx].weight
                self.offRequestViolations[staffId] += 1 #aqui
                solution.CostoParcial[staffId] += staffMember.shiftOffRequests[idx].weight

            lastShift = shift
        
        lastShift = memberSchedule[0]
        consecutiveShifts = 0
        consecutiveDaysOff = 0
        if lastShift != ' ':
            consecutiveShifts += 1

        for shift in memberSchedule[1:]:
            if shift != ' ':
                if lastShift == ' ' and self.maxConsecutiveShifts != 0:
                    if self.minConsecutiveDaysOff > consecutiveDaysOff:
                        self.minConsecutiveDaysOff = consecutiveDaysOff

                consecutiveDaysOff = 0
                consecutiveShifts += 1
            else:
                if lastShift != ' ':
                    if self.maxConsecutiveShifts < consecutiveShifts:
                        self.maxConsecutiveShifts = consecutiveShifts
                    if self.minConsecutiveShifts > consecutiveShifts:
                        self.minConsecutiveShifts = consecutiveShifts
                
                consecutiveShifts = 0
                consecutiveDaysOff += 1
            lastShift = shift
        
        if consecutiveShifts != 0:
            if self.maxConsecutiveShifts < consecutiveShifts:
                self.maxConsecutiveShifts = consecutiveShifts
            if self.minConsecutiveShifts > consecutiveShifts:
                self.minConsecutiveShifts = consecutiveShifts

        if self.minConsecutiveShifts == float('inf'):
            self.minConsecutiveShifts = 0

        totalWeekends = int(problem.horizon / 7)
        weekendWorking = [[memberSchedule[week*7 + x] != ' ' for x in [5, 6]] for week in range(totalWeekends)]

        weekendIndex=[] #OJO CON ESTO QUE NO HACE NADA
        if problem.horizon % 7 == 6:
            weekendIndex.append([memberSchedule[problem.horizon - 1] != ' '])
        for weekend in weekendWorking:
            if any(weekend):
                self.weekends += 1

        ViolaCambioTurno1 = False
        #Para forzar cambio de turno
        for shift, count in shiftsTaken.items():
            if staffMember.maxShifts.get(shift, problem.horizon) < count:
                self.hardViolations += 1
                self.maxShiftsViolations[staffId] +=1 #aqui
                solution.CostoParcial[staffId] += problem.costos.maxShiftsWeight + 10
                solution.scoreCuantificableParcial[staffId] += problem.costos.maxShiftsWeight
                solution.scoreInterno += 10
                ViolaCambioTurno1 = True

        ViolaCambioTurno2 = False
        #Test "keep shift" new condition
        shiftsTaken.pop(' ', None)
        if staffMember.Contratado == True:
            for shift, count in shiftsTaken.items():
                if shift == list(problem.shifts.keys())[-1] and count > 0:
                    for other_shifts, other_count in shiftsTaken.items():
                        if other_shifts == shift:
                            continue
                        if other_count > 0:
                            self.hardViolations += 1
                            self.maxShiftsViolations[staffId] +=1 #aqui
                            solution.CostoParcial[staffId] += problem.costos.maxShiftsWeight + 10
                            solution.scoreCuantificableParcial[staffId] += problem.costos.maxShiftsWeight
                            solution.scoreInterno += 10
                            ViolaCambioTurno2 = True

        if ViolaCambioTurno1 and ViolaCambioTurno2:
            self.hardViolations -= 1
            self.maxShiftsViolations[staffId] -=1 #aqui
            solution.scoreCuantificableParcial[staffId] -= problem.costos.maxShiftsWeight
            solution.scoreInterno += problem.costos.maxShiftsWeight + 20

        if self.totalMinutes < staffMember.minTotalMinutes:
#            self.hardViolations += 1
            # solution.CostoParcial[staffId] += (problem.HorasPorSemana*60-self.totalMinutes) * (staffMember.Costo/60)
            # solution.scoreCuantificableParcial[staffId] += (problem.HorasPorSemana*60-self.totalMinutes) * (staffMember.Costo/60)
            solution.CostoParcial[staffId] += (staffMember.minTotalMinutes-self.totalMinutes) * (staffMember.Costo/60)
            solution.scoreCuantificableParcial[staffId] += (staffMember.minTotalMinutes-self.totalMinutes) * (staffMember.Costo/60)
            self.minTotalMinutesViolations[staffId] += 1 #aqui


        #COSTOS
        if self.totalMinutes > staffMember.maxTotalMinutes:
            if self.totalMinutes > staffMember.maxTotalMinutes+problem.maxExtraHours*60:
                self.hardViolations += 1
                solution.CostoParcial[staffId] += problem.costos.maxTotalMinutesWeight
                solution.scoreCuantificableParcial[staffId] += problem.costos.maxTotalMinutesWeight
                self.maxTotalMinutesViolations[staffId] += 1 #aqui

        # print(problem.feriados)
        minutosLegalesPorDia = problem.HorasPorSemana*(60/6)
        if staffMember.Contratado == True:
            for dia, shift in enumerate(memberSchedule):
                if shift != ' ':
                    if problem.shifts[shift].length > minutosLegalesPorDia:
                        if dia != 6 or dia not in problem.feriados:
                            solution.CostoParcial[staffId] += minutosLegalesPorDia * (staffMember.Costo/60) #largo del bloque en minutos, costo en horas
                            solution.scoreCuantificableParcial[staffId] += minutosLegalesPorDia * (staffMember.Costo/60)
                            solution.HorasContratoSemanales[staffId] += minutosLegalesPorDia / 60
                        else:
                            solution.CostoParcial[staffId] += minutosLegalesPorDia * (staffMember.CostoExtra100/60) #largo del bloque en minutos, costo en horas
                            solution.scoreCuantificableParcial[staffId] += minutosLegalesPorDia * (staffMember.CostoExtra100/60)
                            solution.HorasContratoSemanales[staffId] += minutosLegalesPorDia / 60
                        if dia != 6 or dia not in problem.feriados:
                            solution.CostoParcial[staffId] += (problem.shifts[shift].length-minutosLegalesPorDia)*(staffMember.CostoExtra50/60)
                            solution.scoreCuantificableParcial[staffId] += (problem.shifts[shift].length-minutosLegalesPorDia)*(staffMember.CostoExtra50/60)
                            solution.HorasExtraSemanales[staffId]  += (problem.shifts[shift].length-minutosLegalesPorDia)/60
                        else:
                            solution.CostoParcial[staffId] += problem.shifts[shift].length*(staffMember.CostoExtra100/60)
                            solution.scoreCuantificableParcial[staffId] += problem.shifts[shift].length*(staffMember.CostoExtra100/60)
                            solution.HorasExtraSemanales[staffId]  += problem.shifts[shift].length/60
                    else:
                        solution.CostoParcial[staffId] += problem.shifts[shift].length * (staffMember.Costo/60) #largo del bloque en minutos, costo en horas
                        solution.scoreCuantificableParcial[staffId] += problem.shifts[shift].length * (staffMember.Costo/60)
                        solution.HorasContratoSemanales[staffId] += problem.shifts[shift].length / 60
            solution.TotalHorasTrabajadas[staffId]=solution.HorasExtraSemanales[staffId]+solution.HorasContratoSemanales[staffId]
        else:
            for dia,shift in enumerate(memberSchedule):
                if shift != ' ':
                    if dia != 6 or dia not in problem.feriados:
                        solution.CostoParcial[staffId] += problem.shifts[shift].length * (staffMember.Costo/60) #largo del bloque en minutos, costo en horas
                        solution.scoreCuantificableParcial[staffId] += problem.shifts[shift].length * (staffMember.Costo/60)
                    else:
                        solution.CostoParcial[staffId] += problem.shifts[shift].length * (staffMember.CostoExtraExterno/60) #largo del bloque en minutos, costo en horas
                        solution.scoreCuantificableParcial[staffId] += problem.shifts[shift].length * (staffMember.CostoExtraExterno/60)


        if self.minConsecutiveShifts < staffMember.minConsecutiveShifts:
            self.hardViolations += 1
            self.minConsecutiveShiftsViolations[staffId] += 1  #aqui
            solution.CostoParcial[staffId] += problem.costos.minConsecutiveShiftsWeight
                        
        if self.maxConsecutiveShifts > staffMember.maxConsecutiveShifts:
            self.hardViolations += 1
            self.maxConsecutiveShiftsViolations[staffId] += 1 #aqui
            solution.CostoParcial[staffId] += problem.costos.maxConsecutiveShiftsWeight
            solution.scoreCuantificableParcial[staffId] += problem.costos.maxConsecutiveShiftsWeight
            
        if self.minConsecutiveDaysOff < staffMember.minConsecutiveDaysOff:
            self.hardViolations += 1
            self.minConsecutiveDaysOffViolations[staffId] += 1 #aqui
            solution.CostoParcial[staffId] += problem.costos.minConsecutiveDaysOffWeight
            
        if self.weekends > staffMember.maxWeekends:
            self.hardViolations += 1
            self.weekendsViolations[staffId] += 1 #aqui
            solution.CostoParcial[staffId] += problem.costos.weekendsWeight
            
    def CalculatePenalty(self):
        return self.offRequestPenalty + self.onRequestPenalty


def CalculatePenalty(solution, problem):
    totalPenalty = 0
    solution.hardViolations = 0
    solution.minConsecutiveShiftsViolations = dict() #aqui
    solution.maxConsecutiveShiftsViolations = dict() #aqui
    solution.minConsecutiveDaysOffViolations = dict() #aqui
    solution.weekendsViolations = dict() #aqui
    solution.maxShiftsViolations = dict() #aqui
    solution.minTotalMinutesViolations = dict() #aqui
    solution.maxTotalMinutesViolations = dict() #aqui
    solution.daysOffViolations = dict() #aqui
    solution.softViolations = 0 #aqui
    solution.onRequestViolations = dict() #aqui
    solution.offRequestViolations = dict() #aqui
    solution.externalViolations = 0 #aqui
    solution.requirementViolations = dict() #aqui
    solution.CostoParcial = dict()
    solution.HorasContratoSemanales = dict()
    solution.HorasExtraSemanales = dict()
    solution.TotalHorasTrabajadas = dict()
    solution.NumeroContratados = 0
    solution.ViolacionContratadosPorDia = dict()
    solution.scoreCuantificableParcial = dict()
    solution.scoreCuantificable = 0
    solution.scoreInterno = 0
    solution.requirementViolationsFaltan = dict()
    solution.requirementViolationsSobran = dict()

    for staffId in problem.staff.keys():
        staffMemberResult = StaffMemberResult()
        staffMemberResult.BuildInfo(solution, problem, staffId )
        totalPenalty += staffMemberResult.hardViolations * problem.hardConstraintWeight
        totalPenalty += staffMemberResult.externalViolations * problem.hardConstraintWeight#aqui
        totalPenalty += staffMemberResult.CalculatePenalty()
        solution.hardViolations += staffMemberResult.hardViolations
        solution.minConsecutiveShiftsViolations[staffId] = staffMemberResult.minConsecutiveShiftsViolations[staffId] #aqui
        solution.maxConsecutiveShiftsViolations[staffId] = staffMemberResult.maxConsecutiveShiftsViolations[staffId] #aqui
        solution.minConsecutiveDaysOffViolations[staffId] = staffMemberResult.minConsecutiveDaysOffViolations[staffId] #aqui
        solution.weekendsViolations[staffId] = staffMemberResult.weekendsViolations[staffId] #aqui
        solution.maxShiftsViolations[staffId] = staffMemberResult.maxShiftsViolations[staffId] #aqui
        solution.maxTotalMinutesViolations[staffId] = staffMemberResult.maxTotalMinutesViolations[staffId] #aqui
        solution.minTotalMinutesViolations[staffId] = staffMemberResult.minTotalMinutesViolations[staffId] #aqui
        solution.daysOffViolations[staffId] = staffMemberResult.daysOffViolations[staffId] #aqui
        solution.softViolations += staffMemberResult.softViolations #aqui (request on or off)
        solution.onRequestViolations[staffId] = staffMemberResult.onRequestViolations[staffId] #aqui
        solution.offRequestViolations[staffId] = staffMemberResult.offRequestViolations[staffId] #aqui
        solution.externalViolations += staffMemberResult.externalViolations #aqui

    CostoTotal = sum(solution.CostoParcial.values())

    dia2name = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
    for day in range(problem.horizon):
        solution.requirementViolationsFaltan[dia2name[day]] = 0 #aqui
        solution.requirementViolationsSobran[dia2name[day]] = 0 #aqui
        for shift, cover in problem.cover[day].items():
            count = 0
            for schedule in solution.schedule.values():
                if schedule[day] == shift:
                    count += 1

            if cover.requirement < count:
                solution.softViolations += 1
                totalPenalty += (count - cover.requirement) * cover.weightForOver 
                solution.requirementViolationsSobran[dia2name[day]] += 1 #aqui
                CostoTotal += (count - cover.requirement) * cover.weightForOver
                solution.scoreInterno += (count - cover.requirement) * cover.weightForOver
            elif cover.requirement > count:
                solution.softViolations += 1
                totalPenalty += (cover.requirement - count) * cover.weightForUnder
                solution.requirementViolationsFaltan[dia2name[day]] += 1 #aqui
                CostoTotal += (cover.requirement - count) * cover.weightForUnder
                solution.scoreInterno += (cover.requirement - count) * cover.weightForUnder

    #Contar contratados trabajando
    if problem.ContratadosObligados:
        items = [item for item in problem.shifts.keys()]
        for day in range(problem.horizon):
            solution.ViolacionContratadosPorDia[dia2name[day]] = 0
            solution.NumeroContratados = 0
            for item in items:
                for staffId, schedule in solution.schedule.items():
                    if problem.staff[staffId].Contratado and schedule[day] == item:
                        solution.NumeroContratados += 1
                        break
            if solution.NumeroContratados < len(items):
                solution.softViolations += len(items) - solution.NumeroContratados
                CostoTotal += (len(items) - solution.NumeroContratados)*problem.hardConstraintWeight
                solution.scoreInterno += (len(items) - solution.NumeroContratados)*problem.hardConstraintWeight
                solution.ViolacionContratadosPorDia[dia2name[day]] += 1


    for shift, cover in problem.cover[0].items():
        solution.Requerimientos = cover.requirement
        break
    solution.IdTurnos = [item for item in problem.shifts.keys()]
    solution.scoreCuantificable = sum(solution.scoreCuantificableParcial.values())
    solution.score = CostoTotal
