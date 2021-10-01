import Solver_codigos.instance
import sys

def ParseHorizon(line, thisInstance):
	# The horizon length in days:
	thisInstance.horizon = int(line)
	thisInstance.cover = [dict() for _ in range(thisInstance.horizon)]

def ParseShifts(line, thisInstance):
	# ShiftID, Length in mins, Shifts which cannot follow this shift | separated
	result = Solver_codigos.instance.Shift()
	sections = line.split(',')

	result.id = sections[0]
	result.length = int(sections[1])
	result.prohibitNext = set()

	for x in sections[2].split('|'):
		result.prohibitNext.add(x)

	thisInstance.shifts[result.id] = result

def ParseStaff(line, thisInstance):
	# ID, MaxShifts, MaxTotalMinutes, MinTotalMinutes, MaxConsecutiveShifts, MinConsecutiveShifts, MinConsecutiveDaysOff, MaxWeekends
	result = Solver_codigos.instance.StaffMember()
	sections = line.split(',')

	result.id = sections[0]
	result.maxShifts = dict()
	result.maxTotalMinutes = int(sections[2])
	result.minTotalMinutes = int(sections[3])
	result.maxConsecutiveShifts = int(sections[4])
	result.minConsecutiveShifts = int(sections[5])
	result.minConsecutiveDaysOff = int(sections[6])
	result.maxWeekends = int(sections[7])
	result.Costo = float(sections[8])
	if int(sections[9])==1:
		result.Contratado = True
	result.CostoExtra50 = float(sections[10])
	result.CostoExtra100 = float(sections[11])
	result.CostoExtra200 = float(sections[12])
	result.CostoExtraExterno = float(sections[13])

	for x in sections[1].split('|'):
		shiftId, maxCount = x.split('=')
		maxCount = int(maxCount)
		# Only add restriction that can be violated
		if (maxCount < thisInstance.horizon):
			result.maxShifts[shiftId] = maxCount
	
	thisInstance.staff[result.id] = result

def ParseDaysOff(line, thisInstance):
	# EmployeeID, DayIndexes (start at zero)
	sections = line.split(',')
	
	staffId = sections[0]
	days = [int(x) for x in sections[1:]]

	thisInstance.staff[staffId].daysOff = set(days)

def ParseShiftOnRequests(line, thisInstance):
	# EmployeeID, Day, ShiftID, Weight
	sections = line.split(',')
	result = Solver_codigos.instance.ShiftRequest()

	result.id = sections[2]
	result.day = int(sections[1])
	result.weight = int(sections[3])

	if thisInstance.hardConstraintWeight < result.weight:
		thisInstance.hardConstraintWeight = result.weight

	thisInstance.staff[sections[0]].shiftOnRequests[result.day] = result

def ParseShiftOffRequests(line, thisInstance):
	# EmployeeID, Day, ShiftID, Weight
	sections = line.split(',')
	result = Solver_codigos.instance.ShiftRequest()

	result.id = sections[2]
	result.day = int(sections[1])
	result.weight = int(sections[3])

	if thisInstance.hardConstraintWeight < result.weight:
		thisInstance.hardConstraintWeight = result.weight

	thisInstance.staff[sections[0]].shiftOffRequests[result.day] = result

def ParseCover(line, thisInstance):
	# Day, ShiftID, Requirement, Weight for under, Weight for over
	result = Solver_codigos.instance.Cover()
	sections = line.split(',')

	result.day = int(sections[0])
	result.shiftId = sections[1]
	result.requirement =  int(sections[2])
	result.weightForUnder = int(sections[3])
	result.weightForOver = int(sections[4])

	if thisInstance.hardConstraintWeight < result.weightForUnder:
		thisInstance.hardConstraintWeight = result.weightForUnder
	if thisInstance.hardConstraintWeight < result.weightForOver:
		thisInstance.hardConstraintWeight = result.weightForOver

	thisInstance.cover[result.day][result.shiftId] = result

def ParseCostos(line, thisInstance):

	result = Solver_codigos.instance.Costos()
	sections = [float(i) for i in line.split(',')]

	result.daysOffWeight = sections[0] #aqui
	result.prohibitNextWeight = sections[1]
	result.maxShiftsWeight = sections[2] #aqui
	result.minTotalMinutesWeight = sections[3] #aqui
	result.maxTotalMinutesWeight = sections[4] #aqui
	result.minConsecutiveShiftsWeight = sections[5] #aqui
	result.maxConsecutiveShiftsWeight = sections[6] #aqui
	result.minConsecutiveDaysOffWeight = sections[7] #aqui
	result.weekendsWeight = sections[8] #aqui
	result.externalWeight = sections[9] #aqui
#	result.requirementUnderWeight = sections[10] #aqui
#	result.requirementOverWeight = sections[11] #aqui

	thisInstance.costos = result

parseMethod = {
	'SECTION_HORIZON': ParseHorizon,
	'SECTION_SHIFTS': ParseShifts,
	'SECTION_STAFF': ParseStaff,
	'SECTION_DAYS_OFF': ParseDaysOff,
	'SECTION_SHIFT_ON_REQUESTS': ParseShiftOnRequests,
	'SECTION_SHIFT_OFF_REQUESTS': ParseShiftOffRequests,
	'SECTION_COVER': ParseCover,
	'SECTION_COSTOS': ParseCostos,
}

def LineType(line):
	if line in parseMethod.keys():
	   return line
	else:
		return 'DATA'

def ParseRoster(filename=None,contents_from_excel=None,MinutosPorSemana=None,
				contratados_obligatorios=False,MaximoHorasExtra=None,feriados=None,dia_inicio=None):
	if filename is not None:
		file = open(filename, 'r')
		contents = file.read()
		file.close()
	elif contents_from_excel is not None:
		contents = contents_from_excel
	else:
		sys.exit("Error:ParseRoster")

	# filter comments and empty lines in file
	contents = [s for s in contents.split('\n') if (s != '' and s[0] != '#')]
	parseType = None
	currentParseMethod = None
	result = Solver_codigos.instance.ProblemInstance()

	for line in contents:
		parseType = LineType(line)

		if parseType != 'DATA':
			currentParseMethod = parseMethod[parseType]
		else:
			currentParseMethod(line, result)

	#Used when using TotalPenalty
	result.hardConstraintWeight *= 1
	result.HorasPorSemana = MinutosPorSemana/60
	result.ContratadosObligados = contratados_obligatorios
	result.maxExtraHours = MaximoHorasExtra

	import datetime
	tmp_dia = dia_inicio
	for dia in range(7):
		if tmp_dia in feriados.values():
			result.feriados.append(tmp_dia.weekday())
		tmp_dia += datetime.timedelta(days=1)

	print("Feriados",result.feriados)
	return result
