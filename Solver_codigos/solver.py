import random
import math
import time
import copy
import Solver_codigos.validator

class SolutionInstance:
    '''
    schedule = {staffId: list(shiftId)}
    '''
    def __init__(self):
        self.horizon = 0
        self.score = 0
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
        self.externalViolations = 0 #aqui
        self.HorasExtraSemanales = dict()
        self.HorasContratoSemanales = dict()
        self.TotalHorasTrabajadas = dict()
        self.schedule = dict()
        self.requirementViolationsFaltan = dict()
        self.requirementViolationsSobran = dict()
        self.CostoParcial = dict()
        self.NumeroContratados = 0     
        self.ViolacionContratadosPorDia = dict()
        self.DomingoTrabajado = dict()
        self.scoreCuantificableParcial = dict()
        self.scoreCuantificable = 0
        self.scoreInterno = 0
        self.Requerimientos = 0
        self.IdTurnos = list()

    def ShallowCopy(self):
        result = SolutionInstance()
        result.horizon = self.horizon
        result.score = self.score
        result.schedule = {x: y for x, y in self.schedule.items()}
        return result

    def PrintDebug(self):
        for staff, schedule in self.schedule.items():
            print('solution.schedule[\'{}\'] ='.format(staff), schedule)
        print("\n")

    def SaveDebug(self,f,solution):
        for staff, schedule in self.schedule.items():
            to_write = ('solution.schedule[\'{}\'] ='.format(staff), schedule)
            f.write(str(to_write))            
            f.write("\n")
        f.write('Score:'+str(solution.score))
        f.write("\n")
        f.write('Hard violations:'+str(solution.hardViolations))
        f.write("\n")
        f.write('Soft violations:'+str(solution.softViolations))
        f.write("\n")
        f.write('External violations:'+str(solution.externalViolations))
        f.write("\n")
        f.write('DaysOff violations:'+str(solution.daysOffViolations))
        f.write("\n")

    def Show(self):
        for schedule in self.schedule.values():
            print('\t'.join(schedule).replace(' ', ''))

def CreateEmptySolution(problem):
    result = SolutionInstance()
    result.horizon = problem.horizon

    for staffId in problem.staff.keys():
        result.schedule[staffId] = [' '] * problem.horizon

    return result

'''
Simulated Annealing has 4 major parts:
    1. A valid start configuration
    2. A random rearrangement scheme
    3. An objective function
    4. An annealing schedule
'''

def calcAvrMinutes(problem):
    return sum([x.length for k, x in problem.shifts.items()])/len(problem.shifts.items())
    
def calcDaysOff(problem, staff):
    avrMin = calcAvrMinutes(problem)
    maxMinStaff = max([x.maxTotalMinutes for k, x in problem.staff.items()])
    maxDaysOff = (problem.horizon * avrMin - maxMinStaff) / avrMin
    return maxDaysOff - len(problem.staff[staff].daysOff)

def GenerateInitialConfiguration(problem):
    result = CreateEmptySolution(problem)
    for key in result.schedule.keys():
        currentMinutes = 0
        staffMaxShifts = problem.staff[key].maxShifts
        impossible_shifts = [shift for shift, count in staffMaxShifts.items() if count == 0]
        avaliable_shifts = (set(problem.shifts.keys()) - set(impossible_shifts)).union(set([' ']))
        last_shift = ' '
        weekends = 0
        consecutiveOn = 0
        daysOff = 0
        maxDays = calcDaysOff(problem, key)
        for day in range(problem.horizon):
            if not day in problem.staff[key].daysOff:
                if last_shift in problem.shifts:
                    prohibitShifts = avaliable_shifts.intersection(problem.shifts[last_shift].prohibitNext)
                else:
                    prohibitShifts = set()
                
                avaliable_shifts -= prohibitShifts
                
                if consecutiveOn == problem.staff[key].maxConsecutiveShifts or \
                    ((day > 2 and ((day + 1) % 7 == 0 or (day + 2) % 7 == 0) \
                    and weekends == problem.staff[key].maxWeekends)):
                    curr_shift = ' '
                else:
                    curr_shift = random.choice(list(avaliable_shifts))
                
                if curr_shift == ' ':
                    consecutiveOn = 0
                else:
                    consecutiveOn += 1
    
                if(curr_shift != ' '):
                    currentMinutes += problem.shifts[curr_shift].length
                if currentMinutes > problem.staff[key].maxTotalMinutes :
                    break
                
                if curr_shift == ' ':
                    daysOff += 1
                
                result.schedule[key][day] = curr_shift;

                if (day > 2 and (day + 1) % 7 == 0) and \
                   (result.schedule[key][day] != ' ' or  result.schedule[key][day - 1] != ' '):
                        weekends += 1
                
                avaliable_shifts.union(prohibitShifts)
                last_shift = curr_shift
                if curr_shift in staffMaxShifts:
                    if staffMaxShifts[curr_shift] == 1:
                        avaliable_shifts -= set([curr_shift])
                    else:
                        staffMaxShifts[curr_shift] -= 1

    Solver_codigos.validator.CalculatePenalty(result, problem)
    return result

def NeighbourMove_TotalReorder(solution, **kw):
    staffId = random.choice(list(solution.schedule.keys()))
    schedule = solution.schedule[staffId]
    startIndex = [0]

    prevShift = schedule[0]
    currShift = ''

    # Find all bounds between different shifts
    for idx in range(1, solution.horizon):
        currShift = schedule[idx]
        if currShift != prevShift:
            startIndex.append(idx)
        prevShift = currShift

    reorderIndex = random.choice(startIndex)
    solution.schedule[staffId] = schedule[reorderIndex:] + schedule[:reorderIndex]

def NeighbourMove_PartialReorder(solution, **kw):
    staffId = random.choice(list(solution.schedule.keys()))
    schedule = solution.schedule[staffId]
    startIndex = [0]

    prevShift = schedule[0]
    currShift = ''

    # Find all bounds between different shifts
    for idx in range(1, solution.horizon):
        currShift = schedule[idx]
        if currShift != prevShift:
            startIndex.append(idx)
        prevShift = currShift

    # Edge case: whole schedule is only one shift
    if len(startIndex) == 1:
        return

    seq1, seq2 = 0, 0
    while seq1 == seq2:
        seq1 = random.randint(0, len(startIndex) - 1)
        seq2 = random.randint(0, len(startIndex) - 1)

    if seq1 > seq2:
        seq1, seq2 = seq2, seq1

    startIndex.append(solution.horizon)

    start1 = startIndex[seq1]
    end1 = startIndex[seq1 + 1]
    start2 = startIndex[seq2]
    end2 = startIndex[seq2 + 1]

    # [0, 1, 5, 7, 9, 11]
    # 4 1 => [1, 5), [9, 11)
    #
    # solution.schedule['A'] = [' ', <'D', 'D', 'D', 'D'>, ' ', ' ', 'D', 'D', <' ', ' '>, 'D', 'D', 'D']
    # solution.result  ['A'] = [' ', <' ', ' '>, ' ', ' ', 'D', 'D', <'D', 'D', 'D', 'D'>, 'D', 'D', 'D']

    solution.schedule[staffId] = \
        schedule[:start1] + \
        schedule[start2:end2] + \
        schedule[end1:start2] + \
        schedule[start1:end1] + \
        schedule[end2:]

def NeighbourMove_SegmentShift(solution, annealCoeff = 0.25, **kw):
    staffId = random.choice(list(solution.schedule.keys()))
    schedule = solution.schedule[staffId]

    segmentLength = max(int(len(schedule) * annealCoeff), 1)
    segmentStart = random.randint(0, len(schedule) - segmentLength)

    shiftDist = 0
    while shiftDist == 0:
        shiftDist = random.randint(-segmentStart, len(schedule) - segmentStart + segmentLength)

    if shiftDist < 0:
        solution.schedule[staffId] = \
            schedule[: segmentStart + shiftDist] + \
            schedule[segmentStart : segmentStart + segmentLength] + \
            schedule[segmentStart + shiftDist : segmentStart] + \
            schedule[segmentStart + segmentLength:]
    else:
        solution.schedule[staffId] = \
            schedule[: segmentStart] + \
            schedule[segmentStart + segmentLength : segmentStart + segmentLength + shiftDist] + \
            schedule[segmentStart : segmentStart + segmentLength] + \
            schedule[segmentStart + segmentLength + shiftDist :]

def NeighbourMove_SwitchShift(solution, **kw):
    staffId = random.choice(list(solution.schedule.keys()))
    day = random.randint(0, solution.horizon - 1)
    tmp_kw = list(kw['shiftTypes'])
    tmp_kw.remove(solution.schedule[staffId][day])
    newShift = random.choice(tmp_kw)
    solution.schedule[staffId][day] = newShift

def NeighbourMove_SwapShifts(solution, **kw):
    staffId = random.choice(list(solution.schedule.keys()))
    day1, day2 = 0, 0
    while day1 == day2:
        day1 = random.randint(0, solution.horizon - 1)
        day2 = random.randint(0, solution.horizon - 1)
    schedule = solution.schedule[staffId]
    schedule[day1], schedule[day2] = schedule[day2], schedule[day1]

#Agregado
def NeighbourMove_SwapStaffShifts(solution, **kw):
    trabajadores = list(solution.schedule.keys())
    if len(trabajadores)<2:
        return
    staffId1 = random.choice(trabajadores)
    trabajadores.remove(staffId1)
    staffId2 = random.choice(trabajadores)

    day = random.randint(0, solution.horizon - 1)
    schedule1 = solution.schedule[staffId1]
    schedule2 = solution.schedule[staffId2]
    schedule1[day], schedule2[day] = schedule2[day], schedule1[day]

def NeighbourMove_SwapEventualesShifts(solution, **kw):
    contratados = [worker for worker in list(solution.schedule.keys()) if 'Trabajador' in worker]
    eventuales = [worker for worker in list(solution.schedule.keys()) if 'Eventual' in worker]
    if len(contratados+eventuales)<2:
        return
    staffId1 = random.choice(contratados)
    staffId2 = random.choice(eventuales)

    non_empty_2 = [dia for dia,x in enumerate(solution.schedule[staffId2]) if x != ' ']
    if len(non_empty_2)>0:
        day = random.choice(non_empty_2)
    else:
        day = random.randint(0, solution.horizon - 1)
    schedule1 = solution.schedule[staffId1]
    schedule2 = solution.schedule[staffId2]
    schedule1[day], schedule2[day] = schedule2[day], schedule1[day]

# Moves with their relative weight
neighbourMoves = [
    [NeighbourMove_TotalReorder, 1],
    [NeighbourMove_PartialReorder, 3],
    [NeighbourMove_SegmentShift, 5],
    [NeighbourMove_SwitchShift, 55],
    [NeighbourMove_SwapStaffShifts, 15],
    [NeighbourMove_SwapShifts, 15]
]
# Moves with their relative weight
neighbourMoves2 = [
    # [NeighbourMove_TotalReorder, 1],
    # [NeighbourMove_PartialReorder, 3],
    # [NeighbourMove_SegmentShift, 5],
    # [NeighbourMove_SwitchShift, 55],
    # [NeighbourMove_SwapStaffShifts, 15],
    # [NeighbourMove_SwapShifts, 15],
    [NeighbourMove_SwapEventualesShifts, 15]
]

def MakeAccum(moves):
    totalWeight = sum([x[1] for x in neighbourMoves])
    accum = 0.0
    for idx, x in enumerate(neighbourMoves):
        accum += x[1] / totalWeight
        moves[idx][1] = accum
    moves[-1][1] = 1.0

def ChooseMove(moves):
    p = random.random()
    idx = 0
    while moves[idx][1] < p:
        idx += 1
    return moves[idx][0]

def FixDaysOff(solution, problem):
    for staffId, staffMember in problem.staff.items():
        schedule = solution.schedule[staffId]
        for idx in staffMember.daysOff:
            if schedule[idx] != ' ':
                otherIdx = idx
                while otherIdx in staffMember.daysOff:
                    otherIdx = random.randint(0, problem.horizon - 1)
                schedule[idx], schedule[otherIdx] = schedule[otherIdx], schedule[idx]

def FixSolution(solution, problem):
    FixDaysOff(solution, problem)

def AnnealingSchedule(mu):
    return mu * 1.05

def Anneal(problem = None, maxTime = float('inf'), runs = 1, useAnnealing = True,
             Temperature = 200, debug = None):
    '''
    Try to solve the given problem while not exceeding 'maxTime'.
    'runs' is the number of tries to solve the problem. Each try should start from
    a different random configuration.
    '''
    # Solution variables
    timePerInstance = maxTime / runs
    bestSolution = None
    bestValidSolution = None

    # Annealing variables
    mu = -1/Temperature
    totalIterations = 0
    accept = 0

    # Initialization
    MakeAccum(neighbourMoves)
    allShiftTypes = list(problem.shifts.keys())
    allShiftTypes.append(' ')

    graphData = [[] for _ in range(runs)]

    for r in range(runs):
        solution = GenerateInitialConfiguration(problem)
        # solution = CreateEmptySolution(problem)
        # validator.CalculatePenalty(solution, problem)

        if debug is not None:
            debug.write(('Starting run<{}> with score {}\n'.format(r, solution.score)))

        endTime = time.time() + timePerInstance

        if bestSolution is None:
            bestSolution = solution

        while True:
            if (time.time() > endTime):
                break

            if accept == 5000:
                mu = AnnealingSchedule(mu)
                accept = 0
                if debug is not None:
                    debug.write(("Temperatura: {0:f}\n".format(-1.0/mu)))

            if totalIterations % 2000 == 0:
                bestValid = 0
                if not (bestValidSolution is None):
                    bestValid = bestValidSolution.score
                graphData[r].append((totalIterations, solution.score, bestValid))

            totalIterations += 1

            newSolution = copy.deepcopy(solution)
            if (time.time() > endTime-timePerInstance*(1/16)):
                ChooseMove(neighbourMoves2)(newSolution, shiftTypes=allShiftTypes)
                useAnnealing = False
            else:
                ChooseMove(neighbourMoves)(newSolution, shiftTypes=allShiftTypes)
            FixSolution(newSolution, problem)
            Solver_codigos.validator.CalculatePenalty(newSolution, problem)

#            if solution.hardViolations == 0 and (bestValidSolution is None or bestValidSolution.score > solution.score): #aqui
            if solution.softViolations == 0 and (bestValidSolution is None or bestValidSolution.score > solution.score): #aqui
#            if (bestValidSolution is None or bestValidSolution.score > solution.score): #aqui
                    bestValidSolution = copy.deepcopy(solution)
                    if debug is not None:
                        debug.write(('Found better valid solution: {}\n'.format( bestValidSolution.score)))

            if newSolution.score <= solution.score or \
                    (useAnnealing and random.random() < math.exp(mu * (newSolution.score - solution.score))):
                solution = newSolution
                accept += 1
                if bestSolution.score > solution.score:
#                    print(solution.score)
                    bestSolution = copy.deepcopy(solution)
                    if debug is not None:
                        debug.write(('Found better solution: {}\n'.format(bestSolution.score)))
                        debug.write(("Total soft: {}\n".format(solution.softViolations)))
                        debug.write(("Total hard: {}\n".format(solution.hardViolations)))
            # print(useAnnealing,bestSolution.score,time.time(),endTime,endTime-timePerInstance*(1/16))            

    if debug is not None:
        debug.write(('Total iterations: {}\n'.format(totalIterations)))
    if bestValidSolution is None:
        return bestSolution, graphData
    else:
        return bestValidSolution, graphData