t=2


######################################################
# Objetivo: Read the worksheet 1 e guarda em variaveis 
# Variaveis: Index, Class, T_ON, T_OFF, T_start
######################################################
import openpyxl

# load the workbook
workbook = openpyxl.load_workbook('Info.xlsx')

# select the worksheet
worksheet = workbook['Lds_inputs']

# read Index of cells
column = worksheet['B']
Index = []
for cell in column:
    if cell.row == 1:
        continue
    Index.append(cell.value)

# read Classification of cells
column = worksheet['C']
Class = []
for cell in column:
    if cell.row == 1:
        continue
    Class.append(cell.value)

# read T_ON of cells
column = worksheet['D']
T_ON = []
for cell in column:
    if cell.row == 1:
        continue
    T_ON.append(cell.value)

# read T_OFF of cells
column = worksheet['E']
T_OFF = []
for cell in column:
    if cell.row == 1:
        continue
    T_OFF.append(cell.value)

# read T_start of cells
column = worksheet['F']
T_start = []
for cell in column:
    if cell.row == 1:
        continue
    T_start.append(cell.value)

# read Power
column = worksheet['G']
Power = []
for cell in column:
    if cell.row == 1:
        continue
    Power.append(cell.value)

# print the values
#print(Index)
#print(Class)
#print(T_ON)
#print(T_OFF)
#print(T_start)
#print(Power)


############################################################################
# Objetivo: Cria matrix ON/OFF de acordo com as especificaçoes de worksheet1 
# Variaveis: ON_or_OFF
############################################################################

#Declaração de variaveis
cont_on = [0] * len(Index)
cont_off = [0] * len(Index)
estado = [0] * len(Index) #0=desligado, 1=ligado
rows = t
cols = len(Index)
ON_or_OFF = [['' for j in range(cols)] for i in range(rows)]
i=0
j=0
# 0=desligado, 1=ligado

for i in range (t):
    for j in range(len(Index)): 
        if (T_OFF[j] != 0): #Flexivel no tempo
            if (T_start[j] == 1 or estado[j] == 1): # ON 
                if cont_on[j] < T_ON[j]:
                    ON_or_OFF[i][j] = 1
                    cont_on[j] += 1
                    estado[j]=1
                else:
                    cont_on[j] = 0
                    ON_or_OFF[i][j] = 0
                    cont_off[j] = 1
                    estado[j]=0
            elif (T_start[j] == 0 or estado[j] == 0): # OFF 
                if cont_off [j] < T_OFF[j]:
                    ON_or_OFF[i][j] = 0
                    cont_off[j] += 1
                    estado[j]=0
                else:
                    cont_off[j] = 0
                    ON_or_OFF[i][j] = 1
                    cont_on[j] = 1
                    estado[j]=1
  

        elif (T_OFF[j] == 0 and (T_start[j] == 1 or estado[j] == 1)): # (Não flexivel ou flexivel em pot) e ON
            if cont_on[j] < T_ON[j]:
                ON_or_OFF[i][j] = 1
                cont_on[j] += 1
                estado[j]=1
            else:
                cont_on[j] = 0
                ON_or_OFF[i][j] = 0
                cont_off[j] = 1
                estado[j]=0

        elif (T_OFF[j] == 0 and (T_start[j] == 0 or estado[j] == 0) ): # (Não flexivel ou flexivel em pot) e OFF
            if ((j== 3 or j== 4 or j== 5) and (i==420 or i==1140 or i==1200 or i==1260)) or ((j== 9 or j== 12 or j== 13) and (i==780 or i==1200)): 
            # ligar luzes em 07:00 , 19:00 , 20:00 , 21:00; ligar load não flexivel as 13:00 ou 20
                ON_or_OFF[i][j] = 1
                cont_on[j] += 1
                estado[j] = 1
            else:
                ON_or_OFF[i][j] = 0
                estado[j] = 0
                cont_on[j] = 0
                cont_off[j] = 1

    j += 1
i += 1


#print(ON_or_OFF)


##############################################################################
# Objetivo: Modelo
##############################################################################

import pyomo.environ as pe
import np


def param_values():
    for i in range(1, 15):
        yield (i, Power[i])

# Declare the model
model = pe.ConcreteModel()

# Declare the large list parameter with the generator function
model.Power = pe.Param(range(1, 15), initialize=param_values())

# Access the values of the parameter
for i in range(1, 11):
    print(f"Value of parameter for index {i}: {model.Power[i].value}")




"""
# Create a model instance
model = pe.ConcreteModel()

# Parameters
model.index = pe.Set(initialize=Index,doc='Index')
model.Power[i] = pe.Param(initialize={Power[i]})
model.Ton = pe.Param(initialize=T_ON)
model.time = pe.Param(initialize=t)



# Variables are the things we want to solve for
model.q = pe.Var(domain=pe.NonNegativeReals, initialize=0.0)

# Create a constraint with a function -> constraints are things that hold true and cannot be violated
def _constraint_one(m):
    return m.q <= 10

# Assign the constraint to the model
model.constraint_one = pe.Constraint(rule=_constraint_one)

# Add a model objective -> objective is the thing we want to maximize or minimize
def _objective_function(m):
    return m.q * m.V * m.kA * m.CAf / (m.q + m.V * m.kB) / (m.q + m.V * m.kA)
model.objective = pe.Objective(expr=_objective_function, sense=pe.maximize)

# Create a solver instance -> Change this according to the problem. SCIP is a good general purpose solver
results = pe.SolverFactory('scip', executable='C:/Program Files/SCIPOptSuite 8.0.3/bin/scip.exe').solve(model)
model.pprint()


# print solutions
qmax = model.q()
CBmax = model.objective()
print('\nFlowrate at maximum CB = ', qmax, 'liters per minute.')
print('\nMaximum CB =', CBmax, 'moles per liter.')
print('\nProductivity = ', qmax*CBmax, 'moles per minute.')


"""
print('\nPower = ', model.Power[i], 'W')

"""

##############################################################################
# Objetivo: Acrescentar os minutos e index á matrix ON_OFF.
# Variaveis: ON/OFF.
##############################################################################

# Create a numerated vector using a for loop and the range function
col = []
for i in range(t):
    col.append(i + 1)
for i in range(t):
    ON_or_OFF[i].insert(0, col[i])
    

row = []
for i in range(len(Index)+1):
    row.append(i)
row [0] = 'Min/Index'

# Insert the new row at the beginning of the matrix
ON_or_OFF.insert(0, row)



############################################################################################################
# Objetivo: Imprime as 5 worksheets.
# Função: Escreve no excel a tabela das especificacoes, a tabela com ON/OFF para cada minuto para cada load,
#         tabela com pesos, dataset depois de optimizado e resultados.
############################################################################################################
import pandas as pd

# Create a Pandas Excel writer using XlsxWriter
writer = pd.ExcelWriter('Dataset.xlsx', engine='xlsxwriter')

#Dataset sem optimização
df2 = pd.DataFrame(ON_or_OFF)
df2.to_excel(writer, sheet_name='Dataset', index=False, header=False)

writer.save()


"""



"""

# EXEMPLO EDUARDO

# Create a model instance
model = pe.ConcreteModel()

# Create model parameters -> parameters are not variables, they are constants
model.V = pe.Param(initialize=V)
model.kA = pe.Param(initialize=kA)
model.kB = pe.Param(initialize=kB)
model.CAf = pe.Param(initialize=CAf)

# Variables are the things we want to solve for
model.q = pe.Var(domain=pe.NonNegativeReals, initialize=0.0)

# Create a constraint with a function -> constraints are things that hold true and cannot be violated
def _constraint_one(m):
    return m.q <= 10

# Assign the constraint to the model
model.constraint_one = pe.Constraint(rule=_constraint_one)

# Add a model objective -> objective is the thing we want to maximize or minimize
def _objective_function(m):
    return m.q * m.V * m.kA * m.CAf / (m.q + m.V * m.kB) / (m.q + m.V * m.kA)
model.objective = pe.Objective(expr=_objective_function, sense=pe.maximize)

# Create a solver instance -> Change this according to the problem. SCIP is a good general purpose solver
results = pe.SolverFactory('scip', executable='C:/Program Files/SCIPOptSuite 8.0.3/bin/scip.exe').solve(model)
model.pprint()


# print solutions
qmax = model.q()
CBmax = model.objective()
print('\nFlowrate at maximum CB = ', qmax, 'liters per minute.')
print('\nMaximum CB =', CBmax, 'moles per liter.')
print('\nProductivity = ', qmax*CBmax, 'moles per minute.')



"""



