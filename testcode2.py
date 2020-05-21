import numpy as np
import xlrd
import math
import xlwt

################ Read from excel ##################
def readExcel (filename):                       # Defines a function that reads the excel file.
    wb = xlrd.open_workbook(filename)
    sheet = wb.sheet_by_index(0)
    return wb, sheet                            # Returns the workbook and the sheet parameter.

loc = ("C:/Users/LarsMjeldheim/OneDrive/Documents/NC State/ISE553/CaseStudy2/BSL3.xlsx")
wb, sheet = readExcel(loc)                      # Returns the workbook and the sheet parameter.

#############  Critical fractile z-score ##################
zCF = sheet.cell_value(1,19)                    # Reads the Z score of the CF = 0.99

################ Holding Cost ######################
h = {}
HoldingWH = 1.0
for i in range(20):
    h[i] = sheet.cell_value(i+3,16)             # Reads the value from the sheet (as assigned earlier)
h[20] = HoldingWH                               # Needs to make the central WH node 21 (20, since Py is zeroed based)
h[21] = sheet.cell_value(20+3,16)               # Needs to make districs 21 at node 22

#################### Sigma ########################
sig = {}
for i in range(20):
    sig[i] = sheet.cell_value(i+3,20)           # Imports the sigma's from the excel
sig[20] = 0                                     # Empty for now. Need to calc the pooled sigma's
sig[21]= sheet.cell_value(20+3,20)

var = 0
for i in range(22):                             # A for loop to compute the sigma for the central WH
    var += sig[i]*sig[i]
SigWH = math.sqrt(var)
sig[20] = SigWH                                 # Assigns the computed sigma to the central WH stage

################ Lead time ######################
L = {}
for i in range(20):
    L[i] = sheet.cell_value(i+3,18)
L[20] = sheet.cell_value(21+3,18)
L[21] = sheet.cell_value(20+3,18)

################# S #############################
S = {}
for i in range(22):
    S[i] = 0                                    # Assigns all S(i) to be 0

################# SI ############################
T = {}
for i in range(22):
    T[i] = 7                                    # Assigns all the processing times to be equal to the periodic review time of 7

################## M ############################
M = {}
for i in range(22):
    M[i] = L[i] + L[20]                         # Assigns all the maximum CST of the system at stage i
    M[i] = int(M[i])
M[20] = int(L[20])

################ A (directed arcs) ###############
A = {}                                          # Assigns all the Arcs for each stage i
A[0] = (20,0)
A[1] = (20,1)
A[2] = (20,2)
A[3] = (20,3)
A[4] = (20,4)
A[5] = (20,5)
A[6] = (20,6)
A[7] = (20,7)
A[8] = (20,8)
A[9] = (20,9)
A[10] = (20,10)
A[11] = (20,11)
A[12] = (20,12)
A[13] = (20,13)
A[14] = (20,14)
A[15] = (20,15)
A[16] = (20,16)
A[17] = (20,17)
A[18] = (20,18)
A[19] = (20,19)
A[20] = (20,21)                 # Only assings the arc above i as the arcs to i's below are already represented
A[21] = (20,21)                 # Assigns an arc already assigned, however needed to compute the cost at the last node


############## Cost Funcstion ###################
def costK(S, SI, k):            # Defines the cost function as shown in Equation 8
    costk = (h[k]*zCF*sig[k]*math.sqrt(SI+T[k]-S)) + sumOut(SI,k) + sumInn(S,k)
    return costk

############### Sum of Outbound ####################
def sumOut(SI,k):                               # Defines the sum of the cost for Outbound arcs at stage k
    costSum = 0
    for i in range(k):                          # This checks all arcs below stage k
        if A[i][1] == k:                        # If there is any arc with stage k in the second possition (i,k), continue
            if A[i][0]<k:                       # If there is any arc with first position i less then k in a (i,k) arc, continue
                costMin = float("inf")
                for s in range(SI+1):           # Find the minimum S for the outbound theta at stage k
                    cost = thetaO(s,i)          # Calls the thetaO function for all s in the range SI
                    if cost < costMin:          # Checks if it is minimum
                        costMin = cost          # If its a minimum, update the costMin here.
                costSum += costMin              # This terms sums up all the Cost for the minimum cost computed
    return costSum                              # Return the sum of the min cost for each arc. Returns 0 if there is no such arcs

################ Sum of Innbound ####################
def sumInn(S,k):                                # Defines the sum of the cost for Innbound arcs at stage k.
    costSum = 0
    for j in range(k):                          # This checks all arcs below stage k
        if A[j][0] == k:                        # If there is any arc with stage k in the first possition (k,i), continue
            if A[j][1]<k:                       # If there is any arc with second position i less then k in a (k,i) arc, continue
                costMin = float("inf")
                for si in range(S,M[k]-T[k] + 1, 1):    # Find the minimum S for the innbound theta at stage k
                    cost = thetaI(si,j)         # Calls the thataI function for all si in the possible range.
                    if cost < costMin:          # Checks if it is minimum
                        costMin = cost          # If it is a minimum, update the 'costMin' function
                costSum += costMin              # This terms sums up all the Cost for the minimum cost computed
    return costSum                              # Return the sum of the min cost for each arc. Returns 0 if there is no such arcs

############## Theta Outbound ###################
def thetaO(S,k):                                # Defines the outbound theta fuction, as in Equation 8
    costMin = float("inf")
    for i in range(M[k]-T[k]+1):                # Checks all possible SI values
        if i + T[k] >= S:                       # This needs to hold true, unless invalid SI value.
            cost = costK(S,i,k)        # HERE IS THE RECURSIVE CALL. Calculates for k's below the k initiating this term in the cost function
            if cost < costMin:                  # Checks if the cost for a given si is minimum
                costMin = cost                  # Updates the min cost
    return costMin                              # Returns the minimum cost at stage k (the k provided into this defenition)

############## Theta Innbound ####################
def thetaI(SI,k):                               # Defines the outbound theta fuction, as in Equation 8
    costMin = float("inf")
    for i in range(S[k]+1):                     # Checks all possible S values
        if i <= SI + T[k]:                      # This needs to hold true, unless invalid S value.
            cost = costK(i,SI,k)       # HERE IS THE RECURSIVE CALL. Calculates for k's below the k initiating this term in the cost function
            if cost < costMin:                  # Checks if the cost for a given si is minimum
                costMin = cost                  # Updates the min cost
    return costMin                              # Returns the minimum cost at stage k (the k provided into this defenition)

################ Final Result ############

minFinalCost = float("inf")                     # Gives an infinite large number, to then update with minimum later
minFinalSI = 0
for si in range(M[21]-T[21]+1):                 # Checks for all possible SI values at the last stage N=22
    finalc = costK(S[21], si, 21)               # Cals the costK funstion, which is going to recusrivly calulate the cost fro all stages.
    if finalc < minFinalCost:                   # Checks if the SI provides a minimum
        minFinalCost = finalc                   # If minimum, update the minFinalCost function
        minFinalSI = si                         # If minimum, stores the SI value that gave the minimum
print('The min cost at N stage is', minFinalCost,'and the SI =', minFinalSI)

########## Final Result at each stage ######
FinalCost = {}                                  # Defines an open dictionary to store values later
minSI = {}                                      # Defines an open dictionary to store values later
SS = {}                                         # Defines an open dictionary to store values later
for k in range(22):                             # Does the same as above, but for all stages.
    FinalCost[k] = float("inf")
    for si in range(M[k]-T[k]+1):
        finalc = costK(S[k], si, k)
        if finalc < FinalCost[k]:
            FinalCost[k] = finalc               # Updates the FinalCost at stage k if it is a minimum
            minSI[k] = si                       # Updates the minSI at stage k if it is a minimum
            SS[k] = zCF*sig[k]*math.sqrt(si+T[k]-S[k])      # Updateds the Safety Stock in stage k if there is a minimum

print('The min cost at each stage', FinalCost)
print('The min SI at each stage', minSI)
print('Safty stock is', SS)

############## Output the result into Excel ##########
workbook = xlwt.Workbook()
resultXlS = workbook.add_sheet("Result")        # Creates a workbook

resultXlS.write(0, 0, 'Stage')
resultXlS.write(0, 1, 'District')
resultXlS.write(0, 2, 'Safety stock')
resultXlS.write(0, 3, 'Safety stock cost')
resultXlS.write(0, 4, 'SI value')
for i in range(22):
    resultXlS.write(i+1, 0, i)
    resultXlS.write(i+1, 2, SS[i])
    resultXlS.write(i+1, 3, FinalCost[i])
    resultXlS.write(i+1, 4, minSI[i])
for i in range(20):
    resultXlS.write(i+1, 1, sheet.cell_value(i+3,0))
resultXlS.write(20+1, 1, 'Central WH')
resultXlS.write(21+1, 1, sheet.cell_value(20+3,0))
workbook.save("Result.xls")
