#Gupta, R.; Gupta, P.; Wang, S.; Melnykov, A.; Jiang, Q.; Seth, A.; Wang, Z.; Morrissey, J.; George, I.; Gandra, S.; Sinha, P.; Storch, G.; Parikh, B.; Genin, G.; Singamanenei, S.; 
#Ultrasensitive lateral-flow assays via plasmonically active antibody-conjugated fluorescent nanoparticles. 
#Nature Biomedical Engineering
#DOI: https://www.nature.com/articles/s41551-022-01001-1


import numpy as np
import matplotlib.pyplot as plt
import math
from numpy import log as ln
from openpyxl import load_workbook

####IMPORTANT NOTES#####
# This code works best when you set your concentration to a unit such that the boundaries
# ranges from 0.0001 to 10000
# There may be instabilities in the nonlinear calculation if you do not change ur units to such range
# For Example: 
#     My assay is from 100 µg/mL to 1000000 µg/mL, I will adjust to 
#     0.1 mg/mL to 1000 mg/mL

#The following line defines the x axis range (currently 0 to 2000 units with a step size of 0.01. You may
#change this value depending on the wideness of your curve.

#The 1 to 5 range represents the y axis, which is the µ value. This should not be changed above 10.

########See notes above to change#######
x, y = np.meshgrid(np.arange(0, 2000, 0.01), np.arange(1, 5, 0.05))

########################################
#Critical Value from Corresponding 2 sided t-test confidence interval: currently 99%. 
z = 4.6041 
#Change this value to another z score if you want a different confidence interval. 

################################
####VARIABLES FROM LANGMUIR#####
################################
#Replace the following values with values obtained from curve fitting
#in Prism or OriginPro using a Langmuir binding isotherm

Smin = 1.46396
Smax = 154.38999
k = 329.93951

σSmin = 0.01138
σSmax = 4.82143
σk = 12.59838

################################
################################
################################

wb = load_workbook("x_valuesBook.xlsx") ###DO NOT CHANGE. MAKE SURE THIS XLSX FILE IS IN THE SAME DIRECTORY AS THIS FILE. 
sheet = wb["Sheet1"]
read = "A"
write = "B"

def Langmuir(x, factor, type): #Function for Langmuir Plotting of RMC
    x = x * factor
    val1 = x + k

    if type == 1:

        eqOne = x/val1
        eqOne *= σSmax
        eqOne **= 2

        eqTwo = k/val1
        eqTwo *= σSmin
        eqTwo **= 2

        eqThree = (Smin - Smax)*x
        eqThree /= val1**2
        eqThree *= σk
        eqThree **= 2

        return eqOne + eqTwo + eqThree

    if type == 2:
        S = (Smax - Smin)*x
        S /= val1
        S += Smin

        return S

    
def getGraph():
    sumB = Langmuir(x, 1, 1) + Langmuir(x, y, 1) 
    sumB = np.sqrt(sumB)
    sumT = Langmuir(x, y, 2) - Langmuir(x, 1, 2)
    equation = sumT/sumB
    equation -= target
    g = plt.contour(x, y, equation, [0])

    return g


if __name__ == "__main__":
    print("Nothing working?: follow this link to Desmos Graphing Calculator:\nhttps://www.desmos.com/calculator/gzl4irdyjx")
    print("Check Desmos calculator to see if graph is correct in the first place")
    print("Running, Should take 30-60 sec")

    target = round(z / (1.7320508), 4)
    graph = getGraph()
    p = graph.collections[0].get_paths()[0]
    v = p.vertices
    x_coords = np.around(v[:, 0], decimals=4)
    y_coords = np.around(v[:, 1], decimals=4)

    for i in range(1, len(sheet[write]) + 1, 1):
        sheet[write+str(i)] = None
        wb.save("x_valuesBook.xlsx")

    for i in range(1, len(sheet[read]) + 1, 1):
        X = sheet[read+str(i)].value
        for μ in np.arange(1, 19, 0.0004): #if many values missing from excel sheet change from 0.0004 to 0.0001, but this may crash older computers.
            sumB = Langmuir(X, 1, 1) + Langmuir(X, μ, 1) 
            sumB = np.sqrt(sumB)

            sumT = Langmuir(X, μ, 2) - Langmuir(X, 1, 2)
            sumT = math.fabs(sumT)
            calc = round(sumT/sumB, 4)
            if target + 0.0008 >= calc >= target - 0.0008:
                answer = μ
                sheet[write+str(i)] = answer
                wb.save("x_valuesBook.xlsx")
                break
    wb.save("x_valuesBook.xlsx")

    min_index = y_coords.argmin()
    min_value = str(round(y_coords[min_index], 4))
    min_pos = str(round(x_coords[min_index], 4))
    print("μ min = " + min_value + " at x = " + min_pos)

    try:
        indexs5 = np.where(y_coords == 5)
        position51 = str(round(x_coords[indexs5[0][0]], 4))
        position52 = str(round(x_coords[indexs5[0][1]], 4))
        print("μ is under 5 from " + position51 + " to " + position52)
    except IndexError:
        print("range not found: μ may not be under 5")
        pass
    try:
        indexs = np.where(y_coords == 2)
        actualIndex = []
        previous = -100
        for i in indexs[0]:
            if round(previous, 0) != round(x_coords[i], 0):
                actualIndex.append(x_coords[i])
            previous = x_coords[i]

        position1 = str(round(actualIndex[0], 4))
        position2 = str(round(actualIndex[1], 4))
        print("μ is under 2 from " + position1 + " to " + position2)

    except IndexError:
        print("range not found: μ may not be under 2")
        pass




