import numpy as np
import matplotlib.pyplot as plt
import math
from numpy import log as ln
from openpyxl import load_workbook


x, y = np.meshgrid(np.arange(0, 2500, 0.01), np.arange(1, 5, 0.05))



wb = load_workbook("x_valuesBook.xlsx")
sheet = wb["Sheet1"]
read = "A"
write = "B"

#z = 4.6041 #Critical Value from Corresponding 2 sided t-test confidence interval
z = 2.7764
##VARIABLE FOR LANGMUIR

Smin = 1.46396
Smax = 154.38999
k = 329.93951

σSmin = 0.01138
σSmax = 4.82143
σk = 12.59838

##Variables for Logisitc

a = 1.46033
b = 176.46
c = 360.4
d = 0.864
e = 1.161

sigmaA = 0.0122
sigmaB = 6.775
sigmaC = 15.28
sigmaD = 0.0463
sigmaE = 0.07955

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

def Logistic(x, factor, type): #Function for Logistic Plotting of RMC
    BA = b - a
    x = x * factor
    val = c / x
    valC = val ** d  # (c/x)^d
    valU = (1 + valC) ** e  # (1+(c/x)^d)^e
    if type == 1:
        valU2 = 1 + valC
        valU2 **= (-e - 1)  # (1+(c/x)^d)^(-e-1)

        eqOne = 1 - (1/valU)
        eqOne *= sigmaA
        eqOne **= 2

        eqTwo = 1/valU
        eqTwo *= sigmaB
        eqTwo **= 2

        eqThree = (e*BA)*(d*valC)*(valU2)
        eqThree /= c
        eqThree *= -1
        eqThree *= sigmaC
        eqThree **= 2

        eqFour = (e*BA)*(ln(val))*(valU2)*(valC)
        eqFour *= -1
        eqFour *= sigmaD
        eqFour **= 2


        eqFive = (BA)*(ln((1+valC)))
        eqFive /= valU
        eqFive *= -1
        eqFive *= sigmaE
        eqFive **= 2

        return eqOne + eqTwo + eqThree + eqFour + eqFive

    if type == 2:
        S = 1 / valU
        S *= BA
        S += a
        return S




def getGraph():
    sumB = Langmuir(x, 1, 1) + Langmuir(x, y, 1) #Replace Langmuir with Logistic if using logistic (need to also do this again, scroll down)
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
        for μ in np.arange(1, 19, 0.0004): #if many values missing from excel sheet change from 0.0004 to 0.0001
            sumB = Langmuir(X, 1, 1) + Langmuir(X, μ, 1) #Replace Langmuir with Logistic if using logistic
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





