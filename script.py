#!/usr/bin/env python3
import sys
import os
import matplotlib.pyplot as plt
import pandas as pd
import numpy
from openpyxl import load_workbook

# These are global Variables - not a good practice but needed for now
# It will be changed at a later date when the concept of classes are introduced
a_MAX_by_g = 0.06
WT_Depth = 1.90
BAR_in_MPa = 0.1
BAR_in_KPa = 100
P_in_Area_KPa = 9.81


def main():
    '''Excel files quite often have multiple sheets and the ability
    to read a specific sheet or all of them is very important.
    To make this easy, the pandas ExcelFile class can be used
    to read multiple sheets from thr SAME excel file and pass it to the
    read_excel method'''
    run()
# end of main


def promptFilePath():
    print("Enter an absolute filepath: ")
    appended_file = input("Absolute File Path: ")
    # check if the given path exists
    _isPathValid = checkPath(appended_file)
    if _isPathValid:
        print("File is OK")
        return appended_file
    else:
        print("ERROR: Path does not exist. Exiting")
        exit()


def checkPath(excel_fileName):
    return os.path.exists(excel_fileName)


def run():
    appended_file = promptFilePath()
    # The getValue function is useful if you want to extract data
    # for one of the columns in your CPTu data
    depth_m = getValList(appended_file, "A")  # from the column A

    # but what if you want to calculate a bunch of stuff?
    # and also just enter the answer itself in the SAME file?

    # append to the EXISTING excel file
    '''
    So this part could have been a part of unit testing. 
    I know the current issues include unused parameters - that can
    potentially be solved by function overloading concept.
    '''
    appendToExcel(appended_file, "depth-diff_m", "0")
    print("depth-diff_m calculated")

    appendToExcel(appended_file, "U-KPa", "U-MPa")
    print("U-KPa calculated")

    appendToExcel(appended_file, "qc-MPa", "qc-Bar")
    print("qc-MPa calculated")

    appendToExcel(appended_file, "fs-KPa", "fs-Bar")
    print("fs-KPa is calculated")

    appendToExcel(appended_file, "qt-Bar", "qc-MPa")
    print("qt-Bar calculated")

    appendToExcel(appended_file, "Rf-Pct", "fs-Bar_qt-Bar")
    print("Rf-Pct calculated")

    appendToExcel(appended_file, "γ_kN_over_m3", "rf-Pct_qt-Bar")
    print("γ_kN_over_m3 calculated")

    appendToExcel(appended_file, "Depth-M", "depth")
    print("Depth-M added")

    appendToExcel(appended_file, "σv_KPa", "depth-diff")
    print("σv_KPa calculated")

    appendToExcel(appended_file, "effective_σv_KPa", "depth-diff")
    print("effective_σv_KPa calculated")

    appendToExcel(appended_file, "Qt_norm", "qt-Bar_σv-KPa")
    print("Qt_norm calculated")

    appendToExcel(appended_file, "Fr-norm", "Qt_norm")
    print("Fr-norm calculated")

    appendToExcel(appended_file, "Rd", "0")
    print("Rd calculated")

    appendToExcel(appended_file, "CSR", "0")
    print("CSR calculated")

    appendToExcel(appended_file, "normalization-factor", "0")
    print("normalization-factor calculated")

    appendToExcel(appended_file, "conestress-normalization", "o")
    print("conestress-normalization calculated")

    appendToExcel(appended_file, "CRR", "o")
    print("CRR calculated")

    appendToExcel(appended_file, "phi", "0")
    print("phi calculated")

    appendToExcel(appended_file, "FOS", "0")
    print("FOS calculated")

    appendToExcel(appended_file, "ICVal", "0")
    print("ICVal calculated")

    plot(appended_file)


def plot(appended_file):
    df = pd.read_excel(appended_file, index_col=0)
    df.plot.scatter(y='Depth-M', x='FOS', legend=True, rot=90)
    plt.show()


def countRows(dataframe):
    return dataframe.shape[0]


def appendToExcel(appended_file, newParam, dataUsed):

    # read the file as a DataFrame object
    df = pd.read_excel(appended_file, index_col=0)
    # -1 because excel is 1-indexed but Python is 0-indexed
    rowTotal = countRows(df) - 1

    if newParam == "depth-diff_m":
        df["depth-diff_m"] = calculateDepthDiff(appended_file, df, rowTotal)
    elif newParam == "U-KPa":
        df[newParam] = calculateUKPa(df)
    elif newParam == "qc-MPa":
        df[newParam] = calculate_qc_Mpa(df)
    elif newParam == "fs-KPa":
        df[newParam] = calculateFs_KPa(df)
    elif newParam == "qt-Bar":
        df[newParam] = df[dataUsed] * 10
    elif newParam == "Rf-Pct":
        df[newParam] = calculateRF(df)
    elif newParam == "γ_kN_over_m3":
        df[newParam] = calculateGamma(df)
    elif newParam == "Depth-M":
        df[newParam] = addDepth(appended_file, rowTotal, df)
    elif newParam == "σv_KPa":
        df[newParam] = calculate_sigmav(df, rowTotal)
    elif newParam == "effective_σv_KPa":
        df[newParam] = calculate_effective_sigmav(appended_file, df, rowTotal)
    elif newParam == "Qt_norm":
        df[newParam] = calculate_QtNorm(df)
    elif newParam == "Fr-norm":
        df[newParam] = calculate_FrNorm(df)
    elif newParam == "Rd":
        df[newParam] = calculateRd(appended_file, df, rowTotal)
    elif newParam == "CSR":
        df[newParam] = calculateCSR(df)
    elif newParam == "normalization-factor":
        df[newParam] = calculateNormalizationFactor(df)
    elif newParam == "conestress-normalization":
        df[newParam] = calculateConestressNormalization(df)
    elif newParam == "CRR":
        df[newParam] = calculateCRR(df)
    elif newParam == "phi":
        df[newParam] = calculatePhi(df)
    elif newParam == "FOS":
        df[newParam] = calculateFOS(df)
    elif newParam == "ICVal":
        df[newParam] = calculateICVal(df)
    df.to_excel("/Users/zmoin/Python Workshop/sampleData.xlsx")


def addDepth(fileName, totalRow, dataframe):
    list_of_dept_tmp = getValList(fileName, "A")
    count = 0
    while count < totalRow:
        dataframe.at[dataframe.index[count],
                     'Depth-M'] = list_of_dept_tmp[count]
        count = count + 1
    return dataframe['Depth-M']


def calculateFs_KPa(dataframe):
    return numpy.round(dataframe["fs-Bar"] * 100, decimals=2)


def calculateUKPa(dataframe):
    return numpy.round(dataframe["u2-M"] * 9.81, decimals=2)


def calculateICVal(dataframe):
    return numpy.round(((3.47-numpy.log10(dataframe["Qt_norm"]))**2 + (numpy.log10(dataframe["Fr-norm"]) + 1.22) ** 2)**0.5, decimals=2)


def calculatePhi(dataframe):
    return numpy.round(numpy.degrees(numpy.arctan((1/2.68)*(numpy.log10((dataframe["qc-MPa"]*1000)/dataframe["σv_KPa"])+0.29))), decimals=2)


def calculateFOS(dataframe):  # does not handle division by 0 but can be added
    return numpy.round(dataframe["CRR"]/dataframe["CSR"], decimals=2)


def calculateCRR(dataframe):
    return numpy.round(0.833 * (dataframe["conestress-normalization"]/100) + 0.05, decimals=2)


def calculateConestressNormalization(dataframe):
    return numpy.round(dataframe["normalization-factor"]*dataframe["qc-MPa"], decimals=2)


def calculateNormalizationFactor(dataframe):
    return numpy.round(((100/dataframe["σv_KPa"])**0.5), decimals=2)


def calculate_qc_Mpa(dataframe):
    return numpy.round(dataframe["qc-Bar"] * BAR_in_MPa, decimals=2)


def calculateCSR(dataframe):
    return numpy.round(0.65*dataframe["Rd"] * (dataframe["σv_KPa"]/dataframe["effective_σv_KPa"]) * a_MAX_by_g, decimals=2)


def calculateRd(fileName, dataframe, totalRows):
    list_of_depth = getValList(fileName, "A")
    count = 0
    while count < totalRows:
        dataframe.at[dataframe.index[count], 'Rd'] = numpy.round(
            1-(0.012 * list_of_depth[count]), decimals=2)
        count = count + 1
    return dataframe["Rd"]


def calculate_FrNorm(dataframe):
    return numpy.round((dataframe["fs-Bar"]*100)/(dataframe["qt-Bar"]*100 - dataframe["σv_KPa"])*100, decimals=2)


def calculate_QtNorm(dataframe):
    return numpy.round((dataframe["qt-Bar"] * 100 - dataframe["σv_KPa"])/dataframe["σv_KPa"], decimals=2)


def calculateGamma(dataframe):
    return numpy.round(9.81*(0.27*numpy.log10(dataframe["Rf-Pct"])+0.36*numpy.log10(dataframe["qt-Bar"]) + 1.236), decimals=2)


def calculateRF(dataframe):
    return numpy.round((dataframe["fs-Bar"]/dataframe["qt-Bar"]) * 100, decimals=2)


def calculateDepthDiff(fileName, dataframe, totalRows):
    list_of_depth = getValList(fileName, "A")
    dataframe.at[dataframe.index[0],
                 'depth-diff_m'] = getValList(fileName, "A")[0]
    count = 1
    while count < totalRows:
        dataframe.at[dataframe.index[count],
                     'depth-diff_m'] = list_of_depth[count + 1] - list_of_depth[count]
        count = count + 1
    return dataframe["depth-diff_m"]


def calculate_effective_sigmav(fileName, dataframe, totalRows):
    list_of_depth_m = getValList(fileName, "A")
    count = 0
    while count < totalRows:
        if dataframe.at[dataframe.index[count], 'u2-M'] < 0:
            dataframe.at[dataframe.index[count], 'effective_σv_KPa'] = numpy.round(
                dataframe.at[dataframe.index[count], 'σv_KPa'], decimals=2)
            count = count + 1
        else:
            dataframe.at[dataframe.index[count], 'effective_σv_KPa'] = numpy.round(
                dataframe.at[dataframe.index[count], 'σv_KPa'] - (9.81*(list_of_depth_m[count] - WT_Depth)), decimals=2)
            count = count + 1
    return dataframe["effective_σv_KPa"]


def calculate_sigmav(dataframe, totalRows):
    count = 1
    dataframe.at[dataframe.index[0], 'σv_KPa'] = numpy.round(
        dataframe.at[dataframe.index[0], 'depth-diff_m'] * dataframe.at[dataframe.index[0], 'γ_kN_over_m3'], decimals=2)
    while count < totalRows:
        dataframe.at[dataframe.index[count], 'σv_KPa'] = numpy.round(
            dataframe.at[dataframe.index[count], 'depth-diff_m'] * dataframe.at[dataframe.index[count], 'γ_kN_over_m3'] + dataframe.at[dataframe.index[count - 1], 'σv_KPa'], decimals=2)
        count = count + 1
    return dataframe["σv_KPa"]


# get value from the passed column
def getValList(excelFile, colNum):
    retVal = pd.read_excel(excelFile, usecols=colNum).values.T[0].tolist()
    return retVal


if __name__ == "__main__":
    main()
