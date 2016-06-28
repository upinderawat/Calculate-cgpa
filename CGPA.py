import sys
import re
import xlrd
import operator



def CGPA(t1,t2,t3,t4,t5,p1,p2,p3,p4):
    return round((4*(t1+t2+t3+t4+t5)+2*(p1+p2+p3+p4))/28,2)
def CalCG(num):
    '''takes number of subject and calculate the CG'''
    if num >= 90:
        return 10.0
    elif num >= 80.0:
        return 9.0
    elif num >= 70:
        return 8.0
    elif num >= 60:
        return 7.0
    elif num >= 50:
        return 6.0
    elif num>= 45:
        return 5.0
    elif num>= 40:
        return 4.0
    else: return 0.0             #the student fails if he scores less than 40 marks

def readPract(filename):
    '''extracts the number from 8th column of every row
    and multipy by 2. To get cgpa
    '''
    prac = dict()
    workbook = xlrd.open_workbook(filename)
    page1 = workbook.sheet_by_index(0)
    for i in range(1,page1.nrows):
        Prac=CalCG(page1.cell_value(i,7)*2)
        prac[page1.cell_value(i,0)]=Prac        #store the Practical(CG) in dict 'prac'

    return prac
def readTheory(filename):
    theory = dict()
    workbook = xlrd.open_workbook(filename)
    page1 = workbook.sheet_by_index(0)
    for i in range(1,page1.nrows):
        TCG=CalCG(page1.cell_value(i,1)+page1.cell_value(i,3)+page1.cell_value(i,4)+page1.cell_value(i,6))
        theory[page1.cell_value(i,0)]=TCG       # store the theory(CG) in dict 'theory'
    return theory

def main():
    rank =1
    result={}
    osp=readPract("OS.xls")
    sep=readPract("SEp.xlsx")
    adap=readPract("ada.xlsx")
    wtp=readPract("WT.xlsx")

    os=readTheory("OStheory.xls")
    cao=readTheory("Cao.xlsx")
    wt=readTheory("WT.xlsx")
    se=readTheory("SE.xlsx")
    ada=readTheory("ada.xlsx")

    file = open("record.txt","w")
    for roll in osp.keys():
        result[roll]=   CGPA(os[roll],cao[roll],wt[roll],se[roll],ada[roll],osp[roll],sep[roll],adap[roll],wtp[roll])
    result = sorted(result.items(),key=operator.itemgetter(1),reverse=True)
    for tup in result:
        file.write(str(rank) + " " + str(tup[0]) +" " + str(tup[1])+ "\n")
        rank= rank+1
    file.close()

if __name__ == "__main__":
    main()
