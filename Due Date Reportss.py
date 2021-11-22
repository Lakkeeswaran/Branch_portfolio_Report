import xlrd
file_location = "C:/Users/LAKKE/Documents/LCC Analysis/CHYUR/Deep Report/LCCss.xlsx"
file_location1 = "C:/Users/LAKKE/Documents/LCC Analysis/CHYUR/personal loan cases.xlsx"
workbook = xlrd.open_workbook(file_location)
workbook1 = xlrd.open_workbook(file_location1)
sheet = workbook.sheet_by_index(0)
sheetpln = workbook1.sheet_by_index(0)
import xlwt
wb = xlwt.Workbook()
ws = wb.add_sheet("Reports")
wtbs = wb.add_sheet("<1000 To be Strike")
wtbn = wb.add_sheet("<1000 To be Nill Arrear")
psnln = wb.add_sheet("Personal Loan")
above3 = wb.add_sheet("Above 3")
CNP = wb.add_sheet("CNP")
dd5th = wb.add_sheet("5th Due Date")
dd10th = wb.add_sheet("10th Due Date")
dd15th = wb.add_sheet("15th Due Date")
dd20th = wb.add_sheet("20th Due Date")
wtbs.write(0,0,"LOAN NO")
wtbs.write(0,1,"RE NAME")
wtbs.write(0,2,"Due Date")
wtbs.write(0,3,"Veh No")
wtbs.write(0,4,"Customer Name")
wtbs.write(0,5,"Description")
wtbs.write(0,6,"Demand")
wtbs.write(0,7,"Receipt")
wtbs.write(0,8,"Strike")
psnln.write(0,0,"LOAN NO")
psnln.write(0,1,"RE NAME")
psnln.write(0,2,"Due Date")
psnln.write(0,3,"Veh No")
psnln.write(0,4,"Customer Name")
psnln.write(0,5,"Description")
psnln.write(0,6,"Demand")
psnln.write(0,7,"Receipt")
psnln.write(0,8,"Closing Arrear")
dd5th.write(0,0,"LOAN NO")
dd5th.write(0,1,"RE NAME")
dd5th.write(0,2,"Due Date")
dd5th.write(0,3,"Veh No")
dd5th.write(0,4,"Customer Name")
dd5th.write(0,5,"Description")
dd5th.write(0,6,"Demand")
dd5th.write(0,7,"Receipt")
dd5th.write(0,8,"Closing Arrear")
dd5th.write(0,9,"Arrear EMI")
dd5th.write(0,10,"EMI Accured")
dd10th.write(0,0,"LOAN NO")
dd10th.write(0,1,"RE NAME")
dd10th.write(0,2,"Due Date")
dd10th.write(0,3,"Veh No")
dd10th.write(0,4,"Customer Name")
dd10th.write(0,5,"Description")
dd10th.write(0,6,"Demand")
dd10th.write(0,7,"Receipt")
dd10th.write(0,8,"Closing Arrear")
dd10th.write(0,9,"Arrear EMI")
dd10th.write(0,10,"EMI Accured")
dd15th.write(0,0,"LOAN NO")
dd15th.write(0,1,"RE NAME")
dd15th.write(0,2,"Due Date")
dd15th.write(0,3,"Veh No")
dd15th.write(0,4,"Customer Name")
dd15th.write(0,5,"Description")
dd15th.write(0,6,"Demand")
dd15th.write(0,7,"Receipt")
dd15th.write(0,8,"Closing Arrear")
dd15th.write(0,9,"Arrear EMI")
dd15th.write(0,10,"EMI Accured")
dd20th.write(0,0,"LOAN NO")
dd20th.write(0,1,"RE NAME")
dd20th.write(0,2,"Due Date")
dd20th.write(0,3,"Veh No")
dd20th.write(0,4,"Customer Name")
dd20th.write(0,5,"Description")
dd20th.write(0,6,"Demand")
dd20th.write(0,7,"Receipt")
dd20th.write(0,8,"Closing Arrear")
dd20th.write(0,9,"Arrear EMI")
dd20th.write(0,10,"EMI Accured")
wtbn.write(0,0,"LOAN NO")
wtbn.write(0,1,"RE NAME")
wtbn.write(0,2,"Due Date")
wtbn.write(0,3,"Veh No")
wtbn.write(0,4,"Customer Name")
wtbn.write(0,5,"Description")
wtbn.write(0,6,"Demand")
wtbn.write(0,7,"Receipt")
wtbn.write(0,8,"Nill Arrear")
ws.write(0,0,"NAME")
ws.write(0,1,"DEMAND")
ws.write(0,2,"COLLECTION")
ws.write(0,3,"COLL %")
ws.write(0,4,"NON PAYER")
ws.write(0,5,"STRIKE COUNT")
ws.write(0,6,"TOTAL RUNNING FILE")
ws.write(0,7,"STRIKE %")
ws.write(0,8,"STRIKE <1000")
ws.write(0,9,"NILL ARR <1000")
ws.write(0,10,"NILL ARR")
ws.write(0,11,"NILL ARR %")
ws.write(0,12,"S&S")
ws.write(0,13,"MATURED")
ws.write(0,14,"Above 3")
ws.write(0,15,"CNP")
ws.write(9,0,"Due date wise Details")
ws.write(9,1,"Total")
ws.write(9,2,"100% Coll")
ws.write(9,3,">90% Coll")
ws.write(9,4,">50% Coll")
ws.write(9,5,"<50% Coll")
ws.write(9,6,"0% Coll")
ws.write(10,0,"5th Due date Files")
ws.write(11,0,"10th Due date Files")
ws.write(12,0,"15th Due date Files")
ws.write(13,0,"20th Due date Files")

wb.save("Due Date Reportss.xls")
import xlwt
import xlrd
from xlutils.copy import copy
rb = xlrd.open_workbook('Due Date Reportss.xls')
wb = copy(rb)
w_sheet = wb.get_sheet(0)
w1_sheet = wb.get_sheet(1)
w2_sheet = wb.get_sheet(2)
pl_sheet = wb.get_sheet(3)
ab3_sheet = wb.get_sheet(4)
cnp_sheet = wb.get_sheet(5)
d5th_sheet = wb.get_sheet(6)
d10th_sheet = wb.get_sheet(7)
d15th_sheet = wb.get_sheet(8)
d20th_sheet = wb.get_sheet(9)

n = 0
arr = []
pl = []
for row in range(1, sheet.nrows, 1):
    arr.append(sheet.cell_value(row,5))

for row in range(1, sheetpln.nrows, 1):
    pl.append(sheetpln.cell_value(row,6))
    
free =[]
plln = []

for i in arr:
    if i not in free:
        free.append(i)

srun = sorted(free)

for row in range(1, len(srun)+1, 1):
    w_sheet.write(row,0,srun[row-1])
    
n=len(srun)
rw=1
rwpls=1
rw1=1
rw5=1
rw10=1
rw15=1
rw20=1

cm1=0
cmpls=0
cm=0
cm5=0
cm10=0
cm15=0
cm20=0

dd=0
dmd=0
cll=0
d5th=0
d10th=0
d15th=0
d20th=0
d100c=[]
da90c=[]
da50c=[]
db50c=[]
de0c=[]
demand=[]
sands = []
mat = []
cnp = []
abov3 = []
totalrunningfile=[]
collection=[]
strikecount=[]
tobestrike=[]
tobenillarr=[]
nillarrcount=[]
nonpayercount=[]
collpercent=[]

for row in range(0, n, 1):
    d100c.append(0)
    da90c.append(0)
    da50c.append(0)
    db50c.append(0)
    de0c.append(0)
    mat.append(0)
    cnp.append(0)
    sands.append(0)
    abov3.append(0)
    demand.append(0)
    collection.append(0)
    totalrunningfile.append(0)
    strikecount.append(0)
    tobestrike.append(0)
    tobenillarr.append(0)
    nillarrcount.append(0)
    nonpayercount.append(0)
    collpercent.append(0)

for value in range(0, len(srun), 1):
    for row in range(1, sheet.nrows, 1):
        if srun[value] == sheet.cell_value(row, 5) :
            if sheet.cell_value(row, 8) == "S&S":
                sands[value] = sands[value] + 1
                demand[value] = demand[value] + sheet.cell_value(row,21) + sheet.cell_value(row,22)
                collection[value] = collection[value] + sheet.cell_value(row,25)
            if sheet.cell_value(row, 8) == "Mat":
                mat[value] = mat[value] + 1
                demand[value] = demand[value] + sheet.cell_value(row,21) + sheet.cell_value(row,22)
                collection[value] = collection[value] + sheet.cell_value(row,25)
            if sheet.cell_value(row, 8) == "Run":
                if sheet.cell_value(row, 6) != "":
                    if sheet.cell_value(row, 41) != "NA":
                        if sheet.cell_value(row, 6) == 5 :
                            d5th=d5th+1
                            dd=0
                            dmd = sheet.cell_value(row,21) + sheet.cell_value(row,22)
                            if dmd != 0:
                                cll = sheet.cell_value(row,24)
                                colp=cll/dmd*100
                                if sheet.cell_value(row,26) <= 0:
                                    d100c[dd] = d100c[dd] + 1
                                if sheet.cell_value(row,26) > 0:
                                    if colp>90:
                                        da90c[dd] = da90c[dd] + 1
                                    if colp < 90 and colp > 50:
                                        da50c[dd] = da50c[dd] + 1
                                    if colp < 50 and cll > 0:
                                        db50c[dd] = db50c[dd] + 1
                                    if colp == 0:
                                        de0c[dd] = de0c[dd] +1
                    if sheet.cell_value(row, 6) == 10 :
                        d10th=d10th+1
                        dd=1
                        dmd = sheet.cell_value(row,21) + sheet.cell_value(row,22)
                        if dmd != 0:
                            cll = sheet.cell_value(row,24)
                            colp=cll/dmd*100
                            if sheet.cell_value(row,26) <= 0:
                                d100c[dd] = d100c[dd] + 1
                            if sheet.cell_value(row,26) > 0:
                                if colp>90:
                                    da90c[dd] = da90c[dd] + 1
                                if colp < 90 and colp > 50:
                                    da50c[dd] = da50c[dd] + 1
                                if colp < 50 and cll > 0:
                                    db50c[dd] = db50c[dd] + 1
                                if colp == 0:
                                    de0c[dd] = de0c[dd] +1
                    if sheet.cell_value(row, 6) == 15 :
                        d15th=d15th+1
                        dd=2
                        dmd = sheet.cell_value(row,21) + sheet.cell_value(row,22)
                        if dmd != 0:
                            cll = sheet.cell_value(row,24)
                            colp=cll/dmd*100
                            if sheet.cell_value(row,26) <= 0:
                                d100c[dd] = d100c[dd] + 1
                            if sheet.cell_value(row,26) > 0:
                                if colp>90:
                                    da90c[dd] = da90c[dd] + 1
                                if colp < 90 and colp > 50:
                                    da50c[dd] = da50c[dd] + 1
                                if colp < 50 and cll > 0:
                                    db50c[dd] = db50c[dd] + 1
                                if colp == 0:
                                    de0c[dd] = de0c[dd] +1
                    if sheet.cell_value(row, 6) == 20 :
                        d20th=d20th+1
                        dd=3
                        dmd = sheet.cell_value(row,21) + sheet.cell_value(row,22)
                        if dmd != 0:
                            cll = sheet.cell_value(row,24)
                            colp=cll/dmd*100
                            if sheet.cell_value(row,26) <= 0:
                                d100c[dd] = d100c[dd] + 1
                            if sheet.cell_value(row,26) > 0:
                                if colp>90:
                                    da90c[dd] = da90c[dd] + 1
                                if colp < 90 and colp > 50:
                                    da50c[dd] = da50c[dd] + 1
                                if colp < 50 and cll > 0:
                                    db50c[dd] = db50c[dd] + 1
                                if colp == 0:
                                    de0c[dd] = de0c[dd] +1
                if sheet.cell_value(row, 47) == "Y":
                    cnp[value] = cnp[value] + 1
                if sheet.cell_value(row, 6) != "":
                    totalrunningfile[value] = totalrunningfile[value] + 1
                demand[value] = demand[value] + sheet.cell_value(row,21) + sheet.cell_value(row,22)
                collection[value] = collection[value] + sheet.cell_value(row,25)
                strike = sheet.cell_value(row,21) + sheet.cell_value(row,22) - sheet.cell_value(row,24)
                if sheet.cell_value(row, 36) != "NA" :
                    if sheet.cell_value(row, 36) > 2.7 :
                        abov3[value] = abov3[value] + 1
                if sheet.cell_value(row, 37) != "NA" :
                    if sheet.cell_value(row, 26) > 0 & 0 < sheet.cell_value(row, 37) < 8 :
                        if sheet.cell_value(row, 6) == 5 :
                            d5th_sheet.write(rw5,cm5,sheet.cell_value(row,0))
                            d5th_sheet.write(rw5,cm5+1,sheet.cell_value(row,5))
                            d5th_sheet.write(rw5,cm5+2,sheet.cell_value(row,6))
                            d5th_sheet.write(rw5,cm5+3,sheet.cell_value(row,9))
                            d5th_sheet.write(rw5,cm5+4,sheet.cell_value(row,10))
                            d5th_sheet.write(rw5,cm5+5,sheet.cell_value(row,16))
                            d5th_sheet.write(rw5,cm5+6,sheet.cell_value(row,21)+sheet.cell_value(row,22))
                            d5th_sheet.write(rw5,cm5+7,sheet.cell_value(row,24))
                            d5th_sheet.write(rw5,cm5+8,sheet.cell_value(row,26))
                            d5th_sheet.write(rw5,cm5+9,sheet.cell_value(row,36))
                            d5th_sheet.write(rw5,cm5+10,sheet.cell_value(row,37))
                            rw5 = rw5 + 1
                        if sheet.cell_value(row, 6) == 10 :
                            d10th_sheet.write(rw10,cm10,sheet.cell_value(row,0))
                            d10th_sheet.write(rw10,cm10+1,sheet.cell_value(row,5))
                            d10th_sheet.write(rw10,cm10+2,sheet.cell_value(row,6))
                            d10th_sheet.write(rw10,cm10+3,sheet.cell_value(row,9))
                            d10th_sheet.write(rw10,cm10+4,sheet.cell_value(row,10))
                            d10th_sheet.write(rw10,cm10+5,sheet.cell_value(row,16))
                            d10th_sheet.write(rw10,cm10+6,sheet.cell_value(row,21)+sheet.cell_value(row,22))
                            d10th_sheet.write(rw10,cm10+7,sheet.cell_value(row,24))
                            d10th_sheet.write(rw10,cm10+8,sheet.cell_value(row,26))
                            d10th_sheet.write(rw10,cm10+9,sheet.cell_value(row,36))
                            d10th_sheet.write(rw10,cm10+10,sheet.cell_value(row,37))
                            rw10 = rw10 + 1
                        if sheet.cell_value(row, 6) == 15 :
                            d15th_sheet.write(rw15,cm15,sheet.cell_value(row,0))
                            d15th_sheet.write(rw15,cm15+1,sheet.cell_value(row,5))
                            d15th_sheet.write(rw15,cm15+2,sheet.cell_value(row,6))
                            d15th_sheet.write(rw15,cm15+3,sheet.cell_value(row,9))
                            d15th_sheet.write(rw15,cm15+4,sheet.cell_value(row,10))
                            d15th_sheet.write(rw15,cm15+5,sheet.cell_value(row,16))
                            d15th_sheet.write(rw15,cm15+6,sheet.cell_value(row,21)+sheet.cell_value(row,22))
                            d15th_sheet.write(rw15,cm15+7,sheet.cell_value(row,24))
                            d15th_sheet.write(rw15,cm15+8,sheet.cell_value(row,26))
                            d15th_sheet.write(rw15,cm15+9,sheet.cell_value(row,36))
                            d15th_sheet.write(rw15,cm15+10,sheet.cell_value(row,37))
                            rw15 = rw15 + 1
                        if sheet.cell_value(row, 6) == 20 :
                            d20th_sheet.write(rw20,cm20,sheet.cell_value(row,0))
                            d20th_sheet.write(rw20,cm20+1,sheet.cell_value(row,5))
                            d20th_sheet.write(rw20,cm20+2,sheet.cell_value(row,6))
                            d20th_sheet.write(rw20,cm20+3,sheet.cell_value(row,9))
                            d20th_sheet.write(rw20,cm20+4,sheet.cell_value(row,10))
                            d20th_sheet.write(rw20,cm20+5,sheet.cell_value(row,16))
                            d20th_sheet.write(rw20,cm20+6,sheet.cell_value(row,21)+sheet.cell_value(row,22))
                            d20th_sheet.write(rw20,cm20+7,sheet.cell_value(row,24))
                            d20th_sheet.write(rw20,cm20+8,sheet.cell_value(row,26))
                            d20th_sheet.write(rw20,cm20+9,sheet.cell_value(row,36))
                            d20th_sheet.write(rw20,cm20+10,sheet.cell_value(row,37))
                            rw20 = rw20 + 1
                if sheet.cell_value(row, 0) in pl :
                    if sheet.cell_value(row, 24) == 0 or sheet.cell_value(row, 26) > 0:
                        pl_sheet.write(rwpls,cmpls,sheet.cell_value(row,0))
                        pl_sheet.write(rwpls,cmpls+1,sheet.cell_value(row,5))
                        pl_sheet.write(rwpls,cmpls+2,sheet.cell_value(row,6))
                        pl_sheet.write(rwpls,cmpls+3,sheet.cell_value(row,9))
                        pl_sheet.write(rwpls,cmpls+4,sheet.cell_value(row,10))
                        pl_sheet.write(rwpls,cmpls+5,sheet.cell_value(row,16))
                        pl_sheet.write(rwpls,cmpls+6,sheet.cell_value(row,21)+sheet.cell_value(row,22))
                        pl_sheet.write(rwpls,cmpls+7,sheet.cell_value(row,24))
                        pl_sheet.write(rwpls,cmpls+8,sheet.cell_value(row,26))
                        rwpls = rwpls + 1
                if sheet.cell_value(row,24) == 0:
                    if sheet.cell_value(row,6) != "":
                        nonpayercount[value] = nonpayercount[value] + 1                
                if sheet.cell_value(row,41) == "Yes":
                    strikecount[value] = strikecount[value] + 1
                if sheet.cell_value(row,41) == "No" and 0 < strike < 1000:
                    w1_sheet.write(rw,cm,sheet.cell_value(row,0))
                    w1_sheet.write(rw,cm+1,sheet.cell_value(row,5))
                    w1_sheet.write(rw,cm+2,sheet.cell_value(row,6))
                    w1_sheet.write(rw,cm+3,sheet.cell_value(row,9))
                    w1_sheet.write(rw,cm+4,sheet.cell_value(row,10))
                    w1_sheet.write(rw,cm+5,sheet.cell_value(row,16))
                    w1_sheet.write(rw,cm+6,sheet.cell_value(row,21)+sheet.cell_value(row,22))
                    w1_sheet.write(rw,cm+7,sheet.cell_value(row,24))
                    w1_sheet.write(rw,cm+8,strike)
                    rw=rw+1
                    tobestrike[value] = tobestrike[value] + 1
                if 0 < sheet.cell_value(row,26) < 1000:
                    w2_sheet.write(rw1,cm1,sheet.cell_value(row,0))
                    w2_sheet.write(rw1,cm1+1,sheet.cell_value(row,5))
                    w2_sheet.write(rw1,cm1+2,sheet.cell_value(row,6))
                    w2_sheet.write(rw1,cm1+3,sheet.cell_value(row,9))
                    w2_sheet.write(rw1,cm1+4,sheet.cell_value(row,10))
                    w2_sheet.write(rw1,cm1+5,sheet.cell_value(row,16))
                    w2_sheet.write(rw1,cm1+6,sheet.cell_value(row,21)+sheet.cell_value(row,22))
                    w2_sheet.write(rw1,cm1+7,sheet.cell_value(row,24))
                    w2_sheet.write(rw1,cm1+8,sheet.cell_value(row,26))
                    rw1=rw1+1
                    tobenillarr[value] = tobenillarr[value] +1
                if sheet.cell_value(row,26)<=0:
                    nillarrcount[value] = nillarrcount[value] +1

matured=0
cl100 = 0
cl90 =0
cla50 = 0
clb50 = 0
cl0 = 0
SandSs = 0
above3=0
cnps=0
monthdemand=0
monthcollection=0
monthstrike=0
monthtotalrunningfile=0
monthnillarrear=0
monthnonpayer=0

for row in range(0, n, 1):
    cl100 = cl100 + d100c[row]
    cl90 = cl90 + da90c[row]
    cla50 = cla50 + da50c[row]
    clb50 = clb50 + db50c[row]
    cl0 = cl0 + de0c[row]
    matured=matured+mat[row]
    SandSs=SandSs+sands[row]
    above3=above3+abov3[row]
    cnps = cnps+cnp[row]
    monthdemand=monthdemand+demand[row]
    monthcollection=monthcollection+collection[row]
    monthstrike=monthstrike+strikecount[row]
    monthtotalrunningfile=monthtotalrunningfile+totalrunningfile[row]
    monthnillarrear=monthnillarrear+nillarrcount[row]
    monthnonpayer=monthnonpayer+nonpayercount[row]

for row in range(0, n, 1):
    w_sheet.write(row+1,1,demand[row])
    w_sheet.write(row+1,2,collection[row])
    w_sheet.write(row+1,3,collection[row]/demand[row]*100)
    w_sheet.write(row+1,4,nonpayercount[row])
    w_sheet.write(row+1,5,strikecount[row])
    w_sheet.write(row+1,6,totalrunningfile[row])
    w_sheet.write(row+1,7,strikecount[row]/totalrunningfile[row]*100)
    w_sheet.write(row+1,8,tobestrike[row])
    w_sheet.write(row+1,9,tobenillarr[row])
    w_sheet.write(row+1,10,nillarrcount[row])
    w_sheet.write(row+1,11,nillarrcount[row]/totalrunningfile[row]*100)
    w_sheet.write(row+1,12,sands[row])
    w_sheet.write(row+1,13,mat[row])
    w_sheet.write(row+1,14,abov3[row])
    w_sheet.write(row+1,15,cnp[row])
    w_sheet.write(row+10,2,d100c[row])
    w_sheet.write(row+10,3,da90c[row])
    w_sheet.write(row+10,4,da50c[row])
    w_sheet.write(row+10,5,db50c[row])
    w_sheet.write(row+10,6,de0c[row])

w_sheet.write(10,1,d5th)
w_sheet.write(11,1,d10th)
w_sheet.write(12,1,d15th)
w_sheet.write(13,1,d20th)
w_sheet.write(14,2,cl100)
w_sheet.write(14,3,cl90)
w_sheet.write(14,4,cla50)
w_sheet.write(14,5,clb50)
w_sheet.write(14,6,cl0)
w_sheet.write(9,7,"Nill %")
w_sheet.write(10,7,d100c[0]/d5th*100)
w_sheet.write(11,7,d100c[1]/d10th*100)
w_sheet.write(12,7,d100c[2]/d15th*100)
w_sheet.write(13,7,d100c[3]/d20th*100)

w_sheet.write(n+1,3,monthcollection/monthdemand*100)
w_sheet.write(n+1,1,monthdemand)
w_sheet.write(n+1,2,monthcollection)
w_sheet.write(n+1,4,monthnonpayer)
w_sheet.write(n+1,5,monthstrike)
w_sheet.write(n+1,7,monthstrike/monthtotalrunningfile*100)
w_sheet.write(n+1,6,monthtotalrunningfile)
w_sheet.write(n+1,10,monthnillarrear)
w_sheet.write(n+1,11,monthnillarrear/monthtotalrunningfile*100)
w_sheet.write(n+1,12,SandSs)
w_sheet.write(n+1,13,matured)
w_sheet.write(n+1,14,above3)
w_sheet.write(n+1,15,cnps)

wb.save("Due Date Reportss.xls")

