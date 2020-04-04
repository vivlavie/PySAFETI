#read rater.txt and plot

f = open('rater.txt')
cutoffrate = 0.
numscn = 0
numLeakRates = 46
scns = []
cutoffrate,numscn,dmy = f.readline().split(";")
numscn = int(numscn)
scns = f.readline().split(";")
scns = scns[:-2]
TVD = []
for i in range(0,numscn):
    for j in range(0,numLeakRates):
        rr = f.readline().split(";")
        if (int(rr[0]) != j+1) or (int(rr[1]) != i+1):
            print(rr[1],i,rr[0],j+1)
            break
        else:
            # rr = float(rr[2:-1])
            rr = rr[2:-1]
            TVD.append([ list( (float(d) for d in rr))])

f.close()

""" import matplotlib.pyplot as plt
fig,axes = plt.subplots(12,4)
# for i in range(0,numscn):
# for i in range(0,1):
for i in range(1,2):
    for j in range(0,numLeakRates):
        axes[int(j/4),j%4].plot(TVD[i*numLeakRates+j][0])
        axes[int(j/4),j%4].set_title("Section {:s} - m0: {:6.1f}".format(scns[i],TVD[i*numLeakRates+j][0][0]))
        # axes[int(j/4),j%4].
 """

import dill
element_dump_filename = 'Bv06_dump'
with open(element_dump_filename,'rb') as element_dump:
    lEvent = dill.load(element_dump)

from scipy.interpolate import interp1d
import numpy as np


print(numscn,'; ',0.1,'; ')
print('; '.join(map(str,scns)),';') # what if there happens an errro with the trailing ' ;' with a blank cell

i = 0
j = 0 #print figure index
# for s in scns:
FigurePLot = False
PrintRater = False

numHoles = 4
numCases = 4
# id_vs_rr = [0.1, 0.2, 0.4, 0.5, 0.8, 1, 1.2, 2, 3, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 22, 26, 30, 34, 38, 40, 46, 50, 55, 60, 65, 70, 75, 85, 95, 105, 115, 125, 150, 175, 200, 250, 300, 400, 600, 800, 1000]
id_vs_rr = [0.1, 0.2, 0.4, 0.5, 0.8, 1, 1.2, 2, 3, 4, 5, 6, 8, 10]
numLeakRates = len(id_vs_rr)
RR = [] #to be of the dimension numRates x numscn
RRSet =  [[False for x in range(numLeakRates)] for x in range(numscn)] 
for s in [scns[0]]:

    i += 1
    if FigurePLot == True:
        fig,axes = plt.subplots(numHoles,numCases)
        j = 0
    for e in lEvent:        
        if (s in e.Key) and (("EOBO" in e.Hole) or ("EOBN" in e.Hole)):        
            print(s, e.Key)
        # TVD for the release rate calcualted in PHAST but re-arranged for 1 sec interval
        # if (s in e.Key) and (True):
            rr0 = e.Discharge.ReleaseRate
            Tend = int(max(e.TVD[:,0]))
            tti = np.arange(0,Tend+1)
            r_e = interp1d(e.TVD[:,0],e.TVD[:,2],kind='linear')
            rri = r_e(tti)
            k = len(rri)-1
            #If to trim the release rate less than 0.1 kg/sec, which should not be made if it is to fit the curve to a larger release rate
            # while (k > 0) and (rri[k] < 0.1):
            #     k -= 1
            # if k>1:
            #     tti = tti[:k]
            #     rri = rri[:k]

            if FigurePLot == True:
                axes[int(j/numHoles),j%numCases].plot(tti,rri)
                axes[int(j/numHoles),j%numCases].set_title("{:s} {:s} - {:s} m0: {:6.1f}".format(s,e.Key, e.Hole, rri[0]))
                j += 1

            if PrintRater == True:
                print(s,e.Hole,j+1,i,rri,sep=';')
                print(j+1,i,'; '.join(map(str,rri)),' ',sep='; ')

            # TVD for the release rates designated for 'rater.txt' using the 1 sec interval for the release rate calculated in PHAST
            # for release rates smaller than a value read from 'e'
            rr_i = 0
            while id_vs_rr[rr_i] < e.Discharge.ReleaseRate:
                if RRSet[i][rr_i] == False:
                    RR.append([i,rr_i,list( (float(id_vs_rr[rr_i]/rri[0]*rr) for rr in rri))])
                    RRSet[i][rr_i] = True
                rr_i += 1
            if rr_i == 0:
                print(i,' skipped')
            else:
                print(i,' Processed')
                # print('Set up to',i,rr_i-1)
            if (("LM" in e.Hole) or ("LA" in e.Hole)) and (e.Discharge.ReleaseRate < id_vs_rr[-1]):
                #Extrapolate the consequence
                for rr_i in range(rr_i,len(id_vs_rr)):
                    RR.append([i,rr_i,list( (float(id_vs_rr[rr_i]/rri[0]*rr) for rr in rri))])
                    RRSet[i][rr_i] = True
                print(i,' read up to',rr_i)
            if rr_i > 0:
                for rr_i in range(0,len(id_vs_rr)):
                    print(RR[rr_i][0],RR[rr_i][2][-3:-1])


# for s in [scns[0]]:
#     ISs = []
#     for e in lEvent:
#         if (s in e.Key) and ("EOBO" in e.Hole):        
#             is1,hole,weather = e.Key.split("\\")
#             ISs.append([is1, hole, e.Discharge.ReleaseRate])
#     print(s,ISs)

#for 023-01, the RR from LM case is 5.6 kg/sec. Extrapolating this to 1000?


