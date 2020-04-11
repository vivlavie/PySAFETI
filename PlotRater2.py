cutoffrate = 0.1
numscn = 0
numLeakRates = 46
FigurePLot = True
FigureEventPLot = False
PrintRater = True



#reading a rater file and plot is done by 'ReadRater.py'



import matplotlib.pyplot as plt
import dill
element_dump_filename = 'Bv06_dump'
with open(element_dump_filename,'rb') as element_dump:
    lEvent = dill.load(element_dump)

from scipy.interpolate import interp1d
import numpy as np

ISs = []
for e in lEvent:
    study, pv, folder = e.Study_Folder_Pv.split('\\')
    if (not (pv in ISs)) and ( not('-DB' in pv)) and ( not('-DC' in pv)):
        ISs.append(pv)
scns = ISs
numscn = len(scns)




i = 0
j = 0 #print figure index
# for s in scns:


numHoles = 4
numCases = 4
id_vs_rr = [0.1, 0.2, 0.4, 0.5, 0.8, 1, 1.2, 2, 3, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 22, 26, 30, 34, 38, 40, 46, 50, 55, 60, 65, 70, 75, 85, 95, 105, 115, 125, 150, 175, 200, 250, 300, 400, 600, 800, 1000]
# id_vs_rr = [0.1, 0.2, 0.4, 0.5, 0.8, 1, 1.2, 2, 3, 4, 5, 6, 8, 10]
numLeakRates = len(id_vs_rr)
# RR = [] #to be of the dimension numRates x numscn
RRSet =  [[[False,[]] for x in range(numLeakRates)] for x in range(numscn+1)] # not only a bool whether 'rr' is set but also the data itself
# for s in [scns[0]]:
i = 0
ScnIndex = {}
ScnsToAnalyse =  scns[35:40]
ScnsToAnalyse =  scns
for s in ScnsToAnalyse:
# for s in scns:
    i += 1 #id for section
    ScnIndex[s] = i    
    es = []
    for hole in ['SM','ME','MA','LA']:
        for e in lEvent:        
            if (s in e.Key) and ((hole+"_EOBO" == e.Hole) or (hole+"_EOBN" == e.Hole)) and (e.Weather == '7.7D'):
            # if (s in e.Key) and ((hole+"_EOBO" == e.Hole) or (hole+"_EOBN" == e.Hole)) and (e.Weather == '7.7D'):
                es.append(e)
    hole = 'LM'
    for e in lEvent:        
        if (s in e.Key) and ((hole+"_EOBO" == e.Hole) or (hole+"_EOBN" == e.Hole)) and (e.Weather == '7.7D') and (not ( s == '045-01-G')):
            es.append(e)    
    
    
    if FigureEventPLot == True:
        fig,axes = plt.subplots(numHoles,numCases)
        j = 0
    
    #Events for section 's' is now in 'es'
    for e in es:                  
        GasFraction = max(1-e.Discharge.LiquidMassFraction,0.05)        
        MassReleasedTotal1 = GasFraction*(e.Discharge.RRs[0]*(120+e.PESD) + (e.Discharge.Ms[0] - e.Discharge.Ms[-1]))
        # MassReleasedTotal2 = sum(rri[1:]+rri[:-1])*0.5 #rri should have been interpolated with dt=1
        MassReleasedTotal = MassReleasedTotal1

        if ("-DB" in e.Key) or ("-DC" in e.Key):
            print("DB or DC shall not be in rater.txt. ", e.Key)
            break
        # TVD for the release rate calcualted in PHAST but re-arranged for 1 sec interval
        # if (s in e.Key) and (True):        
        rr0 = GasFraction*e.Discharge.ReleaseRate
        Tend = int(max(e.TVD[:,0]))
        tti = np.arange(0,Tend+1)
        # tti = np.arange(0,3600)
        tt_rr = np.array(e.TVD)
        tt = tt_rr[:,0]
        rr = tt_rr[:,2]
        r_e = interp1d(tt,GasFraction*rr,kind='linear') #the data is reade having been scaled considering the 'GasFraction'
        rri = r_e(tti)
        # k = len(rri)-1
        
        if FigureEventPLot == True: #To plot Discharge profile read from lEvent:
            axes[int(j/numHoles),j%numCases].plot(tti,rri)
            axes[int(j/numHoles),j%numCases].set_title("{:s} {:s} - {:s} m0: {:6.1f}".format(s,e.Key, e.Hole, rri[0]))
            j += 1        

        # TVD for the release rates designated for 'rater.txt' using the 1 sec interval for the release rate calculated in PHAST
        # for release rates smaller than a value read from 'e'
        #Iteration for different initial release rate
        rr_i = 0
        ri0 = id_vs_rr[rr_i]
        while ri0 < GasFraction*e.Discharge.ReleaseRate:
            if RRSet[i][rr_i][0] == False:
                rri_scaled = list( (float(ri0/rri[0]*rr) for rr in rri))
                RRSet[i][rr_i][1] = rri_scaled                
                """ k = 0                
                while (rri_scaled[k] <= ri0) and (k < (len(rri)-1)):
                    RRSet[i][rr_i][1][k] = max(ri0,rri_scaled[k])                    
                    k += 1
                RRSet[i][rr_i][1][k] = max(ri0,rri_scaled[k])                
                while (k < 3600) and (sum(RRSet[i][rr_i][1]) < MassReleasedTotal):
                    RRSet[i][rr_i][1].append(ri0)
                    k += 1 """
                # RR.append([i,rr_i,list( (float(id_vs_rr[rr_i]/rri[0]*rr) for rr in rri))])                
                RRSet[i][rr_i][0] = True
                if rri_scaled == []:
                    print(i,j,'rri scaled []')
            rr_i += 1
            ri0 = id_vs_rr[rr_i]        
        
        if (("LM" in e.Hole) or ("LA" in e.Hole) or ((s=='045-01-G') and ('MA' in e.Hole)) or ((s=='046-01-04-L') and ('SM' in e.Hole))) and (GasFraction*e.Discharge.ReleaseRate < id_vs_rr[-1]):
            #Extrapolate the consequence
            print("Am in in LM/LA logic?",s,e.Hole)
            for rr_j in range(rr_i,len(id_vs_rr)):                
                if RRSet[i][rr_j][0] == False:
                    rri_scaled = list( (float(id_vs_rr[rr_j]/rri[0]*rr) for rr in rri))
                    RRSet[i][rr_j][1] = rri_scaled
                    # RR.append([i,rr_j,list( (float(id_vs_rr[rr_j]/rri[0]*rr) for rr in rri))])
                    RRSet[i][rr_j][0] = True
                    if rri_scaled == []:
                        print(i,j,'rri scaled [] for LM')
                    # print("RR: {:8.2f} & Scale {:8.2f}".format(id_vs_rr[rr_j], id_vs_rr[rr_j]/rri[0]))
            if PrintRater == False:
                print(i,' read up to',rr_j)
            # if rr_i > 0:
            #     for rr_i in range(0,len(id_vs_rr)):
            #         print(RR[rr_i][0],RR[rr_i][2][-3:-1])
    for j in range(0,numLeakRates):
        #Total release mass shall be considered
        k = 0
        rri = RRSet[i][j][1]
        MassReleased = 0.5*(rri[k+1] + rri[k])*1 #Time difference, 1 sec
        while (MassReleased < MassReleasedTotal) and (k < len(rri)-2):
            k += 1
            MassReleased += 0.5*(rri[k+1] + rri[k])*1 #Time difference, 1 sec
            #if the release rate is too low, the total MassRelesed would be less than that durign 1 hour.
        
        #if MassRelease is already larger than MassReleasedTotal for k=0,

        # 'k' then should be the index for the end of the time or till the release made upto the MassReleasedTotal
        # k = len(RRSet[i][j][1])
        # Release rate less than 0.1 kg/sec should be neglected
        if rri == []:
            print(i,j,"rri none before removing mdot < 0.1")
        while (rri[k-1] < 0.1) and (k>1):
            k -= 1
        rri_trimmed = rri[0:k]            #Still rri[0:0] ... not []
        RRSet[i][j][1] = rri_trimmed

        if rri == []:
            print(i,j,"rri none after removing mdot < 0.1")


        #Mass released update
        k = 0        
        if len(rri) > 1:
            MassReleased2 = 0.5*(rri[k+1] + rri[k])*1 #Time difference, 1 sec
            while (MassReleased2 < MassReleasedTotal) and (k < len(rri)-2):
                k += 1
                MassReleased2 += 0.5*(rri[k+1] + rri[k])*1 #Time difference, 1 sec            
            print(MassReleasedTotal, "?=", MassReleased, MassReleased2, ":", rri[0], "released during ", len(rri), " seconds")
        
        else: #at leat rri having more than 1 element
            print(MassReleasedTotal, "?=", MassReleased, ":", rri[0], "released during ", len(rri), " seconds")

    # print(j+1,i,'; '.join(map(str,"{:8.2f}".format(rri[0:5]))),' ',sep='; ')#not working ...
    # print(j+1,i,'; '.join(map(str,rri)),' ',sep='; ',file=raterfile)
       


# for s in [scns[10]]:
# for s in scns:
if PrintRater == True:    
    raterfile = open('rater_vis.txt','w')
    print(0.1,'; ',numscn,'; ',file=raterfile)
    print('; '.join(map(str,scns)),';',file=raterfile) # what if there happens an errro with the trailing ' ;' with a blank cell
    for s in ScnsToAnalyse:
        i = ScnIndex[s]
        #After construction RR, print it out with rr > 0.1 kg/sec up to the total amount
               
        #index 'i' for identification of the section
        # print(j+1,i,'; '.join(map(str,rri)),' ',sep='; ',file=raterfile)
        #If to trim the release rate less than 0.1 kg/sec, which should not be made if it is to fit the curve to a larger release rate
        for j in range(0,numLeakRates):            
            rri = RRSet[i][j][1]
            if rri == []:
                print(j+1,i,'; '.join(map("{:.2f}".format,[id_vs_rr[j]])),' 0.1;',sep='; ',file=raterfile)
            else:
                print(j+1,i,'; '.join(map("{:.2f}".format,rri)),' 0.1;',sep='; ',file=raterfile)
        
        # is_to_plot = '013-01-N1-G'
        if FigurePLot == True:                  
            is_to_plot = s
            i = ScnIndex[is_to_plot]
            fig,axes = plt.subplots(8,5)       
            for j in range(1,numLeakRates-10):
            # for j in range(0,1):
                rri = RRSet[i][j][1]
                if not ( rri == []):
                    tti = range(0,len(rri))
                    axes[int(j/5),j%5].plot(tti,rri)
                    axes[int(j/5),j%5].set_title("{:6.1f}/{:4d}".format(rri[0],tti[-1]))
                    #Another title
                    # axes[int(j/5),j%5].set_title("Section {:s} - mdot(t=0): {:6.1f}".format(is_to_plot,rri[0]))
                else:
                    print("Skipping plotting ",s," for rate ",id_vs_rr[j])
            fig.savefig("rater_{}.png".format(s))
            plt.close()
    raterfile.close()        
        

# for s in [scns[0]]:
#     ISs = []
#     for e in lEvent:
#         if (s in e.Key) and ("EOBO" in e.Hole):        
#             is1,hole,weather = e.Key.split("\\")
#             ISs.append([is1, hole, e.Discharge.ReleaseRate])
#     print(s,ISs)

#for 023-01, the RR from LM case is 5.6 kg/sec. Extrapolating this to 1000?

