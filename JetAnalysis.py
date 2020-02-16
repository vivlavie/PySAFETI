#Sort in ascending order for J00, J05, ...
# #for J in [J00, J05, J10, J30, J60]:
#For each time
""" Threshold = '05'
Jset = J05
RRIndex = 1
RRThreshold = 0.01 """

DimFreq = 1E-4

#Threshold, Jset, RRIndex, RRThreshold = '05', J05, 1, 0.05
RRThreshold = 0.1
Jsets = {'05':J05}
RRIndices = {'05':1}
# Jsets = {'00':J00, '05':J05,'10':J10,'30':J30,'60':J60}
# RRIndices = {'00':0, '05':1,'10':2,'30':3,'60':4}
#load JetAnalysis.py
# for Threshold in ['00', '05','10','30','60']:
for Threshold in ['05']:
    RRIndex = RRIndices[Threshold]
    Jset = Jsets[Threshold]

    #Threshold, Jset, RRIndex, RRThreshold = '05', J05, 1, 0.05
    #JetAnalysis.py
    for J in [Jset]:    
        #For each module
        # for mod in ['P03','P04','P05','S03','S04','S05']:
        # for mod in ['P04','P05','S04','S05']:
        for mod in ['S04']:            
            jd = []            
            for e in lEvent:
                if e.Module == mod:
                    key = e.Key
                    #print(key)
                    """ lEq.append(key)
                    lJF.append(JetFrequency[key])
                    lJL.append(J[key]) """
                    if e.Discharge.RRs[RRIndex] > RRThreshold:
                        jd.append([key, JetFrequency[key], J[key]])
                    # cf_jl = np.vstack([cf_jl,[JetFrequency[key], J[key]]])
            if len(jd) > 0:
                Js = sorted(jd, key = lambda fl: fl[2])
                # for j in Js:
                    # topside,pressure_vessel,eqtag,hole,weather = j[0].split("\\")
                    # print("{} {:30s} {:.2e} {:6.1f}".format(mod, eqtag,j[1],j[2]))
                cJs = Js
                ec = np.zeros([len(cJs),2])
                ec[-1,:]=[cJs[-1][1],cJs[-1][2]]
                InterpolationSuccess = False
                for ir in range(len(cJs)-2,-1,-1):
                    cp = cJs[ir+1][1]
                    cJs[ir][1] += cp
                    cn = cJs[ir][1]
                    ec[ir,:] = [cn,cJs[ir][2]]
                    if cn >= DimFreq and cp < DimFreq:
                        jp = cJs[ir+1][2]
                        jn = cJs[ir][2]
                        j0 = (jn-jp)/(cn-cp)*(DimFreq - cp) + jp
                        print('{}: {}: Dimensioning jet length {:8.1f} with RR limit {:6.2f}'.format(Threshold,mod, j0,RRThreshold))
                        InterpolationSuccess = True
                        break
                    if (ir == len(cJs)-2 ) and (cp >= DimFreq):
                        jp = cJs[ir+1][2]                        
                        j0 = jp
                        print('{}: {}: Dimensioning jet length {:8.1f} with RR limit {:6.2f}-The largest event'.format(Threshold,mod, j0,RRThreshold))
                        InterpolationSuccess = True
                if InterpolationSuccess == False:
                    print('{}: {}: Cumulative frequency {:.2e} not larger than {} Interpolatin fail'.format(Threshold,mod,cn,DimFreq))
            else:
                print('{}: {}: No Event for data processing'.format(Threshold, mod))            
            if ir == 0 and j0 == None:
                print('{}: {}: No dimensioning jet length'.format(Threshold,mod))

            """ fig,ax1 = plt.subplots()
                    
            CF = ec[1:,0]
            JFL = ec[1:,1]
            
            masscolor = 'tab:blue'
            ax1.set_xlabel('Jet Length [m]')
            ax1.set_ylabel('Cumulative Frequency [#/year]',color=masscolor)
            ax1.semilogy(JFL,CF,color=masscolor)
            ax1.set_ylim(bottom=1E-6)
            ax1.tick_params(axis='y',labelcolor=masscolor)
            ax1.grid(True,which="major")           
            
            
            plt.title("Jet length from " +mod+" at " +Threshold)            
            plt.show()
            fig.savefig("{}T{}.png".format(mod,Threshold))
            plt.close() """
e_dim_jet_low = next((x for x in lEvent if x.Key == cJs[ir-1][0]), None)         
e_dim_jet_high = next((x for x in lEvent if x.Key == cJs[ir][0]), None)

def print_cJs(cJs):
    F=0.
    print("{:20s} (Freq. ) - Jet Length  CumFreq".format("Scenario"))
    for e in cJs:
        ts,pv,IS,hole,weather = e[0].split("\\")
        hole = hole.split("_")[0]
        F += e[1]
        print("{:20s} ({:8.2e}) -  {:8.1f}   {:8.2e}".format(IS+"_"+hole,e[1],e[2],F))
