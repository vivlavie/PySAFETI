#PyDimXModSpreading.py
#Effect of process fire into adjacent module
#Coordinates of each module center is necessary
#


#Pool fire is neglected

from openpyxl import load_workbook
import math
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
import scipy.ndimage
from scipy.ndimage.filters import gaussian_filter
import dill
import sys
import matplotlib.pylab as pltparam

# sys.stdout = open('DimCube2.txt','w')

WantPlot = True
cube_result2file = True


Modules = ['P02','S02','P03','S03','P04','S04','P05','S05']
Coords = {}

Coords['LD1'] = [47.4, 3.5, 35, 14.1, 15.5, 0] #x,t,z,dx,dy,dz [m]
Coords['P01'] = [63.7, 7.8, 35, 33.6, 13.9, 17] #x,t,z,dx,dy,dz [m]
Coords['P02'] = [98.8, 4, 35, 20.2, 19.6, 17] #x,t,z,dx,dy,dz [m]
Coords['P03'] = [120.4, 4, 35, 20.6, 23, 9] #x,t,z,dx,dy,dz [m]
Coords['P04'] = [142.1, 4, 35, 25.3, 23, 17.5] #x,t,z,dx,dy,dz [m]
Coords['P05'] = [168.2, 4, 35, 25.4, 23, 16] #x,t,z,dx,dy,dz [m]
Coords['S00'] = [48, -22.7, 35, 13.5, 17.5, 14.5] #x,t,z,dx,dy,dz [m]
Coords['S01'] = [63.4, -27, 35, 34.5, 23, 10] #x,t,z,dx,dy,dz [m]
Coords['S02'] = [98.6, -27, 35, 20.7, 23, 0] #x,t,z,dx,dy,dz [m]
Coords['S03'] = [120.5, -27, 35, 20.9, 23, 19.4] #x,t,z,dx,dy,dz [m]
Coords['S04'] = [142, -27, 35, 25.5, 23, 17] #x,t,z,dx,dy,dz [m]
Coords['S05'] = [168.2, -27, 35, 28, 23, 17] #x,t,z,dx,dy,dz [m]
Coords['LD2'] = [200, -23.4, 29.1, 20.8, 10.1, 5.9] #x,t,z,dx,dy,dz [m]
Coords['R01'] = [45, -3.2, 35, 40, 6.4, 0] #x,t,z,dx,dy,dz [m]
Coords['R02'] = [85.5, -3.2, 35, 34.5, 6.4, 18.6] #x,t,z,dx,dy,dz [m]
Coords['R03'] = [120.3, -3.2, 35, 21.4, 6.4, 18.9] #x,t,z,dx,dy,dz [m]
Coords['R04'] = [142, -3.2, 35, 26, 6.4, 18.9] #x,t,z,dx,dy,dz [m]
Coords['R05'] = [168.1, -3.2, 35, 28.4, 6.4, 18.9] #x,t,z,dx,dy,dz [m]
Coords['Turret']= [202.5, -9.5, 0, 19, 19, 42] #x,t,z,dx,dy,dz [m]
Coords['KOD'] = [199.7, 10.1, 29.1, 16.3, 14.5, 14.7] #x,t,z,dx,dy,dz [m]


ModX = {}
ModX['P02'] = [['P03'],['S02'],['S03']]
ModX['P03'] = [['P02','P04'],['S03'],['S02','S04']]
ModX['P04'] = [['P03','P05'],['S04'],['S03','S05']]
ModX['P05'] = [['P04','KOD'],['S05'],['S04']]

ModX['S02'] = [['S03'],['P02'],['P03']]
ModX['S03'] = [['S02','S04'],['P03'],['P02','P04']]
ModX['S04'] = [['S03','S05'],['P04'],['P03','P05']]
ModX['S05'] = [['S04'],['P05'],['P04']]


def print_cum_cube(cube,AA):
    F=0.
    print("{:35s} Mod  (Freq. ) - Jet Duration  CumFreq".format(cube))
    for e in AA[::-1]:
        # pv,hole,weather = e[2].split("\\")
        # hole = hole.split("_")[0]
        F += e[1]
        pv,hole,weather = e[2].split("\\")
        # mod = Modules[pv]
        print("{:35s} {:8.2e} {:8.1f}   {:8.2e}".format(e[2],e[1],e[0],F))
        if F>1.0E-3:
            break
def print_cum_cube_file(cube,AA,a_file):
    F=0.
    print("{:35s} Mod  (Freq. ) - Jet Duration  CumFreq".format(cube),file = a_file)
    for e in AA[::-1]:
        # pv,hole,weather = e[2].split("\\")
        # hole = hole.split("_")[0]
        F += e[1]
        pv,hole,weather = e[2].split("\\")
        # mod = Modules[pv]
        print("{:35s} {:8.2e} {:8.1f}   {:8.2e}".format(e[2],e[1],e[0],F),file = a_file)
        if F>1.0E-3:
            break

import dill
# Area = "ProcessArea"
element_dump_filename = 'Bv06_dump'
# icubeloc='SCE_CUBE_XYZ2_Process'

with open(element_dump_filename,'rb') as element_dump:
    lEvent = dill.load(element_dump)


def jffit(m):
    jl_lowe = 2.8893*np.power(55.5*m,0.3728)
    if m>5:
        jf = -13.2+54.3*math.log10(m)
    elif m>0.1:
        jf= 3.736*m + 6.
    else:
        jf = 0.
    # print(m, jl_lowe,jf)
    return jl_lowe

#Failure of fire detector for immediate ignition
P_FD_Fail = 0.05 #default value? #should be adjusted for hole size or leak rate? To be suggested to use the release rate!
#Failure of gas detectin for delayed ignition
P_GD_Fail = 0.05

ImpingeDurationArray = []

D = {}
for m in ModX:
    # print(m)
    # print(ModX[m][0]) # front or back
    # print(ModX[m][1]) # side
    # print(ModX[m][2]) # diagonal
    xm = Coords[m][0] + 0.5*Coords[m][3]
    ym = Coords[m][1] + 0.5*Coords[m][4]
    
    for mm in ModX[m][0]: #for modues in X direction
        dx = min(abs(xm - Coords[mm][0]), abs(xm - (Coords[mm][0] + Coords[mm][3])))        
        D[m + mm] = np.sqrt(dx*dx)
    for mm in ModX[m][1]: #for modues in Y direction
        dy = min(abs(ym - Coords[mm][1]), abs(ym - (Coords[mm][1] + Coords[mm][4])))        
        D[m + mm] = np.sqrt(dy*dy)
    for mm in ModX[m][2]: #for modues in X direction
        dx = min(abs(xm - Coords[mm][0]), abs(xm - (Coords[mm][0] + Coords[mm][3])))                
        dy = min(abs(ym - Coords[mm][1]), abs(ym - (Coords[mm][1] + Coords[mm][4])))        
        D[m + mm] = np.sqrt(dx*dx + dy*dy)

CubeImpingeDuration = {}      #To record source-target exceedance duration

# for mm in {key: D[key] for key in ['P03P02']}:    #For each combinatin of source (module) & target (distance),
for mm in D:
    s = mm[0:3]
    t = mm[3:]
    rr = D[mm]

    ImpingeDurationArray = []

   
    for ei in range(0,len(lEvent)):
        e = lEvent[ei]
        # print(id,e.Module, e.Key)
        if s == e.Module:
            #Array of jet length
            jl_e = e.jfscale*2.8893*np.power(55.5*e.TVD[:,2],0.3728) #the 3rd column of the matrix, Lowesmith formulat
            t_e = interp1d(jl_e,e.TVD[:,0],kind='linear') #Read the time when 'jl' is equal to the distance 'rr'            

            if e.Discharge.ReleaseRate <= 1:
                P_FD_Fail = 0.1
                P_GD_Fail = 0.1
            elif e.Discharge.ReleaseRate <= 10:
                P_FD_Fail = 0.05
                P_GD_Fail = 0.05
            elif e.Discharge.ReleaseRate > 10:
                P_FD_Fail = 0.005
                P_GD_Fail = 0.005
            else:
                print('something wrong in P_FD_Fail')
            
            if True:
                if ('EXBX' in e.Key) or ('EXBN' in e.Key):
                    #When ESD fails, whether it is dueto Fire detection failure should be dinstinguished
                    #ESD fail even upon Fire detection successful
                    #Probability of failure of ESD and BDV is already reflected in e.JetFire.Frequency???
                    ff = e.JetFire.Frequency*(1-P_FD_Fail)
                    if max(jl_e) < rr:                                     
                        ImpingeDurationArray.append([0.,0.,e.Key+"FO_Jet"])                    
                    elif min(jl_e) > rr:                
                        ImpingeDurationArray.append([e.TVD[-1,0],ff,e.Key+"FO_Jet"])                             
                    else:                
                        ImpingeDurationArray.append([t_e(rr),ff,e.Key+"FO_Jet"])                    

                    #ESD is not activated due to fire detection failure, any release irrespective of success of ESDV and BDV will result in EXBX release case      
                    ff = e.JetFire.Frequency/e.PESD/e.PBDV*P_FD_Fail
                    if max(jl_e) < rr:                                     
                        ImpingeDurationArray.append([0.,0.,e.Key+"FX_Jet"])                    
                    elif min(jl_e) > rr:                
                        ImpingeDurationArray.append([e.TVD[-1,0],ff,e.Key+"FX_Jet"])                             
                    else:                
                        ImpingeDurationArray.append([t_e(rr),ff,e.Key+"FX_Jet"])                    
                else:  #Fire detection succedded
                    #Probability of success of ESD is already reflected in e.JetFire.Frequency???
                    ff = e.JetFire.Frequency*(1-P_FD_Fail)
                    if max(jl_e) < rr:                                     
                        ImpingeDurationArray.append([0.,0.,e.Key+"FO_Jet"])                    
                    elif min(jl_e) > rr:                
                        ImpingeDurationArray.append([e.TVD[-1,0],ff,e.Key+"FO_Jet"])                             
                    else:                
                        ImpingeDurationArray.append([t_e(rr),ff,e.Key+"FO_Jet"])                    
      
    #To pin-point a scenario that give the dimensioning scenario
    IDAsorted = sorted(ImpingeDurationArray, key = lambda fl: fl[0]) #with the longest duration at the bottom
    
    cf = 0
    jp = IDAsorted[-1]
    DimFreq = 1.0E-4
    di = 0
    cfp = jp[1]#frequency for the longest duration jet fire
    InterpolationSuccess = False
    if cfp > DimFreq:
        di = jp[0]
        scn = jp[2]
    else:
        for j in IDAsorted[-2:0:-1]:
            cf = cfp + j[1]
            if cf >= DimFreq and cfp < DimFreq:
                dp = jp[0]
                dn = j[0]
                # di = (dn-dp)/(cf-cfp)*(DimFreq - cfp) + dp
                # print('Dimensioning jet duration {:8.1f}'.format(di))
                scnp = jp[2]
                scn = j[2]
                InterpolationSuccess = True
                #Choose the scenario close to the threshold
                if abs(cf-DimFreq) > abs(cfp-DimFreq):
                    scn = scnp
                    di = dp
                else:
                    di = dn
                break
            cfp = cf
            # print(cfp,cf)
            jp = j
    
    
    if cube_result2file == True:
        f_cube_result = open(mm +".txt","w")
    if InterpolationSuccess:
        CubeImpingeDuration[mm] = di
        print(s,t,rr,di)
        if cube_result2file == True:                
            print_cum_cube_file(mm,IDAsorted,f_cube_result) 
        
    elif not InterpolationSuccess:
        print(" No dimensioning scenario",mm)
        scn = 'No dimensioning scenario for ' + mm
        # print_cum_cube(id+" No dimensioning scenario ",IDAsorted)        
        if cube_result2file == True:
            print_cum_cube_file(mm+" No dimensioning scenario ",IDAsorted,f_cube_result)        
    


    # if (InterpolationSuccess == True) and (WantPlot == True):
    if (WantPlot == True):
        ll = len(IDAsorted)
        ec = np.zeros([ll,2])
        i=0
        ec[i,1] = IDAsorted[-1][0] #the longest duration
        ec[i,0] = IDAsorted[-1][1] #Frequency for the longest duration
        for i in range(1,len(IDAsorted)):
            ec[i,1] = IDAsorted[ll-i-1][0]
            ec[i,0] = ec[i-1,0] + IDAsorted[ll-i-1][1]

        # plt.figure(figsize=(5.91, 3.15))
        CF = ec[1:,0]
        JFL = ec[1:,1]
        masscolor = 'tab:blue'
        fig,ax1 = plt.subplots()
        ax1.set_xlabel('Jet Impingement Duration [sec]')
        ax1.set_ylabel('Cumulative Frequency [#/year]',color=masscolor)
        ax1.semilogy(JFL,CF,color=masscolor)
        ax1.set_ylim(top=5E-4,bottom=1E-6)
        ax1.set_xlim(left=0, right=3600)
        ax1.tick_params(axis='y',labelcolor=masscolor)
        ax1.xaxis.set_major_locator(plt.FixedLocator([300, 600, 1800, 3600]))
        ax1.annotate(scn,xy=(di,1.0E-4),xytext=(di,2E-4),horizontalalignment='left',verticalalignment='top',arrowprops = dict(facecolor='black',headwidth=4,width=2,headlength=4))
        # ax.xaxis.set_major_formatter(plt.FixedFormatter(['2/3','3/4','4/5','S05']))
        # ax.yaxis.set_major_locator(plt.FixedLocator([-27, -3.1, 3.1, 27]))
        # ax.yaxis.set_major_formatter(plt.FixedFormatter(['ER_S','Tray_S','Tray_P','ER_P']))
        ax1.grid(True,which="major")

        if InterpolationSuccess:
            plt.title("{:5s} -> {:5s} ({:8.2f}) for {:8.1f} [sec]".format(s,t,rr, di))            
        else:
            plt.title("{:5s} -> {:5s} ({:8.2f}) - No dimensioning fire".format(s,t, rr))            
        plt.tight_layout()
        plt.show()

        fig.savefig("{}.png".format(mm))
        plt.close()

    if cube_result2file == True:
        f_cube_result.close()

