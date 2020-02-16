from openpyxl import load_workbook
import math
import numpy as np
# import matplotlib.pyplot as plt
from scipy.interpolate import interp1d
# import scipy.ndimage
# from scipy.ndimage.filters import gaussian_filter
import dill

import matplotlib.pylab as pltparam
pltparam.rcParams["figure.figsize"] = (8,3)
pltparam.rcParams['lines.linewidth'] = 2
# pltparam.rcParams['lines.color'] = 'r'
pltparam.rcParams['axes.grid'] = True 

element_dump_filename = 'v09_dump_jet'

with open(element_dump_filename,'rb') as elements_dump:
    lEvent = dill.load(elements_dump)


RR = np.array([0.75, 1.5, 3, 6, 12, 24, 48, 96])
DD = np.zeros(RR.size)

EqRR = {} #'eq' as key ans 5x1 arrays as value
for e in lEvent:
    pv = e.Study_Folder_Pv
    if pv in EqRR.keys():
        if "NE" in e.Hole:
            EqRR[pv][0,0] = e.Discharge.ReleaseRate
            EqRR[pv][0,1] = e.Holemm*e.Holemm
        if "SM" in e.Hole:
            EqRR[pv][1,0] = e.Discharge.ReleaseRate
            EqRR[pv][1,1] = e.Holemm*e.Holemm
        if "ME" in e.Hole:
            EqRR[pv][2,0] = e.Discharge.ReleaseRate
            EqRR[pv][2,1] = e.Holemm*e.Holemm
        if "MA" in e.Hole:
            EqRR[pv][3,0] = e.Discharge.ReleaseRate
            EqRR[pv][3,1] = e.Holemm*e.Holemm
        if "FBR" in e.Hole:
            EqRR[pv][4,0] = e.Discharge.ReleaseRate
            EqRR[pv][4,1] = e.Holemm*e.Holemm
    else:
        EqRR[pv] = np.zeros([5,2])

RRFit = {}

for e in EqRR:
    RRFit[e] = np.zeros([8,1])
    # Afit = np.zeros(len(RR),2)
for e in EqRR:
    pv = e
    if pv in RRFit.keys():
        Ain = EqRR[e] # Col 0 for Release rate, Col 1 for Hole size*Hole size  (dd)        
        drrfit = interp1d(Ain[:,0],Ain[:,1],kind='linear')    
        ri = 0
        for r in RR:
            if r <  Ain[0,0]:
                r4 = Ain[1,0]
                r3 = Ain[0,0]
                dd4 = Ain[1,1]
                dd3 = Ain[0,1]
                dd = dd4+(dd4-dd3)/(r4-r3)*(r-r4)            
            elif r > Ain[4,0]:
                r4 = Ain[4,0]
                r3 = Ain[3,0]
                dd4 = Ain[4,1]
                dd3 = Ain[3,1]
                dd = dd4+(dd4-dd3)/(r4-r3)*(r-r4)
            else:
                dd = drrfit(r)
            # print("{:40s} {:8.2f} {:8.2f}".format(e, r, dd))
            RRFit[pv][ri,0] = dd
            ri += 1
for e in EqRR:
    print("{:40s} | {:8.2f} | {:8.2f} | {:8.2f} | {:8.2f} |\
        {:8.2f} | {:8.2f} | {:8.2f} | {:8.2f} | {:8.2f} | {:8.2f} | {:8.2f} | {:8.2f}".format(e, X[e], Y[e], ReleaseHeight[y], PP[e], \
        RRFit[e][0,0],RRFit[e][1,0],RRFit[e][2,0],RRFit[e][3,0],RRFit[e][4,0],RRFit[e][5,0],RRFit[e][6,0],RRFit[e][7,0]))    
