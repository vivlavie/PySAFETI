#Print Jet fire intime

print("{:70s} {:10s}{:10s}{:10s}{:10s}{:10s}".format("Scenario","T=0","T=5min","T=10min","T=30min","T=60min"))
for key in J00.keys():
    print("{:70s} {:10.2f}{:10.2f}{:10.2f}{:10.2f}{:10.2f}".format(key, J00[key], J05[key],J10[key],J30[key],J60[key]))

for key in J00.keys():
    if "Flow line" in key:
        print("{:70s} {:10.2f}{:10.2f}{:10.2f}{:10.2f}{:10.2f}".format(key, J00[key], J05[key],J10[key],J30[key],J60[key]))
for e in lEvent:
    if (("Flow line" in e.Path) and ("NE_TLV" in e.Path)):
        print(e.Path, e.Weather)
        print(e.Discharge.Duration)
        print(e.Discharge.Ts)
        print(e.Discharge.Ms)
        print(e.Discharge.RRs)
        print(e.TVD[-1,:])


for e in lEvent:
    if ("S04" in e.Module):
        #print(e.Path[30:], e.Weather, e.Discharge.RRs)
        """ print(e.Discharge.Duration)
        print(e.Discharge.Ts)
        print(e.Discharge.Ms)
        print(e.Discharge.RRs)
        print(e.TVD[-1,:]) """
        #Print Jet Fire
        j00,j05,j10,j30,j60 = e.JetFire.JetLengths
        fmt = "{:50s} | {:.2e} | {:6.0f} | {:8.2f} | {:8.2f} | {:8.2f} | {:8.2f}| {:8.2f}".\
            format(e.Key, e.JetFire.Frequency, e.JetFire.SEP, j00,j05,j10,j30,j60)
        print(fmt)
        

#Print sorted Jet length
for j in Js:
    system,pressure_vessel,eqtag,hole,weather = j[0].split()
    print("{:30s} {:.2e} {:6.1f}".format(eqtag,j[1],j[2]))


for mod in ['S04']:
#For each time
    #for J in [J00, J05, J10, J30, J60]:
    for J in [J05]:
        """ lEq = []
        lJF = []
        lJL = [] """
        jd = []
        for e in lEvent:
            if e.Module == mod:
                key = e.Path+"\\"+e.Weather
                """ lEq.append(key)
                lJF.append(JetFrequency[key])
                lJL.append(J[key]) """
                jd.append([key, JetFrequency[key], J[key]])
#Read the exceedance curve


#Print out Excd curve
 for j in Js:
     ...:     print("{:8.2e} {:8.2e}".format(j[1], j[2]))