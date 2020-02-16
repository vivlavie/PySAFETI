f = open('Hull_Discharge.csv','w')

for e in lEvent:
    ts,pv,eq,hole,wtr = e.Key.split("\\")
    if (e.Discharge != None) and (e.TVD != []):
        tvd = e.TVD
        rr = e.Discharge.ReleaseRate
        spilt_mass = rr*e.Discharge.Duration
        fmt = "{:8s} | {:20s} | {:10s} | {:6.1f} | {:8.1f} | {:8.0f} | {:8.1f}\n".\
            format(e.Module, eq, hole, e.Discharge.ReleaseRate, e.Discharge.Duration, spilt_mass, e.TVD[-1,0])
        # print(fmt)
    else:
        fmt = "{:8s} | {:20s} | {:10s} |\n".\
            format(e.Module, eq,hole,wtr)
    f.write(fmt)
f.close()