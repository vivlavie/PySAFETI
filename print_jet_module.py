f = open('JetFireTop.csv','w')

for e in lEvent:
    ts,pv,eq,hole,wtr = e.Key.split("\\")
    if e.JetFire != None:
        j00,j05,j10,j30,j60 = e.JetFire.JetLengths
        fmt = "{:8s} | {:20s} | {:10s} | {:13s} | {:.2e} | {:6.0f} | {:8.2f} | {:8.2f} | {:8.2f} | {:8.2f}| {:8.2f}\n".\
            format(e.Module, eq,hole,wtr, e.JetFire.Frequency, e.JetFire.SEP, j00,j05,j10,j30,j60)
        # print(fmt)
    else:
        fmt = "{:8s} | {:20s} | {:10s} |\n".\
            format(e.Module, eq,hole,wtr)
    f.write(fmt)
f.close()