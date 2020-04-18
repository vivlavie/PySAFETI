f = open('rater_vis.txt')
scns = []
cutoffrate,numscn,dmy = f.readline().split(";")
numscn = int(numscn)
scns = f.readline().split("; ")
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

scns_to_plot = scns[14:15]
import matplotlib.pyplot as plt
fig,axes = plt.subplots(12,4)
# for i in range(0,numscn):
# for i in range(0,1):
# for i in range(1,2):
for s in scns_to_plot:
    for i in range(0,len(scns)):
        if scns[i] == s:
            break
    for j in range(0,numLeakRates):
        Tend = len(TVD[i*numLeakRates+j][0])
        axes[int(j/4),j%4].plot(TVD[i*numLeakRates+j][0])
        axes[int(j/4),j%4].set_title("{:s}{:6.1f}[kg/s]{:5d}[sec]".format(scns[i],TVD[i*numLeakRates+j][0][0],Tend))
        # axes[int(j/4),j%4].