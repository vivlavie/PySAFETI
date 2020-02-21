
print("Path          Module (     X,     Y,     Z)  Hole   Wind  P[barg]     T[degC]                         Disp  Jet  EPF  LFP  Exp  Fireball")
for e in lEvent:    
    ts,pv,IS,hole = e.Path.split("\\")
    f = "{:15s} {:4s} ({:6.1f},{:6.1f},{:6.1f}) {:6.1f} {:6s} P({:8.2f}) T({:8.2f}) ESD({:}) BDV({:}) ".\
                format(IS, e.Module, e.X, e.Y,  e.ReleaseHeight, e.Holemm, e.Weather[9:],e.Pressure, e.Temperature, e.ESDSuccess,e.BDVSuccess)
    if e.Dispersion != None:
        f += "  O  "
    else:
        f += "  X  "
    if e.JetFire != None:
        f += "{:5.1f}".e.JetFire.Length
    else:
        f += "  X  "
    if e.EarlyPoolFire != None:
        f += "  O  "
    else:
        f += "  X  "
    if e.LatePoolFire != None:
        f += "  O  "
    else:
        f += "  X  "
    if e.Explosion != None:
        f += "  O  "
    else:
        f += "  X  "
    if e.Fireball != None:
        f += "  O  "
    else:
        f += "  X  "
    print(f)
    

    
fexpmeoh=0
MeOH_Exp = []
for e in lEvent:
    if ("046-01" in e.Path) & (e.Explosion != None):
        ts,pv,MeohIS,hole = e.Path.split("\\")
        hole_name,tlv = hole.split("_")
        fexpmeoh += e.Explosion.Frequency
        # MeOH_Exp.append([MeohIS,hole_name,e.Explosion.Distance2OPs[-1],e.Explosion.Frequency,fexpmeoh])
        MeOH_Exp.append([MeohIS,hole_name,e.Explosion.Distance2OPs[-1],e.Explosion.Frequency])
        # print("{:12s} {:2s}  {:8.1f} F:{:6.2e} CF:{:6.2e}".format(MeohIS,hole_name,e.Explosion.Distance2OPs[-1],e.Explosion.Frequency,fexpmeoh))
#fexpmeho ~ 1.37E-4/year
MeOH_Exp_Sorted = sorted(MeOH_Exp, key = lambda fl: fl[2]) #with the longest duration at the bottom


def EventTree(e):
    # Pwss = 1/8*1/2 #8 wind directions & 2 wind velocities
    Pwss = 1
    Pi = e.PImdIgn # Probability of immediate ignition #input into SAFETI hole        
    Pdi = e.PDelIgn

    Ph = 4/6 #horizontal fraction for vertical and horizontal release cases
    Pj_h = 1/4*Ph # probability of horizontal jet fire Pjh x ph = 4/6 x 1/4 = 1/6 with no accompanying pool fire
    Pp_h = 1/2*Ph #probability of pool fire upon horizontal release with no accompanying jet fire
    Pjp_h =0.5/4*Ph #probability of jet & pool fire upon horizontal release

    Pj_v = 1/2*(1-Ph)#probabilty of vertical jet fire upon vertical release
    Pp_v = 1/2*(1-Ph)#
    Pjp_v = 0.5/4*(1-Ph)#

    #short release
    Ps = 1 #fraction for short release effects; release is shorter than cut-off time '20 s'
    Pbp = 1 #prob of bleve & pool fire
    Pb = 0 #prob of bleve without a pool fire
    Pfp = 0 #prob for short duration release immediate ignition resulting in flash fire and pool
    Pf = 0 #Flash fire along
    Pp = 0#prob for short duration release immediate ignition resulting in  to have pool fire
    Pep = 0#prob for short duration release immediate ignition resulting in  to have immediate explosion with pool fire
    Pe = 0#prob for short duration release immediate ignition resulting in  to have immediate explosion without pool fire
    Pp = 0#prob for short duration release immediate ignition resulting in  to have early pool fire without explosion


    Prp = 0.15 #prob of residual pool fire
    Pirp = 0 #Conditional probability of ignition of a residual pool fire??? Residual probability of non-ignition at the time when the cloud leaves the pool
    Po = (1-Pi)*Pwss*Pirp*Prp#prob of igniton of residual pool (MPACT theory p.54, p.140 Descriptions are different, Exception 5)
    Pf_d = 0.6 #prob of delayed flash fire
    Pe_d = 0.4 #prob of explosion upon delayed ignition

    #Immediate, Not short
    P_i_ns_h_j = (1-Ps)*Ph*Pwss*Pj_h*Pi
    P_i_ns_h_p = (1-Ps)*Ph*Pwss*Pp_h*Pi
    P_i_ns_h_jp = (1-Ps)*Ph*Pwss*Pjp_h*Pi
    P_i_ns_v_j = (1-Ps)*(1-Ph)*Pwss*Pj_v*Pi
    P_i_ns_v_p = (1-Ps)*(1-Ph)*Pwss*Pp_v*Pi
    P_i_ns_v_jp = (1-Ps)*(1-Ph)*Pwss*Pjp_v*Pi
    #Immediate, Short
    P_i_s_BP = Ps*Pwss*Pbp*Pi #BLEVE & Pool
    P_i_s_B = Ps*Pwss*Pb*Pi #BLEVE only
    P_i_s_FP = Ps*Pwss*Pfp*Pi #Flash fire and pool
    P_i_s_F = Ps*Pwss*Pf*Pi #Flash fire only
    P_i_s_EP = Ps*Pwss*Pep*Pi #Explosion & pool fire
    P_i_s_E = Ps*Pwss*Pe*Pi #Explosion only
    P_i_s_P = Ps*Pwss*Pp*Pi #Pool fire only

    #Delayed
    P_d_FP = (1-Pi)*Pwss*Pfp*Pdi #Delayed flash and pool fire
    P_d_E = (1-Pi)*Pwss*Pe*Pdi #Delayed explosion
    P_d_p = (1-Pi)*Pwss*Prp*Pirp #Redisual pool fire

    print("----Imd Ignition---Not short--|-Horizontal---Jet fire       : {:8.2e}".format(P_i_ns_h_j))
    print("                 |             |             |-Pool fire      : {:8.2e}".format(P_i_ns_h_p))
    print("                 |             |             -Jet & Pool fire: {:8.2e}".format(P_i_ns_h_jp))
    print("                 |             |-Vertical-----Jet fire       : {:8.2e}".format(P_i_ns_v_j))
    print("                 |                           |-Pool fire      : {:8.2e}".format(P_i_ns_v_p))
    print("                 |                           -Jet & Pool fire: {:8.2e}".format(P_i_ns_v_jp))
    print("                 --Short----------------------BLEVE & Poolfire: {:8.2e}".format(P_i_s_BP))
    print("                                            |-BLEVE only      : {:8.2e}".format(P_i_s_B))
    print("                                            |-Flash & Pool fire:{:8.2e}".format(P_i_s_FP))
    print("                                            |-Flash only: {:8.2e}".format(P_i_s_F))
    print("----Delayed-------Ignited--------------------Flash and Pool fire:{:8.2e}".format(P_d_FP))
    print("                                            |- Explosion : {:8.2e}".format(P_d_E))
    print("                                            |- Pool fire: {:8.2e}".format(P_d_p))