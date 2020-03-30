import os

NumDirections = 6

def EventTree_OLF2007(pv,hole,weather,lEvent):
    import shutil
    #Failure of fire detector for immediate ignition
    P_FD_Fail = 0.0
    #Failure of gas detectin for delayed ignition
    P_GD_Fail = 0.0
    et_filename = 'ET_'+pv+'_'+hole+'_'+weather+'.xlsx'
    if et_filename in os.listdir():
        iExl=load_workbook(filename=et_filename)
    else:
        shutil.copy('ET_template.xlsx',et_filename)
        iExl=load_workbook(filename=et_filename)    
    
    shET = iExl['et']
    basic_freq_set = False
    for e in lEvent:

        if ((pv in e.Key) and (hole in e.Hole) and (weather == e.Weather)):
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
            epv,ehole,eweather = e.Key.split("\\")
            print("{:20s}{:10s} {:6s} {:8.2e} {:8.2e} {:8.2e} - Leak Freq {:8.2e} ".format(epv,ehole,eweather,e.Frequency,e.PESD,e.PBDV,e.Frequency/e.PESD/e.PBDV))
            if basic_freq_set == False:
                shET.cell(1,6).value = pv
                shET.cell(1,7).value = hole
                shET.cell(1,8).value = weather
                shET.cell(30,2).value = e.Frequency/e.PESD/e.PBDV
                shET.cell(31,2).value = e.Discharge.ReleaseRate
                shET.cell(20,6).value = e.PImdIgn
                shET.cell(26,10).value = e.PDelIgn
                shET.cell(15,7).value = 1-P_FD_Fail
                shET.cell(42,7).value = 1-P_GD_Fail
                # shET.cell(10,8).value = 1-0.01*numESDVs[pv]
                shET.cell(10,8).value = e.PESD #PESD: p of successful ESD
                # if "BN" in e.Hole:
                #     shET.cell(9,9).value = 1
                # else:
                #     shET.cell(9,9).value = 1-0.005
                shET.cell(9,9).value = e.PBDV #PBDV: p of successful BDV


                basic_freq_set = True
            else:
                if abs(shET.cell(30,2).value - e.Frequency/e.PESD/e.PBDV)/(e.Frequency/e.PESD/e.PBDV) > 0.01:
                    print(pv, epv, ehole,'Basic frequency wrong',shET.cell(30,2).value, e.Frequency/e.PESD/e.PBDV, e.PESD, e.PBDV)
                    # break
                # else:
                #     print(epv,hole,'Basic frequency match')
                    
                if shET.cell(20,6).value != e.PImdIgn:
                    print(epv, hole, 'ImdIgn Wrong')
                    # break
                # else:
                #     print(epv,hole,'ImdIgn match')
                    
            if e.Explosion == None:
                e.Explosion = Explosion(0.,0.,0.)
            #Eithter fire detection for jet fire or gas detection for explosion has succedded
            if ("EOBO" in ehole) or ("EOBN" in ehole):
                shET.cell(9,13).value = e.JetFire.Frequency*NumDirections*(1-P_FD_Fail) #Jet fire frequency = LeakFreq x PEO x PBO x PImgIgn / NumDir
                shET.cell(25,13).value = e.Explosion.Frequency*(1-P_GD_Fail) #VCE                
                # shET.cell(28,13).value = e.Frequency*(1-e.PImdIgn)*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                # shET.cell(31,13).value = e.Frequency*(1-e.PImdIgn)*(1-e.PDelIgn)*(1-P_GD_Fail) #Gas disperison
                shET.cell(28,13).value = e.Frequency*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                shET.cell(31,13).value = e.Frequency*(1-e.PDelIgn)*(1-P_GD_Fail) #Gas disperison
                # print(pv,epv,ehole,shET.cell(9,13).value, shET.cell(25,13).value, shET.cell(31,13).value)
            if "EOBX" in ehole:
                shET.cell(12,13).value = e.JetFire.Frequency*NumDirections*(1-P_FD_Fail)
                shET.cell(34,13).value = e.Explosion.Frequency*(1-P_GD_Fail)                
                # shET.cell(37,13).value = e.Frequency*(1-e.PImdIgn)*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                # shET.cell(40,13).value = e.Frequency*(1-e.PImdIgn)*(1-e.PDelIgn)*(1-P_GD_Fail)
                shET.cell(37,13).value = e.Frequency*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                shET.cell(40,13).value = e.Frequency*(1-e.PDelIgn)*(1-P_GD_Fail)
                # print(pv,epv,ehole,shET.cell(12,13).value, shET.cell(34,13).value, shET.cell(37,13).value)
            if ("EXBO" in ehole) or ("EXBN" in ehole):
                shET.cell(15,13).value = e.JetFire.Frequency*NumDirections*(1-P_FD_Fail)
                shET.cell(43,13).value = e.Explosion.Frequency*(1-P_GD_Fail)
                # shET.cell(46,13).value = e.Frequency*(1-e.PImdIgn)*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                # shET.cell(49,13).value = e.Frequency*(1-e.PImdIgn)*(1-e.PDelIgn)*(1-P_GD_Fail)
                shET.cell(46,13).value = e.Frequency*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                shET.cell(49,13).value = e.Frequency*(1-e.PDelIgn)*(1-P_GD_Fail)
                # print(pv,epv,ehole,shET.cell(15,13).value, shET.cell(43,13).value, shET.cell(49,13).value)
            if "EXBX" in ehole:
                #Two cases should be distinguished; either detectors are successful or no
                #Detection Success but EXBX
                shET.cell(18,13).value = e.JetFire.Frequency*NumDirections*(1-P_FD_Fail)
                shET.cell(52,13).value = e.Explosion.Frequency*(1-P_GD_Fail) #explosion
                # shET.cell(55,13).value = e.Frequency*(1-e.PImdIgn)*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                # shET.cell(58,13).value = e.Frequency*(1-e.PImdIgn)*(1-e.PDelIgn)*(1-P_GD_Fail) #Dispersion only                
                shET.cell(55,13).value = e.Frequency*(e.PDelIgn)*(1-e.PExp_Ign)*(1-P_GD_Fail) #Flash fire
                shET.cell(58,13).value = e.Frequency*(1-e.PDelIgn)*(1-P_GD_Fail) #Dispersion only                

                #Detection failure & EXBX
                shET.cell(22,13).value = e.JetFire.Frequency*NumDirections/e.PESD/e.PBDV*(P_FD_Fail)                
                shET.cell(61,13).value = e.Explosion.Frequency/e.PESD/e.PBDV*(P_GD_Fail)
                # shET.cell(64,13).value = e.Frequency/e.PESD/e.PBDV*(1-e.PImdIgn)*(e.PDelIgn)*(1-e.PExp_Ign)*(P_GD_Fail) #Flash fire
                # shET.cell(67,13).value = e.Frequency/e.PESD/e.PBDV*(1-e.PImdIgn)*(1-e.PDelIgn)*(P_GD_Fail)
                shET.cell(64,13).value = e.Frequency/e.PESD/e.PBDV*(e.PDelIgn)*(1-e.PExp_Ign)*(P_GD_Fail) #Flash fire
                shET.cell(67,13).value = e.Frequency/e.PESD/e.PBDV*(1-e.PDelIgn)*(P_GD_Fail)
                
                # print(pv,epv,ehole,shET.cell(22,13).value, shET.cell(61,13).value, shET.cell(66,13).value)


    iExl.save(et_filename)
    

                

    
    








def EventTree_SAFETI(e):
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