from __future__ import division
import os, sys, math, re
from PyPDF2 import PdfFileMerger
import psse34
import psspy
import redirect
import shutil
_i = psspy.getdefaultint()
_f = psspy.getdefaultreal()
_s = psspy.getdefaultchar()
MisMatch_Tol = 0.001
cwd = os.getcwd()
Plotting_tools = os.path.join(os.path.dirname(cwd),'Plotting_tools')
sys.path.append(Plotting_tools)
import Plotting as pl
# ! Dynamic Parameters -----------------------------------
nprt = 50000                # number of time steps between the printing of the channel values (input; unchanged).
nplt = 50                  # number of time steps between the writing of the output channel values to the current channel output file (input; unchanged).
crtplt = 50                 # number of time steps between the plotting of those channel values that have been designated as CRT output channels (input; unchanged).
# ========================================================
def Initial(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'Initial'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info,org_sav_file)
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' %os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (test_cases['Case No'][i], Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                print 'v_ref_pu = ', v_ref_pu
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW, Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW, Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            report_file('%s_docu.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            runtime = 20
            ierr = psspy.docu(0, 1, [0, 3, 1])
            psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def ADT(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'ADT' #Angle disturbance test
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                print(str(test_cases['Name 2'][i]).lower())
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            network_impedance_line = Project_info['network_impedance_line']
            ierr, Zsys = psspy.brndt2(network_impedance_line[0], network_impedance_line[1], network_impedance_line[2], 'RX')
            psspy.ltap(network_impedance_line[0],network_impedance_line[1],network_impedance_line[2], 0.1,888,r"""Dummy_TRF""", _f)
            psspy.purgbrn(network_impedance_line[0],888,network_impedance_line[2])
            psspy.two_winding_data_3(network_impedance_line[0],888,network_impedance_line[2],[1,network_impedance_line[0],1,0,0,0,33,0,network_impedance_line[0],0,1,0,1,1,1],[0.0, 0.0001, 100.0, 1.0,0.0,0.0, 1.0,0.0,0.0,0.0,0.0, 1.0, 1.0, 1.0, 1.0,0.0,0.0, 1.1, 0.9, 1.1, 0.9,0.0,0.0,0.0],r"""IDTRF""")
            psspy.branch_data(888, network_impedance_line[1],network_impedance_line[2], realar1=Zsys.real, realar2=Zsys.imag)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            refs = []
            for ref in re.split('[ :,;]', test_cases['Ref 1'][i]):
                refs.append(float(ref))
            runtime = refs[0]
            psspy.run(0, runtime, nprt, nplt, crtplt)
            for j in range(1,10):
                if refs[j] >0:
                    psspy.two_winding_data_3(network_impedance_line[0],888,network_impedance_line[2],realari6 = refs[j+10])
                    runtime = refs[j]
                    psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def APT(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'APT'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_icon(M_ppc + 1, 1)
            psspy.change_cctbusomod_icon(600, r"""VESTAS_PPCV8400""", 2, 1)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            refs = []
            for ref in (re.split('[ :,;]', test_cases['Ref 1'][i])):
                refs.append(float(ref))
            runtime = refs[0]
            psspy.run(0, runtime, nprt, nplt, crtplt)
            for j in range(1,10):
                if refs[j] >0:
                    if 'VWRE_J' in Project_info:
                        psspy.change_var(L_ppc_VWPO + 2, refs[j+10])
                    else:
                        psspy.change_var(L_ppc + 0, refs[j + 10])
                    runtime = refs[j]
                    psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def Fault_SMIB(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'Fault_SMIB'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    POC_dummy_line = Project_info['ppc_line']
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                print(str(test_cases['Name 2'][i]).lower())
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            network_impedance_line = Project_info['network_impedance_line']
            ierr, Zsys1 = psspy.brndt2(network_impedance_line[0], network_impedance_line[1], network_impedance_line[2], 'RX')
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
            else:
                J_ppc, M_ppc, L_ppc, K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
            start(new_sav_file, test_cases, i, vestas_dlls, add_dlls)
            refs = []
            for ref in (re.split('[ :,;]', test_cases['Ref 1'][i])):
                refs.append(float(ref))
            runtime = refs[0]
            psspy.run(0, runtime, nprt, nplt, crtplt)
            for j in range(1,10):
                if refs[j] > 0:
                    if refs[j + 10]<1:
                        Vflt = refs[j + 10]
                        if Vflt == 0:
                            Vflt = 0.01
                        SBASE = 100
                        fault_B = ((1 - Vflt) / (Vflt * Zsys1)).imag * SBASE
                        fault_G = ((1 - Vflt) / (Vflt * Zsys1)).real * SBASE
                        psspy.dist_bus_fault(POC_dummy_line[1], 1, 0.0, [fault_G, fault_B])
                        runtime = refs[j]
                        psspy.run(0, runtime, nprt, nplt, crtplt)
                    elif refs[j + 10]==1:
                        psspy.dist_clear_fault(1)
                        runtime = refs[j]
                        psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def RPT(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'RPT'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    wfbase_MW = 336
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
                psspy.change_icon(M_ppc_VWRE + 0, 1)
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_icon(M_ppc + 2, 0)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            refs = []
            for ref in re.split('[ :,;]', test_cases['Ref 1'][i]):
                refs.append(float(ref))
            runtime = refs[0]
            psspy.run(0, runtime, nprt, nplt, crtplt)
            for j in range(1,10):
                if refs[j] > 0:
                    if 'VWRE_J' in Project_info:
                        psspy.change_var(L_ppc_VWRE + 0, refs[j+10])
                    else:
                        psspy.change_var(L_ppc + 1, refs[j + 10])
                    runtime = refs[j]
                    psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def PFT(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'PFT'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
                psspy.change_icon(M_ppc_VWRE + 0, 2)  # Power Factor Control mode: 2
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_icon(M_ppc + 2, 1)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            refs = []
            for ref in re.split('[ :,;]', test_cases['Ref 1'][i]):
                refs.append(float(ref))
            runtime = refs[0]
            psspy.run(0, runtime, nprt, nplt, crtplt)
            for j in range(1,10):
                if refs[j] > 0:
                    if 'VWRE_J' in Project_info:
                        psspy.change_var(L_ppc_VWRE + 3, refs[j+10])
                    else:
                        psspy.change_var(L_ppc + 3, refs[j + 10])
                    runtime = refs[j]
                    psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def VCT(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'VCT'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_icon(M_ppc + 2, 2)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            refs = []
            for ref in re.split('[ :,;]', test_cases['Ref 1'][i]):
                refs.append(float(ref))
            runtime = refs[0]
            psspy.run(0, runtime, nprt, nplt, crtplt)
            for j in range(1,10):
                if refs[j] > 0:
                    if 'VWRE_J' in Project_info:
                        psspy.change_var(L_ppc_VWRE + 4, refs[j+10])
                    else:
                        psspy.change_var(L_ppc + 2, refs[j + 10])
                    runtime = refs[j]
                    psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def VDT(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'VDT'
    output_files_dir = os.path.join(results, Test)
    dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            refs = []
            for ref in re.split('[ :,;]', test_cases['Ref 1'][i]):
                refs.append(float(ref))
            dyr_file = os.path.join(os.path.dirname(org_sav_file), 'FCT Dyr Dll\\Case_%02d.dyr'%i)
            dyrfile = open(dyr_file, "w")
            dyrfile.write(r"""%d, 'USRMDL', %d, 'PLBVFU1', 1, 1, 3, 4, 3, 6, 1, 0, 'Case_%02d', 1, 50, 0, 0/""" % (Project_info['SMIB'][0], Project_info['SMIB'][1],i))
            dyrfile.write("\n")
            dyrfile.close()
            plb_file = os.path.join(os.path.dirname(org_sav_file), 'FCT Dyr Dll\\Case_%02d.plb'%i)
            plbfile = open(plb_file, "w")
            for j in range(10):
                if (refs[j] > 0 or j == 0):
                    plbfile.write("{: <8}".format(refs[j]))
                    plbfile.write("{: <8}".format(refs[j+10]))
                    plbfile.write("50.0")
                    if refs[j+1] >= refs[j] and j<>9:
                        plbfile.write("\n")
            runtime = max(refs[0:10])
            plbfile.write("\n")
            plbfile.close()
            new_add_dyrs = [x for x in add_dyrs if "GENCLS" not in x]
            new_add_dyrs.append(dyr_file)
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,new_add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,new_add_dyrs,vestas_dlls,add_dlls)
            copy_list([plb_file], os.path.dirname(new_sav_file))
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def FCT_PLB(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'FCT_PLB'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            if 'VWRE_J' in Project_info:
                if int(test_cases['Ref 2'][i]) == 2:
                    Pctrl_mode = 'FSM'
                elif int(test_cases['Ref 2'][i]) == 1:
                    Pctrl_mode = 'PCtrl'
            else:
                if int(test_cases['Ref 2'][i]) == 1:
                    Pctrl_mode = 'FSM'
                elif int(test_cases['Ref 2'][i]) == 0:
                    Pctrl_mode = 'PCtrl'
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_%s_PSSE' % (test_cases['Case No'][i], Test,Pctrl_mode))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s_%s' % (test_cases['Case No'][i], Test,Pctrl_mode))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            refs = []
            for ref in re.split('[ :,;]', test_cases['Ref 1'][i]):
                refs.append(float(ref))
            dyr_file = os.path.join(os.path.dirname(org_sav_file), 'FCT Dyr Dll\\Case_%02d.dyr'%i)
            dyrfile = open(dyr_file, "w")
            dyrfile.write(r"""%d 'USRMDL' %d 'PLBVFU1', 1, 1, 3, 4, 3, 6, -1,1,'Case_%02d',1,50,0,0 /""" % (Project_info['SMIB'][0], Project_info['SMIB'][1],i))
            dyrfile.write("\n")
            dyrfile.close()
            plb_file = os.path.join(os.path.dirname(org_sav_file), 'FCT Dyr Dll\\Case_%02d.plb'%i)
            plbfile = open(plb_file, "w")
            for j in range(10):
                if (refs[j] > 0 or j == 0):
                    plbfile.write("{: <8}".format(refs[j]))
                    plbfile.write("{: <8}".format(1))
                    plbfile.write('%s' %refs[j+10])
                    if refs[j+1] >= refs[j] and j<>9:
                        plbfile.write("\n")
            plbfile.write("\n")
            plbfile.close()
            new_add_dyrs = [x for x in add_dyrs if "GENCLS" not in x]
            new_add_dyrs.append(dyr_file)
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,new_add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
                psspy.change_icon(M_ppc_VWPO + 0, int(test_cases['Ref 2'][i]))  # Change Control model
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,new_add_dyrs,vestas_dlls,add_dlls)
            if 'VWRE_J' in Project_info:
                psspy.change_icon(M_ppc_VWPO + 0, int(test_cases['Ref 2'][i]))
            else:
                psspy.change_icon(M_ppc+1, int(test_cases['Ref 2'][i]))
            copy_list([plb_file], os.path.dirname(new_sav_file))
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)

            runtime = 20
            psspy.run(0, runtime, nprt, nplt, crtplt)
            if 'VWRE_J' in Project_info:
                if test_cases['Ref 3'][i] <1 and int(test_cases['Ref 2'][i]) == 1:
                    psspy.change_var(L_ppc_VWPO + 2, test_cases['Ref 3'][i])
                elif test_cases['Ref 3'][i] <1 and int(test_cases['Ref 2'][i]) == 2:
                    psspy.change_var(L_ppc_VWPO + 0, test_cases['Ref 3'][i])
            else:
                if test_cases['Ref 3'][i] <1:
                    psspy.change_var(L_ppc + 0, test_cases['Ref 3'][i])
            runtime = max(refs[0:10])
            psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def FRB(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'FRB'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            refs = []
            for ref in re.split('[ :,;]', test_cases['Ref 1'][i]):
                refs.append(float(ref))
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            runtime = refs[0]
            psspy.run(0, runtime, nprt, nplt, crtplt)
            for j in range(1,10):
                if refs[j] > 0:
                    if 'VWRE_J' in Project_info:
                        psspy.change_var(L_ppc_VWPO + 90, refs[j+10])
                    else:
                        print('FRB function have not been implemented in new PSSE V8 model')
                        # psspy.change_var(L_ppc + , refs[j + 10])
                    runtime = refs[j]
                    psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def WTGtrip(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'WTGtrip'
    wtg_buses = Project_info['wtg_buses']
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    for wtg_bus in wtg_buses:
        for i in range(len(test_cases['Case No'])):
            if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
                new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_%d_PSSE' % (i, Test,wtg_bus[0]))
                if overlay != '':
                    check_path(overlay)
                    overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                    check_path(overlay_dir)
                scr = test_cases['SCR'][i]
                xr = test_cases['XR'][i]
                p_poc_MW = test_cases['Initial P (MW)'][i]
                Q_POC = test_cases['Initial Q (MVAr)'][i]
                V_POC = test_cases['V_actual_POC (pu)'][i]
                v_ref_pu = test_cases['Vref_droop (pu)'][i]
                if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                    adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
                elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                    adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
                else:
                    shutil.copy(org_sav_file, output_files_dir)
                    new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
                psspy.psseinit(80000)
                log_file('%s_1.dat' % new_sav_file[0:-4])
                psspy.case(new_sav_file)
                if 'nem' in NEM_Conv_file.lower():
                    sys.path.append(os.path.dirname(NEM_Conv_file))
                    import NEM_Conv
                    NEM_Conv.NEM_convert()
                else:
                    Convert()
                if 'VWRE_J' in Project_info:
                    J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                    psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                    psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                    msu_caps = Project_info['msu_caps']
                    for msu_cap in msu_caps:
                        psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
                else:
                    J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
                runtime = 10
                psspy.run(0, runtime, nprt, nplt, crtplt)
                psspy.machine_chng_2(wtg_bus[0], wtg_bus[1], intgar1=0)
                runtime += 20
                psspy.run(0, runtime, nprt, nplt, crtplt)
                os.chdir(os.path.dirname(main_python))
                psspy.stoprecording()
                delete_first_line('%s.py' % new_sav_file[0:-4])
                psspy.close_powerflow()
                psspy.pssehalt_2()
                if overlay != '':
                    shutil.copy2('%s.out' % new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def FCT_K2(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'FCT_K2'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            if int(test_cases['Ref 2'][i]) == 2:
                Pctrl_mode = 'FSM'
            elif int(test_cases['Ref 2'][i]) == 1:
                Pctrl_mode = 'PCtrl'
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_%s_PSSE' % (i, Test,Pctrl_mode))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            add_dlls_new = add_dlls
            # add_dlls_new.append(os.path.join(os.path.dirname(vestas_dlls[0]),'dsusrusercode.dll'))
            add_dyrs.append(os.path.join(os.path.dirname(org_sav_file),'FCT Dyr Dll\\%s.dyr' %test_cases['Ref 1'][i]))
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls_new)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
                psspy.change_icon(M_ppc_VWPO + 0, int(test_cases['Ref 2'][i]))  # Change Control model to P control mode
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            runtime = 5
            psspy.run(0, runtime, nprt, nplt, crtplt)
            if test_cases['Ref 1'][i] in ['FCT_05', 'FCT_06', 'FCT_07', 'FCT_08'] and int(test_cases['Ref 2'][i]) == 1:
                psspy.change_var(L_ppc_VWPO + 2, 0.52153)
            elif test_cases['Ref 1'][i] in ['FCT_05', 'FCT_06', 'FCT_07', 'FCT_08'] and int(test_cases['Ref 2'][i]) == 2:
                psspy.change_var(L_ppc_VWPO + 0, 0.52153)
            runtime = float(test_cases['Ref 3'][i])
            psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            add_dyrs.pop(-1)
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def MSU_logic_Q(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'MSU_logic_Q'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    if 'VWRE_J' in Project_info:
        phs,qhs,dphs,dqhs = [],[],[],[]
        for i in range(6):
            if Project_info['VWRE_J'][i+28] <1 and Project_info['VWRE_J'][i+28]<> 0:
                phs.append(Project_info['VWRE_J'][i+28])
                qhs.append(Project_info['VWRE_J'][i + 34])
                dphs.append(Project_info['VWRE_J'][i + 40])
                dqhs.append(Project_info['VWRE_J'][i + 46])
    else:
        phs, qhs = [], []
        for i in range(10):
            if Project_info['PPC_V8_J'][i+233] <1 and Project_info['PPC_V8_J'][i+233]<>0:
                phs.append(Project_info['PPC_V8_J'][i+233])
                qhs.append(Project_info['PPC_V8_J'][i + 243])
        dph =Project_info['PPC_V8_J'][253]
        dqh =Project_info['PPC_V8_J'][254]
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                psspy.change_icon(M_ppc_VWPO + 37, 1)  ### enable P higher than initial condition
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
                psspy.change_icon(M_ppc_VWRE + 0, 1)  # Reactive power control mode 1
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_icon(M_ppc + 2, 0)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            runtime = 10
            psspy.run(0, runtime, nprt, nplt, crtplt)
            if 'VWRE_J' in Project_info:
                psspy.change_var(L_ppc_VWPO + 2, phs[-1] + dphs[-1] + 0.005)
                psspy.change_var(L_ppc_VWRE + 0, qhs[0] - dqhs[0] - 0.005)
                runtime += 20
                psspy.run(0, runtime, nprt, nplt, crtplt)
                for j in range(len(qhs)):
                    psspy.change_var(L_ppc_VWRE + 0, qhs[j] + dqhs[j] + 0.005)
                    runtime += 20
                    psspy.run(0, runtime, nprt, nplt, crtplt)
                for k in range(len(qhs)-1,-1,-1):
                    psspy.change_var(L_ppc_VWRE + 0, qhs[k] - dqhs[k] - 0.005)
                    runtime += 20
                    psspy.run(0, runtime, nprt, nplt, crtplt)
            else:
                psspy.change_var(L_ppc + 0, phs[-1] + dph + 0.005)
                psspy.change_var(L_ppc + 1, qhs[0] - dqh - 0.005)
                runtime += 20
                psspy.run(0, runtime, nprt, nplt, crtplt)
                for j in range(len(qhs)):
                    psspy.change_var(L_ppc + 1, qhs[j] + dqh + 0.005)
                    runtime += 20
                    psspy.run(0, runtime, nprt, nplt, crtplt)
                for k in range(len(qhs)-1,-1,-1):
                    psspy.change_var(L_ppc + 1, qhs[k] - dqh - 0.005)
                    runtime += 20
                    psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def MSU_logic_P(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'MSU_logic_P'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info, org_sav_file)
    if 'VWRE_J' in Project_info:
        phs,qhs,dphs,dqhs = [],[],[],[]
        for i in range(6):
            if Project_info['VWRE_J'][i+28] <1 and Project_info['VWRE_J'][i+28]<> 0:
                phs.append(Project_info['VWRE_J'][i+28])
                qhs.append(Project_info['VWRE_J'][i + 34])
                dphs.append(Project_info['VWRE_J'][i + 40])
                dqhs.append(Project_info['VWRE_J'][i + 46])
    else:
        phs, qhs = [], []
        for i in range(10):
            if Project_info['PPC_V8_J'][i+233] <1 and Project_info['PPC_V8_J'][i+233]<>0:
                phs.append(Project_info['PPC_V8_J'][i+233])
                qhs.append(Project_info['PPC_V8_J'][i + 243])
        dph =Project_info['PPC_V8_J'][253]
        dqh =Project_info['PPC_V8_J'][254]

    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' % os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            if overlay != '':
                check_path(overlay)
                overlay_dir = os.path.join(overlay, 'Case_%02d_%s' % (test_cases['Case No'][i], Test))
                check_path(overlay_dir)
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW,Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW,Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))
            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                psspy.change_icon(M_ppc_VWPO + 37, 1) ### enable P higher than initial condition
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
                psspy.change_icon(M_ppc_VWRE + 0, 1)  # Reactive power control mode 1
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_icon(M_ppc + 2, 0)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            runtime = 10
            psspy.run(0, runtime, nprt, nplt, crtplt)
            if 'VWRE_J' in Project_info:
                psspy.change_var(L_ppc_VWPO + 2, phs[0] - dphs[0] - 0.005)
                psspy.change_var(L_ppc_VWRE + 0, qhs[-1] + dqhs[-1] + 0.005)
                runtime += 40
                psspy.run(0, runtime, nprt, nplt, crtplt)
                for j in range(len(phs)):
                    psspy.change_var(L_ppc_VWPO + 2, phs[j] + dphs[j] + 0.005)
                    runtime += 20
                    psspy.run(0, runtime, nprt, nplt, crtplt)
                for k in range(len(phs)-1,-1,-1):
                    psspy.change_var(L_ppc_VWPO + 2, phs[k] - dphs[k] - 0.005)
                    runtime += 20
                    psspy.run(0, runtime, nprt, nplt, crtplt)
            else:
                psspy.change_var(L_ppc + 0, phs[0] - dph - 0.005)
                psspy.change_var(L_ppc + 1, qhs[-1] + dqh + 0.005)
                runtime += 40
                psspy.run(0, runtime, nprt, nplt, crtplt)
                for j in range(len(phs)):
                    psspy.change_var(L_ppc + 0, phs[j] + dph + 0.005)
                    runtime += 20
                    psspy.run(0, runtime, nprt, nplt, crtplt)
                for k in range(len(phs)-1,-1,-1):
                    psspy.change_var(L_ppc + 0, phs[k] - dph - 0.005)
                    runtime += 20
                    psspy.run(0, runtime, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
            if overlay != '':
                shutil.copy2('%s.out' %new_sav_file[0:-4], overlay_dir)
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
# ========================================================
def start(new_case, test_cases, i, vestas_dlls, add_dlls):
    os.chdir(os.path.dirname(new_case))
    if 'Time step (s)' in test_cases:
        if type(test_cases['Time step (s)'][i]) == float:
            psspy.dynamics_solution_param_2(realar3=float(test_cases['Time step (s)'][i]))
    if 'Acceleration factor' in test_cases:
        if type(test_cases['Acceleration factor'][i]) == float:
            psspy.dynamics_solution_param_2(realar1=float(test_cases['Acceleration factor'][i]))
    if 'frq_filter' in test_cases:
        if type(test_cases['frq_filter'][i]) == float:
            psspy.dynamics_solution_param_2(realar4=float(test_cases['frq_filter'][i]))
    psspy.save('%s_conv.sav' % new_case[0:-4])
    psspy.snap([-1, -1, -1, -1, -1], '%s' % new_case[0:-4])
    psspy.close_powerflow()
    psspy.pssehalt_2()
    copy_list(vestas_dlls, os.path.dirname(new_case))
    copy_list(add_dlls, os.path.dirname(new_case))
    psspy.psseinit(80000)
    log_file('%s_2.dat' % new_case[0:-4])
    psspy.startrecording(1, '%s.py' % new_case[0:-4])
    psspy.case('%s_conv.sav' % os.path.basename(new_case[0:-4]))
    psspy.rstr('%s' % os.path.basename(new_case[0:-4]))
    for dll in vestas_dlls:
        psspy.addmodellibrary(os.path.basename(dll))
    for dll in add_dlls:
        psspy.addmodellibrary(os.path.basename(dll))
    psspy.strt_2([0, 0], '%s.out' % os.path.basename(new_case[0:-4]))
# ========================================================
def run_LoadFlow():
    psspy.fnsl([1,0,0,1,1,0,99,0])
    psspy.fnsl([1,0,0,1,1,0,99,0])
    psspy.fnsl([1,0,0,1,1,0,99,0])
    psspy.fnsl([1,0,0,1,1,0,99,0])
    psspy.fnsl([1,0,0,1,1,0,99,0])   #Full Newton-Raphson
    psspy.fnsl([1,0,0,1,1,0,99,0])
    psspy.fnsl([1,0,0,1,1,0,99,0])
    blownup = psspy.solved()
    if blownup == 0:
        return 0
    else:
        return 1
# ========================================================
def check_wfbase_MW(Project_info,sav_file):
    psspy.psseinit(80000)
    psspy.case(sav_file)
    MBASEs = []
    for wtg_bus in Project_info['wtg_buses']:
        ierr, MBASE = psspy.macdat(wtg_bus[0], wtg_bus[1], 'MBASE')
        MBASEs.append(MBASE)
    wfbase_MW = sum(MBASEs)
    return wfbase_MW
    psspy.close_powerflow()
    psspy.pssehalt_2()
# ========================================================
def tune_capbank(p_poc_MW,q_poc_MVar,wfbase_MW,Project_info):
    if 'VWRE_J' in Project_info:
        ph1 =Project_info['VWRE_J'][28]
        ph2 =Project_info['VWRE_J'][29]
        ph3 =Project_info['VWRE_J'][30]
        ph4 =Project_info['VWRE_J'][31]
        ph5 =Project_info['VWRE_J'][32]
        ph6 =Project_info['VWRE_J'][33]
        qh1 =Project_info['VWRE_J'][34]
        qh2 =Project_info['VWRE_J'][35]
        qh3 =Project_info['VWRE_J'][36]
        qh4 =Project_info['VWRE_J'][37]
        qh5 =Project_info['VWRE_J'][38]
        qh6 =Project_info['VWRE_J'][39]
    elif 'PPC_V8_J' in Project_info:
        ph1 = Project_info['PPC_V8_J'][233]
        ph2 = Project_info['PPC_V8_J'][234]
        ph3 = Project_info['PPC_V8_J'][235]
        ph4 = Project_info['PPC_V8_J'][236]
        ph5 = Project_info['PPC_V8_J'][237]
        ph6 = Project_info['PPC_V8_J'][238]
        ph7 = Project_info['PPC_V8_J'][239]
        ph8 = Project_info['PPC_V8_J'][240]
        ph9 = Project_info['PPC_V8_J'][241]
        ph10 = Project_info['PPC_V8_J'][242]
        qh1 =Project_info['PPC_V8_J'][243]
        qh2 =Project_info['PPC_V8_J'][244]
        qh3 =Project_info['PPC_V8_J'][245]
        qh4 =Project_info['PPC_V8_J'][246]
        qh5 =Project_info['PPC_V8_J'][247]
        qh6 =Project_info['PPC_V8_J'][248]
        qh7 =Project_info['PPC_V8_J'][249]
        qh8 =Project_info['PPC_V8_J'][250]
        qh9 =Project_info['PPC_V8_J'][251]
        qh10 =Project_info['PPC_V8_J'][252]
    else:
        print('Incorrect project info')
    msu_caps = Project_info['msu_caps']
    p_poc_pu_wfbase = p_poc_MW/wfbase_MW
    q_poc_pu_wfbase = q_poc_MVar/wfbase_MW
    if p_poc_pu_wfbase < ph1 or q_poc_pu_wfbase <qh1:
        num_cap = 0
    elif p_poc_pu_wfbase < ph2 or q_poc_pu_wfbase <qh2:
        num_cap = 1
    elif p_poc_pu_wfbase < ph3 or q_poc_pu_wfbase <qh3:
        num_cap = 2
    elif p_poc_pu_wfbase < ph4 or q_poc_pu_wfbase <qh4:
        num_cap = 3
    elif p_poc_pu_wfbase < ph5 or q_poc_pu_wfbase <qh5:
        num_cap = 4
    elif p_poc_pu_wfbase < ph6 or q_poc_pu_wfbase <qh6:
        num_cap = 5
    else:
        num_cap = 6
    for msu_cap in msu_caps:
        psspy.switched_shunt_chng_3(msu_cap[0], intgar9=1, realar11=msu_cap[int(num_cap)+1])
    return(int(num_cap))
# ========================================================
def adjust_Q_and_V_POC_SMIB(case, new_case, scr, xr, P_POC, Q_POC, V_POC,wfbase_MW,Project_info):
    wtg_buses = Project_info['wtg_buses']
    update_wtg_buses = []
    for i in range(len(wtg_buses)):
        if wtg_buses[i][2] == 1:
            update_wtg_buses.append(wtg_buses[i])
    wtg_buses = update_wtg_buses
    POC_dummy_line = Project_info['ppc_line']
    network_impedance_line = Project_info['network_impedance_line']
    bus_inf = Project_info['SMIB']
    def find_Vsched(V_POC):
        psspy.fdns([1, 0, 0, 1, 1, 0, 0, 0])
        ierr, Vinf = psspy.busdat(bus_inf[0], 'PU')
        ierr, Vpcc = psspy.busdat(POC_dummy_line[0], 'PU')
        dV_PCC = V_POC - Vpcc
        Vsch = Vpcc
        count = 0
        while abs(dV_PCC) > MisMatch_Tol and count<=100:  # MisMatch_Tol = 0.001
            count +=1
            Vsch = Vinf + dV_PCC / 5
            psspy.plant_data(bus_inf[0],realar1=Vsch)  # set the infinite bus scheduled voltage to the estimated voltage for this condtion
            err = run_LoadFlow()
            ierr, Vinf = psspy.busdat(bus_inf[0], 'PU')
            ierr, Vpcc = psspy.busdat(POC_dummy_line[0], 'PU')
            dV_PCC = V_POC - Vpcc
        if count == 100:
            print('Warning: not able to find the Voltage Schedule at infinite bus')
        return Vsch
    sc_pcc_pu_Sbase = (scr*wfbase_MW)/100.0  #100MVA System base
    Rsys = math.sqrt(((1.0 / sc_pcc_pu_Sbase) ** 2) / (xr ** 2 + 1.0))
    Xsys = Rsys * xr
    psspy.psseinit(80000)
    psspy.case(case)
    num_cap = tune_capbank(P_POC, Q_POC, wfbase_MW, Project_info)
    psspy.branch_chng_3(network_impedance_line[0], network_impedance_line[1], network_impedance_line[2], realar1=Rsys,realar2=Xsys)
    Qs = []
    Ps = []
    MBASEs = []
    for i in range(len(wtg_buses)):
        ierr, MBASE = psspy.macdat(wtg_buses[i][0], wtg_buses[i][1], 'MBASE')
        MBASEs.append(MBASE)
    SMVA = sum(MBASEs)

    for i, wtg_bus in enumerate(wtg_buses):
        psspy.machine_chng_2(wtg_bus[0], wtg_bus[1], realar1=P_POC * MBASEs[i] / SMVA, realar2=0, realar3=0, realar4=0)
    run_LoadFlow()
    for i in range(len(wtg_buses)):
        ierr, Q = psspy.macdat(wtg_buses[i][0], wtg_buses[i][1], 'Q')
        Qs.append(Q)
        ierr, P = psspy.macdat(wtg_buses[i][0], wtg_buses[i][1], 'P')
        Ps.append(P)
    ierr, P_POC_act = psspy.brnmsc(POC_dummy_line[0], POC_dummy_line[1], POC_dummy_line[2], 'P')
    dP_POC = P_POC - P_POC_act
    count = 0
    while abs(dP_POC) > MisMatch_Tol * 50 and count <=100:
        count +=1
        for i, wtg_bus in enumerate(wtg_buses):
            P_add = ((dP_POC / 2) * (MBASEs[i] / SMVA))
            Ps[i] = Ps[i] + P_add
            psspy.machine_data_2(wtg_bus[0], wtg_bus[1], realar1=Ps[i])
        ierr, P_POC_act = psspy.brnmsc(POC_dummy_line[0], POC_dummy_line[1], POC_dummy_line[2], 'P')
        dP_POC = P_POC - P_POC_act
        run_LoadFlow()
    if count ==100:
        print('Warning: Warning: Not able to run active power loop ')
    ierr, Q_POC_act = psspy.brnmsc(POC_dummy_line[0], POC_dummy_line[1], POC_dummy_line[2], 'Q')
    dQ_POC = Q_POC - Q_POC_act
    count = 0
    while abs(dQ_POC) > MisMatch_Tol and count <=100:
        count += 1
        for i, wtg_bus in enumerate(wtg_buses):
            Q_add = ((dQ_POC / 2) * (MBASEs[i] / SMVA))
            Qs[i] = Qs[i] + Q_add
            psspy.machine_data_2(wtg_bus[0], wtg_bus[1], realar2=Qs[i], realar3=Qs[i], realar4=Qs[i])
        ierr, Q_POC_act = psspy.brnmsc(POC_dummy_line[0], POC_dummy_line[1], POC_dummy_line[2], 'Q')
        dQ_POC = Q_POC - Q_POC_act
        find_Vsched(V_POC)
        run_LoadFlow()
    if count==100:
        print('Warning: Warning: Not able to run reactive power loop ')
    for i, wtg_bus in enumerate(wtg_buses):
        psspy.machine_data_2(wtg_bus[0], wtg_bus[1], realar2=Qs[i], realar3=Qs[i], realar4=Qs[i])
    run_LoadFlow()
    ierr, P_POC_act = psspy.brnmsc(POC_dummy_line[0], POC_dummy_line[1], POC_dummy_line[2], 'P')
    dP_POC = P_POC - P_POC_act
    count = 0
    while abs(dP_POC) > MisMatch_Tol and count <=100:
        count +=1
        for i, wtg_bus in enumerate(wtg_buses):
            P_add = ((dP_POC / 2) * (MBASEs[i] / SMVA))
            Ps[i] = Ps[i] + P_add
            psspy.machine_data_2(wtg_bus[0], wtg_bus[1], realar1=Ps[i])
        ierr, P_POC_act = psspy.brnmsc(POC_dummy_line[0], POC_dummy_line[1], POC_dummy_line[2], 'P')
        dP_POC = P_POC - P_POC_act
        run_LoadFlow()
    if count == 100:
        print('Warning: Warning: Not able to run active power loop ')
    psspy.save(new_case)
    psspy.close_powerflow()
    psspy.pssehalt_2()
# ========================================================
def adjust_Q_according_Vdroop_ref(case, new_case, scr, xr, p_poc_MW, v_ref_pu,wfbase_MW,Project_info):
    import math
    network_impedance_line = Project_info['network_impedance_line']
    wtg_buses = Project_info['wtg_buses']
    V_deadband_pu = Project_info['VWRE_J'][1]
    droop_percent = Project_info['VWRE_J'][4]
    q_base_pu_wfbase = Project_info['VWRE_J'][11]
    update_wtg_buses = []
    for i in range(len(wtg_buses)):
        if wtg_buses[i][2] == 1:
            update_wtg_buses.append(wtg_buses[i])
    wtg_buses = update_wtg_buses
    poc_dummy_line = Project_info['ppc_line']
    for name in Project_info:
        if 'PPC_PSSE_VWRE' in name:
            q_base_pu_wfbase = Project_info['VWRE_J'][11]
            droop_percent = Project_info['VWRE_J'][4]
            V_deadband_pu = Project_info['VWRE_J'][1]
        elif 'PPC_V8' in name:
            q_base_pu_wfbase = Project_info['PPC_V8_J'][215]
            droop_percent = Project_info['PPC_V8_J'][211] * 100
            V_deadband_pu = Project_info['PPC_V8_J'][209]
    sc_pcc_pu_Sbase = (scr*wfbase_MW)/100.0  #100MVA System base
    Rsys = math.sqrt(((1.0 / sc_pcc_pu_Sbase) ** 2) / (xr ** 2 + 1.0))
    Xsys = Rsys * xr
    psspy.psseinit(80000)
    psspy.case(case)
    psspy.branch_chng_3(network_impedance_line[0], network_impedance_line[1], network_impedance_line[2], realar1=Rsys,realar2=Xsys)
    Qs = []
    Ps = []
    MBASEs = []
    for i in range(len(wtg_buses)):
        ierr, MBASE = psspy.macdat(wtg_buses[i][0], wtg_buses[i][1], 'MBASE')
        MBASEs.append(MBASE)
    SMVA = sum(MBASEs)
    for i, WTG_bus in enumerate(wtg_buses):
        psspy.machine_chng_2(WTG_bus[0], WTG_bus[1], realar1=p_poc_MW * MBASEs[i] / SMVA, realar2=0, realar3=0,realar4=0)
    run_LoadFlow()
    for i in range(len(wtg_buses)):
        ierr, Q = psspy.macdat(wtg_buses[i][0], wtg_buses[i][1], 'Q')
        Qs.append(Q)
        ierr, P = psspy.macdat(wtg_buses[i][0], wtg_buses[i][1], 'P')
        Ps.append(P)
    ierr, P_POC_act = psspy.brnmsc(poc_dummy_line[0], poc_dummy_line[1], poc_dummy_line[2], 'P')
    dP_POC = p_poc_MW - P_POC_act
    count = 0
    while abs(dP_POC) > MisMatch_Tol * 50 and count<=100:
        count += 1
        for i, WTG_bus in enumerate(wtg_buses):
            P_add = ((dP_POC / 2) * (MBASEs[i] / SMVA))
            Ps[i] = Ps[i] + P_add
            psspy.machine_data_2(WTG_bus[0], WTG_bus[1], realar1=Ps[i])
        ierr, P_POC_act = psspy.brnmsc(poc_dummy_line[0], poc_dummy_line[1], poc_dummy_line[2], 'P')
        dP_POC = p_poc_MW - P_POC_act
        run_LoadFlow()
    if count == 100:
        print('Warning: Not able to run active power loop')
    ierr, Q_POC_act = psspy.brnmsc(poc_dummy_line[0], poc_dummy_line[1], poc_dummy_line[2], 'Q')
    ierr, V_POC_Meas_PU = psspy.busdat(poc_dummy_line[0], 'PU')
    if v_ref_pu - V_POC_Meas_PU < -V_deadband_pu:
        Q_POC_ref = ((v_ref_pu - V_POC_Meas_PU + V_deadband_pu) / (droop_percent*0.01)) * q_base_pu_wfbase * wfbase_MW
    elif v_ref_pu - V_POC_Meas_PU > V_deadband_pu:
        Q_POC_ref = ((v_ref_pu - V_POC_Meas_PU - V_deadband_pu) / (droop_percent*0.01)) * q_base_pu_wfbase * wfbase_MW
    else:
        Q_POC_ref = 0
    dQ_POC = Q_POC_ref - Q_POC_act
    ierr, P_POC_act = psspy.brnmsc(poc_dummy_line[0], poc_dummy_line[1], poc_dummy_line[2], 'P')
    num_cap = tune_capbank(P_POC_act, Q_POC_ref, wfbase_MW, Project_info)
    count = 0
    while abs(dQ_POC) > MisMatch_Tol and count<=100:
        count += 1
        for i, WTG_bus in enumerate(wtg_buses):
            Q_add = ((dQ_POC / 2) * (MBASEs[i] / SMVA))
            Qs[i] = Qs[i] + Q_add
            psspy.machine_data_2(WTG_bus[0], WTG_bus[1], realar2=Qs[i], realar3=Qs[i], realar4=Qs[i])
        run_LoadFlow()
        ierr, Q_POC_act = psspy.brnmsc(poc_dummy_line[0], poc_dummy_line[1], poc_dummy_line[2], 'Q')
        ierr, V_POC_Meas_PU = psspy.busdat(poc_dummy_line[0], 'PU')
        if v_ref_pu - V_POC_Meas_PU < -V_deadband_pu:
            Q_POC_ref = ((v_ref_pu - V_POC_Meas_PU + V_deadband_pu) / (droop_percent * 0.01)) * q_base_pu_wfbase * wfbase_MW
        elif v_ref_pu - V_POC_Meas_PU > V_deadband_pu:
            Q_POC_ref = ((v_ref_pu - V_POC_Meas_PU - V_deadband_pu) / (droop_percent * 0.01)) * q_base_pu_wfbase * wfbase_MW
        else:
            Q_POC_ref = 0
        dQ_POC = Q_POC_ref - Q_POC_act
        ierr, P_POC_act = psspy.brnmsc(poc_dummy_line[0], poc_dummy_line[1], poc_dummy_line[2], 'P')
        num_cap = tune_capbank(P_POC_act, Q_POC_ref, wfbase_MW, Project_info)
        run_LoadFlow()
    if count == 100:
        print('Warning: Not able to run reactive power loop')
    for i, WTG_bus in enumerate(wtg_buses):
        psspy.machine_data_2(WTG_bus[0], WTG_bus[1], realar2=Qs[i], realar3=Qs[i], realar4=Qs[i])
    run_LoadFlow()
    ierr, P_POC_act = psspy.brnmsc(poc_dummy_line[0], poc_dummy_line[1], poc_dummy_line[2], 'P')
    dP_POC = p_poc_MW - P_POC_act
    count = 0
    while abs(dP_POC) > MisMatch_Tol and count<=100:
        count += 1
        for i, WTG_bus in enumerate(wtg_buses):
            P_add = ((dP_POC / 2) * (MBASEs[i] / SMVA))
            Ps[i] = Ps[i] + P_add
            psspy.machine_data_2(WTG_bus[0], WTG_bus[1], realar1=Ps[i])
        ierr, P_POC_act = psspy.brnmsc(poc_dummy_line[0], poc_dummy_line[1], poc_dummy_line[2], 'P')
        dP_POC = p_poc_MW - P_POC_act
        run_LoadFlow()
    if count == 100:
        print('Warning: Not able to run active power loop')
    run_LoadFlow()
    psspy.save(new_case)
    psspy.close_powerflow()
    psspy.pssehalt_2()
# ========================================================
def Convert():
    psspy.cong(0)
    psspy.conl(0,1,1,[0,0],[ 100.0,0.0,0.0, 100.0])
    psspy.conl(0,1,2,[0,0],[ 100.0,0.0,0.0, 100.0])
    psspy.conl(0,1,3,[0,0],[ 100.0,0.0,0.0, 100.0])
    psspy.ordr(1)
    psspy.fact()
    psspy.tysl(0)
# ========================================================
def dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs='',vestas_dlls='',add_dlls=''):
    wtg_buses = Project_info['wtg_buses']
    msu_lines = Project_info['msu_lines']
    ppc_line = Project_info['ppc_line']
    #! Dynamic Solution Parameters -----------------------------------
    max_solns = 200            # network solution maximum number of iterations
    sfactor = 0.3              # acceleration sfactor used in the network solution
    con_tol = 0.0001           # convergence tolerance used in the network solution
    dT = 0.001                 # simulation step time
    frq_filter = 0.04          # filter time constant used in calculating bus frequency deviations
    int_delta_thrsh = 0.06     # intermediate simulation mode time step threshold used in extended term simulations
    islnd_delta_thrsh = 0.14   # large (island frequency) simulation mode time step threshold used in extended term simulations
    islnd_sfactor = 1.0        # large (island frequency) simulation mode acceleration factor used in extended term simulations
    islnd_con_tol = 0.0005     # large (island frequency) simulation mode convergence tolerance used in extended term simulations
    #! Import dyr files
    print('new dyr file : %s' % os.path.basename(vestas_dyr))
    psspy.dyre_new([_i, _i, _i, _i], vestas_dyr, "", "", "")
    if add_dyrs != '':
        for add_dyr in add_dyrs:
            print('add dyr file : %s' %os.path.basename(add_dyr))
            psspy.dyre_add([_i, _i, _i, _i], add_dyr, "", "")
    #! Import dll files
    if vestas_dlls != '':
        for dll in vestas_dlls:
            psspy.addmodellibrary(dll)
    if add_dlls != '':
        for dll in add_dlls:
            psspy.addmodellibrary(dll)
    #! Dynamic parameter settings
    psspy.dynamics_solution_param_2(intgar1=max_solns, realar1=sfactor, realar2 =con_tol, realar3=dT, realar4 =frq_filter,
                        realar5=int_delta_thrsh, realar6=islnd_delta_thrsh, realar7=islnd_sfactor, realar8=islnd_con_tol)
    #! setup channels
    psspy.set_next_channel(1)
    i = 0
    for wtg_bus in wtg_buses:
        i += 1
        psspy.machine_array_channel([-1, 4, wtg_bus[0]], wtg_bus[1], "v_wtg%02d_meas_pu"%i)
        psspy.machine_array_channel([-1, 2, wtg_bus[0]], wtg_bus[1], "p_wtg%02d_meas_pu_sbase"%i)
        psspy.machine_array_channel([-1, 3, wtg_bus[0]], wtg_bus[1], "q_wtg%02d_meas_pu_sbase"%i)
    i = 0
    for msu_line in msu_lines:
        i += 1
        psspy.branch_p_and_q_channel([-1, -1, -1, msu_line[0], msu_line[1]], msu_line[2], ["p_msu%d_meas_MW" %i, "q_msu%d_meas_MVar"%i])
    # psspy.voltage_channel([-1, -1, -1, ppc_line[1]], "v_pcc_meas_pu")
    psspy.voltage_and_angle_channel([-1, -1, -1, ppc_line[1]], ["v_pcc_meas_pu", "angle_pcc_meas_deg"])
    psspy.branch_p_and_q_channel([-1, -1, -1, ppc_line[0], ppc_line[1]], ppc_line[2], ["p_pcc_meas_MW", "q_pcc_meas_MVar"])
    psspy.bus_frequency_channel([-1, ppc_line[1]], "f_dev_pcc_meas_pu")

    J_wtg_GSVWs, M_wtg_GSVWs, L_wtg_GSVWs, K_wtg_GSVWs = [],[],[],[]
    J_wtg_GSPQs, M_wtg_GSPQs, L_wtg_GSPQs, K_wtg_GSPQs = [],[],[],[]
    J_wtg_GSLHs, M_wtg_GSLHs, L_wtg_GSLHs, K_wtg_GSLHs = [],[],[],[]
    J_wtg_GSMEs, M_wtg_GSMEs, L_wtg_GSMEs, K_wtg_GSMEs = [],[],[],[]
    J_wtg_GSVFs, M_wtg_GSVFs, L_wtg_GSVFs, K_wtg_GSVFs = [],[],[],[]

    for wtg_bus in wtg_buses:
        ierr, J_wtg_GSVW = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WGEN', 'CON')
        J_wtg_GSVWs.append(J_wtg_GSVW)
        ierr, M_wtg_GSVW = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WGEN', 'ICON')
        M_wtg_GSVWs.append(M_wtg_GSVW)
        ierr, L_wtg_GSVW = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WGEN', 'VAR')
        L_wtg_GSVWs.append(L_wtg_GSVW)
        ierr, K_wtg_GSVW = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WGEN', 'STATE')
        K_wtg_GSVWs.append(K_wtg_GSVW)

        ierr, J_wtg_GSPQ = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WELEC', 'CON')
        J_wtg_GSPQs.append(J_wtg_GSPQ)
        ierr, M_wtg_GSPQ = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WELEC', 'ICON')
        M_wtg_GSPQs.append(M_wtg_GSPQ)
        ierr, L_wtg_GSPQ = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WELEC', 'VAR')
        L_wtg_GSPQs.append(L_wtg_GSPQ)
        ierr, K_wtg_GSPQ = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WELEC', 'STATE')
        K_wtg_GSPQs.append(K_wtg_GSPQ)

        ierr, J_wtg_GSLH = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WMECH', 'CON')
        J_wtg_GSLHs.append(J_wtg_GSLH)
        ierr, M_wtg_GSLH = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WMECH', 'ICON')
        M_wtg_GSLHs.append(M_wtg_GSLH)
        ierr, L_wtg_GSLH = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WMECH', 'VAR')
        L_wtg_GSLHs.append(L_wtg_GSLH)
        ierr, K_wtg_GSLH = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WMECH', 'STATE')
        K_wtg_GSLHs.append(K_wtg_GSLH)

        ierr, J_wtg_GSME = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WPICH', 'CON')
        J_wtg_GSMEs.append(J_wtg_GSME)
        ierr, M_wtg_GSME = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WPICH', 'ICON')
        M_wtg_GSMEs.append(M_wtg_GSME)
        ierr, L_wtg_GSME = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WPICH', 'VAR')
        L_wtg_GSMEs.append(L_wtg_GSME)
        ierr, K_wtg_GSME = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WPICH', 'STATE')
        K_wtg_GSMEs.append(K_wtg_GSME)

        ierr, J_wtg_GSVF = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WAERO', 'CON')
        J_wtg_GSVFs.append(J_wtg_GSVF)
        ierr, M_wtg_GSVF = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WAERO', 'ICON')
        M_wtg_GSVFs.append(M_wtg_GSVF)
        ierr, L_wtg_GSVF = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WAERO', 'VAR')
        L_wtg_GSVFs.append(L_wtg_GSVF)
        ierr, K_wtg_GSVF = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WAERO', 'STATE')
        K_wtg_GSVFs.append(K_wtg_GSVF)

    ierr, J_ppc_VWPO = psspy.windmind(wtg_buses[0][0], wtg_buses[0][1], 'WGUST', 'CON')
    ierr, M_ppc_VWPO = psspy.windmind(wtg_buses[0][0], wtg_buses[0][1], 'WGUST', 'ICON')
    ierr, L_ppc_VWPO = psspy.windmind(wtg_buses[0][0], wtg_buses[0][1], 'WGUST', 'VAR')
    ierr, K_ppc_VWPO = psspy.windmind(wtg_buses[0][0], wtg_buses[0][1], 'WGUST', 'STATE')

    ierr, J_ppc_VWRE = psspy.windmind(wtg_buses[0][0], wtg_buses[0][1], 'WAUX', 'CON')
    ierr, M_ppc_VWRE = psspy.windmind(wtg_buses[0][0], wtg_buses[0][1], 'WAUX', 'ICON')
    ierr, L_ppc_VWRE = psspy.windmind(wtg_buses[0][0], wtg_buses[0][1], 'WAUX', 'VAR')
    ierr, K_ppc_VWRE = psspy.windmind(wtg_buses[0][0], wtg_buses[0][1], 'WAUX', 'STATE')

    for i in range(len(Project_info['VWPO_L'])):
        if Project_info['VWPO_L'][i] == 1:
            psspy.var_channel([-1, L_ppc_VWPO + i], "VWPO_VAR_L%d"%i)
    for i in range(len(Project_info['VWRE_L'])):
        if Project_info['VWRE_L'][i] == 1:
            psspy.var_channel([-1, L_ppc_VWRE + i], "VWRE_VAR_L%d"%i)
    for i in range(len(Project_info['VWPO_K'])):
        if Project_info['VWPO_K'][i] == 1:
            psspy.state_channel([-1, K_ppc_VWPO + i], "VWPO_STATE_K%d"%i)
    for i in range(len(Project_info['VWRE_K'])):
        if Project_info['VWRE_K'][i] == 1:
            psspy.state_channel([-1, K_ppc_VWRE + i], "VWRE_STATE_K%d"%i)
    return J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE
# ========================================================
def dynamic_intialise_v8(Project_info, vestas_dyr, add_dyrs='', vestas_dlls='', add_dlls=''):
    wtg_buses = Project_info['wtg_buses']
    msu_lines = Project_info['msu_lines']
    ppc_line = Project_info['ppc_line']
    ppc_machine = Project_info['ppc_machine']
    #! Dynamic Solution Parameters -----------------------------------
    max_solns = 200            # network solution maximum number of iterations
    sfactor = 0.3              # acceleration sfactor used in the network solution
    con_tol = 0.0001           # convergence tolerance used in the network solution
    dT = 0.001                 # simulation step time
    frq_filter = 0.04         # filter time constant used in calculating bus frequency deviations
    int_delta_thrsh = 0.06     # intermediate simulation mode time step threshold used in extended term simulations
    islnd_delta_thrsh = 0.14   # large (island frequency) simulation mode time step threshold used in extended term simulations
    islnd_sfactor = 1.0        # large (island frequency) simulation mode acceleration factor used in extended term simulations
    islnd_con_tol = 0.0005     # large (island frequency) simulation mode convergence tolerance used in extended term simulations
    #! Import dyr files
    psspy.dyre_new([_i, _i, _i, _i], vestas_dyr, "", "", "")
    if add_dyrs != '':
        for add_dyr in add_dyrs:
            psspy.dyre_add([_i, _i, _i, _i], add_dyr, "", "")
    #! Import dll files
    if vestas_dlls != '':
        for dll in vestas_dlls:
            psspy.addmodellibrary(dll)
    if add_dlls != '':
        for dll in add_dlls:
            psspy.addmodellibrary(dll)
    #! Dynamic parameter settings
    psspy.dynamics_solution_param_2(intgar1=max_solns, realar1=sfactor, realar2 =con_tol, realar3=dT, realar4 =frq_filter,
                        realar5=int_delta_thrsh, realar6=islnd_delta_thrsh, realar7=islnd_sfactor, realar8=islnd_con_tol)
    #! setup channels
    psspy.set_next_channel(1)
    i=0
    for wtg_bus in wtg_buses:
        i+=1
        psspy.machine_array_channel([-1, 4, wtg_bus[0]], wtg_bus[1], "v_wtg%02d_meas_pu"%i)
        psspy.machine_array_channel([-1, 2, wtg_bus[0]], wtg_bus[1], "p_wtg%02d_meas_pu_sbase"%i)
        psspy.machine_array_channel([-1, 3, wtg_bus[0]], wtg_bus[1], "q_wtg%02d_meas_pu_sbase"%i)
    i = 0
    for msu_line in msu_lines:
        i += 1
        psspy.branch_p_and_q_channel([-1, -1, -1, msu_line[0], msu_line[1]], msu_line[2], ["p_msu%d_meas_MW" %i, "q_msu%d_meas_MVar"%i])
    # psspy.voltage_channel([-1, -1, -1, ppc_line[1]], "v_pcc_meas_pu")
    psspy.voltage_and_angle_channel([-1, -1, -1, ppc_line[1]], ["v_pcc_meas_pu", "angle_pcc_meas_deg"])
    psspy.branch_p_and_q_channel([-1, -1, -1, ppc_line[0], ppc_line[1]], ppc_line[2], ["p_pcc_meas_MW", "q_pcc_meas_MVar"])
    psspy.bus_frequency_channel([-1, ppc_line[1]], "f_dev_pcc_meas_pu")

    J_wtgs,M_wtgs,L_wtgs,K_wtgs = [],[],[],[]

    for wtg_bus in wtg_buses:
        ierr, J_wtg = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WGEN', 'CON')
        J_wtgs.append(J_wtg)
        ierr, M_wtg = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WGEN', 'ICON')
        M_wtgs.append(M_wtg)
        ierr, L_wtg = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WGEN', 'VAR')
        L_wtgs.append(L_wtg)
        ierr, K_wtg = psspy.windmind(wtg_bus[0], wtg_bus[1], 'WGEN', 'STATE')
        K_wtgs.append(K_wtg)

    start = Project_info['PPC_V8_title'].find('VESTAS')
    final = int(Project_info['PPC_V8_title'].find('504') - 2)
    model_name = Project_info['PPC_V8_title'][start:final]

    ierr, J_ppc = psspy.cctmind_buso(ppc_machine[0], model_name, 'CON')  # PPC VESTAS WIND TURBINE CONTROL
    ierr, M_ppc = psspy.cctmind_buso(ppc_machine[0], model_name, 'ICON')  # PPC VESTAS WIND TURBINE CONTROL
    ierr, L_ppc = psspy.cctmind_buso(ppc_machine[0], model_name, 'VAR')  # PPC VESTAS WIND TURBINE CONTROL
    ierr, K_ppc = psspy.cctmind_buso(ppc_machine[0], model_name, 'STATE')  # PPC VESTAS WIND TURBINE CONTROL

    for i in range(len(Project_info['PPC_V8_L'])):
        if Project_info['PPC_V8_L'][i] == 1:
            psspy.var_channel([-1, L_ppc + i], "PPC_VAR_L%d"%i)
    for i in range(len(Project_info['PPC_V8_K'])):
        if Project_info['PPC_V8_K'][i] == 1:
            psspy.state_channel([-1, K_ppc + i], "PPC_STATE_K%d"%i)
    return J_ppc,M_ppc,L_ppc,K_ppc
# =======================================================
def read_excel_tools(excel_file,sheet_index_or_name):
    import os, xlrd, shutil
    temp = 'C:\TEMP'
    check_path(temp)
    temp_excel_file = os.path.join(temp, os.path.basename(excel_file))
    try:
        xl_workbook = xlrd.open_workbook(excel_file)
        delTempFileFlag = False
    except:
        command = "xcopy \"" + excel_file + "\" " + "\"" + temp_excel_file + "*\""
        os.system(command)
        xl_workbook = xlrd.open_workbook(temp_excel_file)
        delTempFileFlag = True
    if delTempFileFlag == True:
        os.remove(temp_excel_file)
        delTempFileFlag = False
    if type(sheet_index_or_name) == int:
        xl_sheet = xl_workbook.sheet_by_index(sheet_index_or_name)
    elif type(sheet_index_or_name) == str:
        xl_sheet = xl_workbook.sheet_by_name(sheet_index_or_name)
    return xl_sheet
# =======================================================
def read_project_info(excel_file,version):
    import os, xlrd, shutil, re
    temp = 'C:\TEMP'
    check_path(temp)
    temp_excel_file = os.path.join(temp, os.path.basename(excel_file))
    try:
        xl_workbook = xlrd.open_workbook(excel_file)
        delTempFileFlag = False
    except:
        command = "xcopy \"" + excel_file + "\" " + "\"" + temp_excel_file + "*\""
        os.system(command)
        xl_workbook = xlrd.open_workbook(temp_excel_file)
        delTempFileFlag = True
    if delTempFileFlag == True:
        os.remove(temp_excel_file)
        delTempFileFlag = False
    sheet_names = xl_workbook.sheet_names()
    models = ['GSVW', 'GSPQ', 'GSLH', 'GSME', 'GSVF', 'VWPO', 'VWRE', 'WTG_GSV8', 'PPC_V8']
    Project_info = {}
    for sheet_name in sheet_names:
        for model in models:
            if model in sheet_name:
                xl_sheet = xl_workbook.sheet_by_name(sheet_name)
                for row_idx in range(0, xl_sheet.nrows):
                    cell_obj = (xl_sheet.cell(row_idx, 0)).value
                    if 'ICON' in cell_obj:
                        start_icon = row_idx
                    elif 'CON' in cell_obj:
                        start_con = row_idx
                    elif 'STATE' in cell_obj:
                        start_state = row_idx
                    elif 'VAR' in cell_obj:
                        start_var = row_idx
                row_title = xl_sheet.row(0)
                for col_idx in range(len(row_title)):
                    if version in row_title[col_idx].value:
                        # Title
                        Project_info[model + '_title'] = (xl_sheet.cell(start_icon, col_idx)).value
                        title = (xl_sheet.cell(start_icon, col_idx)).value
                        # read ICON
                        num_icon = int(re.split('[ :,;]', title)[-4])
                        num_con = int(re.split('[ :,;]', title)[-3])
                        num_state = int(re.split('[ :,;]', title)[-2])
                        num_var = int(re.split('[ :,;]', title)[-1])
                        col_idx_data = []
                        for row_idx in range(start_icon+1, start_icon+1+num_icon):
                            cell_obj = (xl_sheet.cell(row_idx, col_idx)).value
                            if isinstance(cell_obj, unicode):
                                cell_obj = str(cell_obj)
                            elif isinstance(cell_obj, float):
                                cell_obj = int(cell_obj)
                            if cell_obj !='':
                                col_idx_data.append(cell_obj)
                        Project_info[model + '_M'] = col_idx_data
                        # read CON
                        col_idx_data = []
                        for row_idx in range(start_con+1, start_con+1+num_con):
                            cell_obj = (xl_sheet.cell(row_idx, col_idx)).value
                            if isinstance(cell_obj, unicode):
                                cell_obj = str(cell_obj)
                            if cell_obj !='':
                                col_idx_data.append(cell_obj)
                        Project_info[model + '_J'] = col_idx_data
                        # read STATE
                        col_idx_data = []
                        for row_idx in range(start_state+1, start_state+1+num_state):
                            cell_obj = int((xl_sheet.cell(row_idx, col_idx)).value)
                            if cell_obj !='':
                                col_idx_data.append(cell_obj)
                        Project_info[model + '_K'] = col_idx_data
                        # read VAR
                        col_idx_data = []
                        for row_idx in range(start_var + 1, start_var+1+num_var):
                            cell_obj = int((xl_sheet.cell(row_idx, col_idx)).value)
                            if cell_obj != '':
                                col_idx_data.append(cell_obj)
                        Project_info[model + '_L'] = col_idx_data
        if sheet_name == 'WTG':
            xl_sheet = xl_workbook.sheet_by_name(sheet_name)
            WTG = []
            for col_idx in range(1, xl_sheet.ncols):
                WTG.append([int((xl_sheet.cell(1, col_idx)).value),str(int(xl_sheet.cell(2, col_idx).value)),int((xl_sheet.cell(3, col_idx)).value)])
            Project_info['wtg_buses'] = WTG
        if sheet_name == 'PPC':
            xl_sheet = xl_workbook.sheet_by_name(sheet_name)
            Project_info['ppc_machine'] = [int((xl_sheet.cell(1, 1)).value),str(int(xl_sheet.cell(2, 1).value))]
        if sheet_name == 'PPC':
            xl_sheet = xl_workbook.sheet_by_name(sheet_name)
            Project_info['ppc_line'] = [int((xl_sheet.cell(3, 1)).value),int(xl_sheet.cell(4, 1).value),str(int(xl_sheet.cell(5, 1).value))]
        if sheet_name == 'network_impedance_line':
            xl_sheet = xl_workbook.sheet_by_name(sheet_name)
            Project_info['network_impedance_line'] = [int((xl_sheet.cell(1, 1)).value),int(xl_sheet.cell(2, 1).value),str(int(xl_sheet.cell(3, 1).value))]
        if sheet_name == 'SMIB':
            xl_sheet = xl_workbook.sheet_by_name(sheet_name)
            Project_info['SMIB'] = [int((xl_sheet.cell(1, 1)).value),int(xl_sheet.cell(2, 1).value)]
        if sheet_name == 'MSU':
            xl_sheet = xl_workbook.sheet_by_name(sheet_name)
            Project_info['msu_lines'] = []
            for col_idx in range(1, xl_sheet.ncols):
                col_idx_data = (int((xl_sheet.cell(12, col_idx)).value),int(xl_sheet.cell(13, col_idx).value),str(int(xl_sheet.cell(14, col_idx).value)))
                Project_info['msu_lines'].append(col_idx_data)
        if sheet_name == 'MSU':
            xl_sheet = xl_workbook.sheet_by_name(sheet_name)
            Project_info['msu_caps'] = []
            for col_idx in range(1, xl_sheet.ncols):
                col_idx_data = (int(xl_sheet.cell(1, col_idx).value),(xl_sheet.cell(5, col_idx)).value,(xl_sheet.cell(6, col_idx)).value,(xl_sheet.cell(7, col_idx)).value,(xl_sheet.cell(8, col_idx)).value,(xl_sheet.cell(9, col_idx)).value,(xl_sheet.cell(10, col_idx)).value,(xl_sheet.cell(11, col_idx)).value)
                Project_info['msu_caps'].append(col_idx_data)

        xl_sheet = xl_workbook.sheet_by_name('UCMSUFB')
        dyrs = ['MSU']
        for row_idx in range(0, xl_sheet.nrows):
            cell_obj = (xl_sheet.cell(row_idx, 0)).value
            if 'ICON' in cell_obj:
                start_icon = row_idx
            elif 'CON' in cell_obj:
                start_con = row_idx
            elif 'STATE' in cell_obj:
                start_state = row_idx
            elif 'VAR' in cell_obj:
                start_var = row_idx
        row_title = xl_sheet.row(0)
        for col_idx in range(len(row_title)):
            for dyr in dyrs:
                if dyr in row_title[col_idx].value:
                    # Title
                    Project_info[dyr + '_description'] = (xl_sheet.cell(start_icon, col_idx)).value
                    title = (xl_sheet.cell(start_icon, col_idx)).value
                    # read ICON
                    num_icon = int(re.split('[ :,;]', title)[-4])
                    num_con = int(re.split('[ :,;]', title)[-3])
                    num_state = int(re.split('[ :,;]', title)[-2])
                    num_var = int(re.split('[ :,;]', title)[-1])
                    col_idx_data = []
                    for row_idx in range(start_icon + 1, start_icon + 1 + num_icon):
                        cell_obj = (xl_sheet.cell(row_idx, col_idx)).value
                        if isinstance(cell_obj, unicode):
                            cell_obj = str(cell_obj)
                        elif isinstance(cell_obj, float):
                            cell_obj = int(cell_obj)
                        if cell_obj != '':
                            col_idx_data.append(cell_obj)
                    Project_info[dyr + '_M'] = col_idx_data
                    # read CON
                    col_idx_data = []
                    for row_idx in range(start_con + 1, start_con + 1 + num_con):
                        cell_obj = (xl_sheet.cell(row_idx, col_idx)).value
                        if isinstance(cell_obj, unicode):
                            cell_obj = str(cell_obj)
                        if cell_obj != '':
                            col_idx_data.append(cell_obj)
                    Project_info[dyr + '_J'] = col_idx_data
                    # read STATE
                    col_idx_data = []
                    for row_idx in range(start_state + 1, start_state + 1 + num_state):
                        cell_obj = int((xl_sheet.cell(row_idx, col_idx)).value)
                        if cell_obj != '':
                            col_idx_data.append(cell_obj)
                    Project_info[dyr + '_K'] = col_idx_data
                    # read VAR
                    col_idx_data = []
                    for row_idx in range(start_var + 1, start_var + 1 + num_var):
                        cell_obj = int((xl_sheet.cell(row_idx, col_idx)).value)
                        if cell_obj != '':
                            col_idx_data.append(cell_obj)
                    Project_info[dyr + '_L'] = col_idx_data

        xl_sheet = xl_workbook.sheet_by_name('UCFI')
        dyrs = ['FCT_01', 'FCT_02', 'FCT_03', 'FCT_04', 'FCT_05', 'FCT_06', 'FCT_07', 'FCT_08']
        for row_idx in range(0, xl_sheet.nrows):
            cell_obj = (xl_sheet.cell(row_idx, 0)).value
            if 'ICON' in cell_obj:
                start_icon = row_idx
            elif 'CON' in cell_obj:
                start_con = row_idx
            elif 'STATE' in cell_obj:
                start_state = row_idx
            elif 'VAR' in cell_obj:
                start_var = row_idx
        row_title = xl_sheet.row(0)
        for col_idx in range(len(row_title)):
            for dyr in dyrs:
                if dyr in row_title[col_idx].value:
                    # Title
                    Project_info[dyr + '_description'] = (xl_sheet.cell(start_icon, col_idx)).value
                    title = (xl_sheet.cell(start_icon, col_idx)).value
                    # read ICON
                    num_icon = int(re.split('[ :,;]', title)[-4])
                    num_con = int(re.split('[ :,;]', title)[-3])
                    num_state = int(re.split('[ :,;]', title)[-2])
                    num_var = int(re.split('[ :,;]', title)[-1])
                    col_idx_data = []
                    for row_idx in range(start_icon + 1, start_icon + 1 + num_icon):
                        cell_obj = (xl_sheet.cell(row_idx, col_idx)).value
                        if isinstance(cell_obj, unicode):
                            cell_obj = str(cell_obj)
                        elif isinstance(cell_obj, float):
                            cell_obj = int(cell_obj)
                        if cell_obj != '':
                            col_idx_data.append(cell_obj)
                    Project_info[dyr + '_M'] = col_idx_data
                    # read CON
                    col_idx_data = []
                    for row_idx in range(start_con + 1, start_con + 1 + num_con):
                        cell_obj = (xl_sheet.cell(row_idx, col_idx)).value
                        if isinstance(cell_obj, unicode):
                            cell_obj = str(cell_obj)
                        if cell_obj != '':
                            col_idx_data.append(cell_obj)
                    Project_info[dyr + '_J'] = col_idx_data
                    # read STATE
                    col_idx_data = []
                    for row_idx in range(start_state + 1, start_state + 1 + num_state):
                        cell_obj = int((xl_sheet.cell(row_idx, col_idx)).value)
                        if cell_obj != '':
                            col_idx_data.append(cell_obj)
                    Project_info[dyr + '_K'] = col_idx_data
                    # read VAR
                    col_idx_data = []
                    for row_idx in range(start_var + 1, start_var + 1 + num_var):
                        cell_obj = int((xl_sheet.cell(row_idx, col_idx)).value)
                        if cell_obj != '':
                            col_idx_data.append(cell_obj)
                    Project_info[dyr + '_L'] = col_idx_data
    return Project_info
# =======================================================
def create_project_dyr(Project_info,dyr_file):
    dyrfile = open(dyr_file, "w")
    dyrfile.write("/ Dyre file for %s\n" %(os.path.basename(dyr_file)[0:-4]))
    print("Created the new dyre file: %s" %dyr_file)
    for machine in Project_info['wtg_buses']:
        if machine[2] == 1:
            for name in Project_info:
                if '_title' in name and 'VWPO' not in name and 'VWRE' not in name and 'PPC_V8' not in name:
                    dyrfile.write("%d 'USRMDL'   %s   %s" %(machine[0],machine[1],Project_info[name]))
                    if len(Project_info[name.replace("_title", "_M")]) > 7:
                        dyrfile.write("\n")
                        dyrfile.write("{: <5}".format(''))
                        for i in range(len(Project_info[name.replace("_title", "_M")])):
                            dyrfile.write("{: <15}".format(Project_info[name.replace("_title", "_M")][i]))
                            if i == len(Project_info[name.replace("_title", "_M")]) - 1:
                                dyrfile.write("/")
                            elif (i + 1) % 7 == 0:
                                dyrfile.write("\n")
                                dyrfile.write("{: <5}".format(''))
                    else:
                        dyrfile.write(",   ")
                        for i in range(len(Project_info[name.replace("_title", "_M")])):
                            if type(Project_info[name.replace("_title", "_M")][i]) == int:
                                dyrfile.write("%1.0f   " % Project_info[name.replace("_title", "_M")][i])
                            elif type(Project_info[name.replace("_title", "_M")][i]) == str:
                                dyrfile.write("%s   " % Project_info[name.replace("_title", "_M")][i])
                    dyrfile.write("\n")
                    dyrfile.write("{: <5}".format(''))
                    for i in range(len(Project_info[name.replace("_title", "_J")])):
                        dyrfile.write("{: <15}".format(Project_info[name.replace("_title", "_J")][i]))
                        if i == len(Project_info[name.replace("_title", "_J")]) - 1:
                            dyrfile.write("/")
                        elif (i + 1) % 7 == 0:
                            dyrfile.write("\n")
                            dyrfile.write("{: <5}".format(''))
                    dyrfile.write("\n\n")
    for name in Project_info:
        if 'VWPO_title' in name or 'VWRE_title' in name:
            dyrfile.write("%d 'USRMDL'   %s   %s" % (Project_info['ppc_machine'][0], Project_info['ppc_machine'][1], Project_info[name]))
            if len(Project_info[name.replace("_title", "_M")]) > 7:
                dyrfile.write("\n")
                dyrfile.write("{: <5}".format(''))
                for i in range(len(Project_info[name.replace("_title", "_M")])):
                    dyrfile.write("{: <15}".format(Project_info[name.replace("_title", "_M")][i]))
                    if (i + 1) % 7 == 0:
                        dyrfile.write("\n")
                        dyrfile.write("{: <5}".format(''))
            else:
                dyrfile.write(",   ")
                for i in range(len(Project_info[name.replace("_title", "_M")])):
                    if type(Project_info[name.replace("_title", "_M")][i]) == int:
                        dyrfile.write("%1.0f   " % Project_info[name.replace("_title", "_M")][i])
                    elif type(Project_info[name.replace("_title", "_M")][i]) == str:
                        dyrfile.write("%s   " % Project_info[name.replace("_title", "_M")][i])
            dyrfile.write("\n")
            dyrfile.write("{: <5}".format(''))
            for i in range(len(Project_info[name.replace("_title", "_J")])):
                dyrfile.write("{: <15}".format(Project_info[name.replace("_title", "_J")][i]))
                if i == len(Project_info[name.replace("_title", "_J")]) - 1:
                    dyrfile.write("/")
                elif (i + 1) % 7 == 0:
                    dyrfile.write("\n")
                    dyrfile.write("{: <5}".format(''))
            dyrfile.write("\n\n")
        elif 'PPC_V8_title' in name:
            dyrfile.write("%d 'USRBUS'   %s" % (Project_info['ppc_machine'][0], Project_info[name]))
            if len(Project_info[name.replace("_title", "_M")]) > 7:
                dyrfile.write("\n")
                dyrfile.write("{: <5}".format(''))
                for i in range(len(Project_info[name.replace("_title", "_M")])):
                    dyrfile.write("{: <15}".format(Project_info[name.replace("_title", "_M")][i]))
                    if (i + 1) % 7 == 0:
                        dyrfile.write("\n")
                        dyrfile.write("{: <5}".format(''))
            else:
                dyrfile.write(",   ")
                for i in range(len(Project_info[name.replace("_title", "_M")])):
                    if type(Project_info[name.replace("_title", "_M")][i]) == int:
                        dyrfile.write("%1.0f   " % Project_info[name.replace("_title", "_M")][i])
                    elif type(Project_info[name.replace("_title", "_M")][i]) == str:
                        dyrfile.write("%s   " % Project_info[name.replace("_title", "_M")][i])
            dyrfile.write("\n")
            dyrfile.write("{: <5}".format(''))
            for i in range(len(Project_info[name.replace("_title", "_J")])):
                dyrfile.write("{: <15}".format(Project_info[name.replace("_title", "_J")][i]))
                if i == len(Project_info[name.replace("_title", "_J")]) - 1:
                    dyrfile.write("/")
                elif (i + 1) % 7 == 0:
                    dyrfile.write("\n")
                    dyrfile.write("{: <5}".format(''))
            dyrfile.write("\n\n")
    dyrfile.write("/ End of %s models.\n" %(os.path.basename(dyr_file)[0:-4]))
    dyrfile.close()
# =======================================================
def create_user_dyrs(Project_info,dyr_folder):
    check_path(dyr_folder)
    for name in Project_info:
        if '_description' in name:
            dyr_file = os.path.join(dyr_folder,'%s.dyr' %name[0:-12])
            dyrfile = open(dyr_file, "w")
            dyrfile.write("/ Dyre file for %s\n" % (os.path.basename(dyr_file)[0:-4]))
            print("Created the new dyre file: %s" % dyr_file)
            dyrfile.write("%d 'USRBUS'   %s" % (Project_info['ppc_line'][1], Project_info[name]))
            if len(Project_info[name.replace("_description", "_M")]) > 7:
                dyrfile.write("\n")
                dyrfile.write("{: <5}".format(''))
                for i in range(len(Project_info[name.replace("_description", "_M")])):
                    dyrfile.write("{: <15}".format(Project_info[name.replace("_description", "_M")][i]))
                    if (i + 1) % 7 == 0:
                        dyrfile.write("\n")
                        dyrfile.write("{: <5}".format(''))
            else:
                dyrfile.write(",   ")
                for i in range(len(Project_info[name.replace("_description", "_M")])):
                    if type(Project_info[name.replace("_description", "_M")][i]) == int:
                        dyrfile.write("%1.0f   " % Project_info[name.replace("_description", "_M")][i])
                    elif type(Project_info[name.replace("_description", "_M")][i]) == str:
                        dyrfile.write("%s   " % Project_info[name.replace("_description", "_M")][i])
            dyrfile.write("\n")
            dyrfile.write("{: <5}".format(''))
            for i in range(len(Project_info[name.replace("_description", "_J")])):
                dyrfile.write("{: <15}".format(Project_info[name.replace("_description", "_J")][i]))
                if i == len(Project_info[name.replace("_description", "_J")]) - 1:
                    dyrfile.write("/")
                elif (i + 1) % 7 == 0:
                    dyrfile.write("\n")
                    dyrfile.write("{: <5}".format(''))
            dyrfile.write("\n\n")
            dyrfile.write("/ End of %s models.\n" %(os.path.basename(dyr_file)[0:-4]))
            dyrfile.close()
# =======================================================
def create_GENCLS_dyr(Project_info,GENCLS_file):
    dyrfile = open(GENCLS_file, "w")
    dyrfile.write("/Existing generator model for the slack bus\n")
    dyrfile.write("%d 'GENCLS' %d     0.0000       0.0000    /" % (Project_info['SMIB'][0], Project_info['SMIB'][1]))
    dyrfile.close()

# =======================================================
def Droop_parameters(Project_info,wfbase_MW):
    if 'VWRE_J' in Project_info:
        POC_V_ref_Channel_id = 'VWRE_VAR_L4'
        V_POC_Channel_id = 'V_PCC_MEAS_PU'
        Q_POC_Channel_id = 'Q_PCC_MEAS_MVAR'
        QDROOP_over_percent = Project_info['VWRE_J'][4]
        QDROOP_lower_percent = Project_info['VWRE_J'][4]
        QBASE_pu = Project_info['VWRE_J'][11]
        deadband_lower = Project_info['VWRE_J'][1]
        deadband_over = Project_info['VWRE_J'][1]
        Qrefupperlimit = Project_info['VWRE_J'][19]
        Qreflowerlimit = Project_info['VWRE_J'][20]
    else:
        POC_V_ref_Channel_id = 'PPC_VAR_L2'
        V_POC_Channel_id = 'V_PCC_MEAS_PU'
        Q_POC_Channel_id = 'Q_PCC_MEAS_MVAR'
        QDROOP_over_percent = Project_info['PPC_V8_J'][211] * 100
        QDROOP_lower_percent = Project_info['PPC_V8_J'][214] * 100
        QBASE_pu = Project_info['PPC_V8_J'][215]
        deadband_lower = Project_info['PPC_V8_J'][209]
        deadband_over = Project_info['PPC_V8_J'][212]
        Qrefupperlimit = Project_info['PPC_V8_J'][216]
        Qreflowerlimit = -Project_info['PPC_V8_J'][217]
    plot_Vdroop = [POC_V_ref_Channel_id,Q_POC_Channel_id, QDROOP_over_percent,QDROOP_lower_percent, QBASE_pu, wfbase_MW, deadband_lower,deadband_over]
    plot_Qdroop = [POC_V_ref_Channel_id,V_POC_Channel_id, QDROOP_over_percent, QDROOP_lower_percent, QBASE_pu, wfbase_MW, deadband_lower, deadband_over, Qrefupperlimit, Qreflowerlimit]
    return plot_Vdroop,plot_Qdroop
# =======================================================
def convert_out_xlsx(outfiles):
    import sys, os
    import psse34
    import psspy
    import dyntools
    signals = []  # will print all signals, else you can indicate them like so   signals = [1,2,4,5]
    for outfile in outfiles:
        if outfile.endswith(".out"):
            excel_file = outfile.split(".out")[0] + "_excel"
            dyntools.CHNF(outfile).xlsout(channels=signals, show=False, outfile=outfile, xlsfile=excel_file)
# =======================================================
def get_data(xl_sheet):
    import re
    chanid = {}
    chandata = {}
    title_row = 1
    title_col = 0
    chan_number_row = 0
    title_row = 0
    short_title = os.path.splitext(filename)[0]
    row_chan = xl_sheet.row(chan_number_row)
    row_title = xl_sheet.row(title_row)
    for chan_idx in range(len(row_chan)):
        chan_idx_data = []
        for row_idx in range(1, xl_sheet.nrows):
            cell_obj = (xl_sheet.cell(row_idx, chan_idx)).value
            if (isinstance(cell_obj, unicode)):
                if len(re.findall(("\d+\.\d+"), cell_obj)) >0:
                    cell_obj = float(re.findall("\d+\.\d+", cell_obj)[0])
                elif len(re.findall(("\d+"), cell_obj)) > 0:
                    cell_obj = float(re.findall("\d+", cell_obj)[0])
                else:
                    print 'This is not float type'
            #print(cell_obj)
            # if type(cell_obj) != float:
            #     cell_obj = float(re.findall("\d+\.\d+", cell_obj)[0])
            chan_idx_data.append(cell_obj)
        if chan_idx ==0:
            chan_idx_data = []
            simtime = 0
            T_step = 0.02
            for row_idx in range(1, xl_sheet.nrows):
                chan_idx_data.append(simtime)
                simtime += T_step
            chanid['time'] = str(row_title[chan_idx].value)
            chandata['time'] = chan_idx_data
        else:
            chanid[int(chan_idx)] = str(row_title[chan_idx].value)
            chandata[int(chan_idx)] = chan_idx_data
    return short_title, chanid, chandata

# =======================================================
def log_file(dat_file_name):
    redirect.psse2py()
    psspy.psseinit()     #initialise PSSE so psspy commands can be called
    psspy.throwPsseExceptions = True
    try:
        os.mkdir(os.path.dirname(dat_file_name))
    except OSError:
        pass
    Progress = dat_file_name
    Outf=open(Progress,'w+')
    Outf.close()
    psspy.t_progress_output(2,Progress,[2,0]) #Use this API to specify the progress output device.

# =======================================================
def report_file(dat_file_name):
    redirect.psse2py()
    psspy.psseinit()     #initialise PSSE so psspy commands can be called
    psspy.throwPsseExceptions = True
    try:
        os.mkdir(os.path.dirname(dat_file_name))
    except OSError:
        pass
    Report = dat_file_name
    Outf=open(Report,'w+')
    Outf.close()
    psspy.t_report_output(2,Report,[2,0]) #Use this API to specify the progress output device.
# ========================================================
def get_files_name(path, fileTypes):
    cases = []
    for type in fileTypes:
        for file in os.listdir(path):
            if (not file.endswith(type)):
                continue
            else:
                cases.append(os.path.join(path,file))
    return(cases)
# ========================================================
def copy(src_folder,destination_fld,fileTypes):
    for file in os.listdir(src_folder):
        for type in fileTypes:
            if file.endswith(type):
                shutil.copy(os.path.join(src_folder,file),destination_fld)
# ========================================================
def copy_list(files,destination_fld):
    for file in files:
        shutil.copy(file,destination_fld)
# ========================================================
def delete_first_line(python_file):
    import os
    with open(python_file, "r") as f:
        lines = f.readlines()
    with open(python_file, "w") as f:
        for i in range(len(lines)):
            if i != 0:
                f.write(lines[i])
            else:
                f.write('# This is the python file for %s \n' %os.path.basename(python_file)[0:-3])
# ========================================================
def read_excel_sheet(excel_file, sheet_name):
    import os,xlrd,shutil
    temp = 'C:\TEMP'
    check_path(temp)
    temp_excel_file = os.path.join(temp, os.path.basename(excel_file))
    try:
        xl_workbook = xlrd.open_workbook(excel_file)
        delTempFileFlag = False
    except:
        command = "xcopy \"" + excel_file + "\" " + "\"" + temp_excel_file + "*\""
        os.system(command)
        xl_workbook = xlrd.open_workbook(temp_excel_file)
        delTempFileFlag = True
    if delTempFileFlag == True:
        os.remove(temp_excel_file)
        delTempFileFlag = False
    xl_sheet = xl_workbook.sheet_by_name(sheet_name)
    row_title = xl_sheet.row(0)
    chanid = {}
    data = {}
    for col_idx in range(len(row_title)):
        col_idx_data = []
        for row_idx in range(1, xl_sheet.nrows):
            cell_obj = (xl_sheet.cell(row_idx, col_idx)).value
            col_idx_data.append(cell_obj)
        chanid[col_idx] = row_title[col_idx].value
        data[row_title[col_idx].value] = col_idx_data
    return data
# ========================================================
def check_path(path):
    try:
        os.mkdir(path)
    except OSError:
        pass
# ========================================================
def dirCreateClean(path,fileTypes):
    import glob,os
    cwd = os.getcwd()
    def subCreateClean(subpath,fileTypes):
        try:
            os.mkdir(subpath)
        except OSError:
            pass
        #delete all files listed in fileTypes in the directory
        os.chdir(subpath)
        for type in fileTypes:
            filelist = glob.glob(type)
            for f in filelist:
                os.remove(f)
    subCreateClean(path,fileTypes)
    os.chdir(cwd)
# ========================================================
def merge_pdf(dir_to_merge):
    import os
    copy_dir = os.path.dirname(dir_to_merge)
    os.chdir(dir_to_merge)
    x = [a for a in os.listdir(dir_to_merge) if a.endswith(".pdf")]
    if len(x)>0:
        merger = PdfFileMerger()
        for pdf in x:
            merger.append(open(pdf, 'rb'))
        with open("Plots.pdf", "wb") as fout:
            merger.write(fout)
        for file in os.listdir(dir_to_merge):
            if file.endswith(".pdf") and (not file.endswith("Plots.pdf")):
                os.remove(os.path.join(dir_to_merge, file))
        for file in os.listdir(dir_to_merge):
            if (not file.endswith(".pdf")):
                continue
            dst_file = dir_to_merge+'\\'+file
            new_dst_file = dir_to_merge+'\\'+ dir_to_merge.split(os.sep)[-1] +'_' + file
            os.rename(dst_file, new_dst_file)
            shutil.copy(new_dst_file,copy_dir)
# ========================================================
def merge_pdf2(dir_to_merge):
    import os
    os.chdir(dir_to_merge)
    x = [a for a in os.listdir(dir_to_merge) if a.endswith(".pdf")]
    if len(x)>0:
        merger = PdfFileMerger()
        for pdf in x:
            merger.append(open(pdf, 'rb'))
        with open("Plots.pdf", "wb") as fout:
            merger.write(fout)
        for file in os.listdir(dir_to_merge):
            if file.endswith(".pdf") and os.path.basename(file) != "Plots.pdf":
                os.remove(os.path.join(dir_to_merge, file))
        for file in os.listdir(dir_to_merge):
            if (not file.endswith(".pdf")):
                continue
            dst_file = dir_to_merge+'\\'+file
            new_dst_file = dir_to_merge+'\\'+ dir_to_merge.split(os.sep)[-1] +'_' + file
            os.rename(dst_file, new_dst_file)
# ========================================================
def plot_all_in_dir(dir_to_plot,Test,Project_info,wfbase_MW):
    dirCreateClean(dir_to_plot, ["*.pdf"])
    filenames = get_files_name(dir_to_plot, ['.out'])
    plot_Vdroop, plot_Qdroop = Droop_parameters(Project_info, wfbase_MW)
    for filename in filenames:
        plot = pl.Plotting()
        plot.read_data(filename)
        if 'VWRE_J' in Project_info:
            if 'APT' in Test:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                plot.subplot_spec(3, (0, 'VWPO_VAR_L2'), scale=wfbase_MW)
                plot.legends[3] = ['P_PCC_MEAS_MW', 'P_PCC_REF_MW_L2']
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
            elif 'VCT' in Test or 'VDT' in Test:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0, plot_Vdroop=plot_Vdroop)
                plot.subplot_spec(1, (0, 'VWRE_VAR_L4'))
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0,plot_Qdroop=plot_Qdroop)
            elif 'RPT' in Test:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
                plot.subplot_spec(5, (0, 'VWRE_VAR_L0'), scale=wfbase_MW)
            elif 'FCT' in Test:
                plot.subplot_spec(0, (0, 'F_DEV_PCC_MEAS_PU'), title='Frequency', ylabel='Frequency (Hz)', scale=50.0,offset=50.0)
                plot.subplot_spec(0, (0, 'VWPO_STATE_K2'), scale=50.0, offset=50.0)
                plot.legends[0] = ['PCC Frequency', 'Filtered_frequency_K2']
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
            elif 'MSU_logic' in Test:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                plot.subplot_spec(3, (0, 'VWPO_VAR_L2'), scale=wfbase_MW)
                plot.legends[3] = ['P_PCC_MEAS_MW', 'P_PCC_REF_MW_L2']
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
                plot.subplot_spec(5, (0, 'VWRE_VAR_L0'), scale=wfbase_MW)
                for i, msu_line in enumerate(Project_info['msu_lines']):
                    plot.subplot_spec(5, (0, 'Q_MSU%d_MEAS_MVAR' % (i + 1)), scale=1.0, offset=0.0)
            elif 'ADT' in Test:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                plot.subplot_spec(2, (0, 'ANGLE_PCC_MEAS_DEG'), title='POC Voltage Angle', ylabel='Angle (deg)',scale=1.0, offset=0.0)
                # for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                #     plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
            else:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                plot.subplot_spec(1, (0, 'VWRE_VAR_L4'))
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
        else:
            if 'APT' in Test:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                plot.subplot_spec(3, (0, 'PPC_VAR_L0'), scale=wfbase_MW)
                plot.legends[3] = ['P_PCC_MEAS_MW', 'P_PCC_REF_MW_L2']
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
            elif 'VCT' in Test or 'VDT' in Test:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0, plot_Vdroop=plot_Vdroop)
                plot.subplot_spec(1, (0, 'PPC_VAR_L2'))
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0,plot_Qdroop=plot_Qdroop)
            elif 'RPT' in Test:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
                plot.subplot_spec(5, (0, 'PPC_VAR_L1'), scale=wfbase_MW)
            elif 'FCT' in Test:
                plot.subplot_spec(0, (0, 'F_DEV_PCC_MEAS_PU'), title='Frequency', ylabel='Frequency (Hz)', scale=50.0,offset=50.0)
                plot.legends[0] = ['PCC Frequency', 'Filtered_frequency_K2']
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
            elif 'MSU_logic' in Test:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                plot.subplot_spec(3, (0, 'PPC_VAR_L0'), scale=wfbase_MW)
                plot.legends[3] = ['P_PCC_MEAS_MW', 'P_PCC_REF_MW_L2']
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
                plot.subplot_spec(5, (0, 'PPC_VAR_L1'), scale=wfbase_MW)
                for i, msu_line in enumerate(Project_info['msu_lines']):
                    plot.subplot_spec(5, (0, 'Q_MSU%d_MEAS_MVAR' % i), scale=1.0, offset=0.0)
            elif 'ADT' in Test:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                plot.subplot_spec(2, (0, 'ANGLE_PCC_MEAS_DEG'), title='POC Voltage Angle', ylabel='Angle (deg)',scale=1.0, offset=0.0)
                # for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                #     plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
            else:
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(0, (0, 'V_WTG%02d_MEAS_PU' % (i + 1)), title='WTG Voltage', ylabel='Voltage (pu)',scale=1.0, offset=0.0)
                plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='POC Voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(2, (0, 'P_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Active Power at WTG',ylabel='Active Power(MW)', scale=100.0, offset=0.0)
                plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)',scale=1.0, offset=0.0)
                for i, wtg_bus in enumerate(Project_info['wtg_buses']):
                    plot.subplot_spec(4, (0, 'Q_WTG%02d_MEAS_PU_SBASE' % (i + 1)), title='Reactive Power at WTG',ylabel='Reactive Power(MVar)', scale=100.0, offset=0.0)
                plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC',ylabel='Reactive power(MVar)', scale=1.0, offset=0.0)
        plot.plot(figname=os.path.splitext(filename)[0], show=0)
    merge_pdf(dir_to_plot)
# =======================================================
def Three_Ph_Lflt(Project_info,study_case,main_python, org_sav_file,vestas_dyr,add_dyrs,vestas_dlls,add_dlls, NEM_Conv_file, results,Plotting_option,delete_option,overlay, MSU_discharging_time,V_auto_mode):
    Test = 'Initial'
    output_files_dir = os.path.join(results, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf","*.sav","*.snp","*.dat","*.dll","*.txt","*.py"])
    dirCreateClean(output_files_dir, ["*.pdf"])
    test_cases = read_excel_sheet(study_case,os.path.basename(main_python)[0:-3])
    wfbase_MW = check_wfbase_MW(Project_info,org_sav_file)
    if overlay != '':
        check_path(overlay)
        overlay_dir = os.path.join(overlay, '%s' % Test)
        check_path(overlay_dir)
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            new_sav_file = '%s.sav' %os.path.join(output_files_dir, 'Case_%02d_%s_PSSE' % (i, Test))
            scr = test_cases['SCR'][i]
            xr = test_cases['XR'][i]
            p_poc_MW = test_cases['Initial P (MW)'][i]
            Q_POC = test_cases['Initial Q (MVAr)'][i]
            V_POC = test_cases['V_actual_POC (pu)'][i]
            v_ref_pu = test_cases['Vref_droop (pu)'][i]
            if 'vdroopref' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_according_Vdroop_ref(org_sav_file, new_sav_file, scr, xr, p_poc_MW, v_ref_pu, wfbase_MW, Project_info)
            elif 'pqv' in str(test_cases['Name 2'][i]).lower():
                adjust_Q_and_V_POC_SMIB(org_sav_file, new_sav_file, scr, xr, p_poc_MW, Q_POC, V_POC, wfbase_MW, Project_info)
            else:
                shutil.copy(org_sav_file, output_files_dir)
                new_sav_file = os.path.join(output_files_dir,os.path.basename(org_sav_file))

            from_bus = int(test_cases['Ref 1'][i])
            to_bus = int(test_cases['Ref 2'][i])
            ckt = str(test_cases['Ref 3'][i])
            near_fault_time = float(test_cases['Ref 4'][i])
            remote_fault_time = float(test_cases['Ref 5'][i])

            psspy.psseinit(80000)
            log_file('%s_1.dat' % new_sav_file[0:-4])
            psspy.case(new_sav_file)
            if 'nem' in NEM_Conv_file.lower():
                sys.path.append(os.path.dirname(NEM_Conv_file))
                import NEM_Conv
                NEM_Conv.NEM_convert()
            else:
                Convert()
            if 'VWRE_J' in Project_info:
                J_ppc_VWPO,M_ppc_VWPO,L_ppc_VWPO,K_ppc_VWPO,J_ppc_VWRE,M_ppc_VWRE,L_ppc_VWRE,K_ppc_VWRE = dynamic_intialise_legacy(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
                psspy.change_con(J_ppc_VWRE + 12, V_auto_mode)  ### Vauto control
                psspy.change_con(J_ppc_VWRE + 57, MSU_discharging_time)  ### MSU discharging time
                msu_caps = Project_info['msu_caps']
                for msu_cap in msu_caps:
                    psspy.switched_shunt_chng_3(msu_cap[0],intgar9 = 1)
            else:
                J_ppc,M_ppc,L_ppc,K_ppc = dynamic_intialise_v8(Project_info,vestas_dyr,add_dyrs,vestas_dlls,add_dlls)
            start(new_sav_file,test_cases,i, vestas_dlls, add_dlls)
            runtime = 10
            psspy.run(0, runtime, nprt, nplt, crtplt)
            psspy.ltap(from_bus, to_bus, ckt, 0.1, 888, 'FLT_BUS', _f)
            psspy.dist_bus_fault(888, 1, _f, [0.0, -0.2E+10])
            ierr, runtime = psspy.dsrval('TIME', 0)
            psspy.run(0, runtime + near_fault_time, nprt, nplt, crtplt)
            psspy.branch_data(888, from_bus, ckt, intgar1 = 0)
            psspy.run(0, runtime + remote_fault_time, nprt, nplt, crtplt)
            psspy.branch_data(888, to_bus, ckt, intgar1 = 0)
            psspy.dscn(888)
            psspy.dist_clear_fault(1)
            if arc_time > 0.01:  # ARC_time: AutoReClosing time
                ierr, runtime = psspy.dsrval('TIME', 0)
                psspy.run(0, runtime + arc_time, nprt, nplt, crtplt)
                psspy.recn(888)
                psspy.dist_bus_fault(888, 1, _f, [0.0, -0.2E+10])
                psspy.branch_data(888, from_bus, ckt, intgar1=1)
                psspy.run(0, runtime + arc_time + near_fault_time, nprt, nplt, crtplt)
                psspy.dscn(888)
                psspy.branch_data(888, from_bus, ckt, intgar1=0)
            ierr, runtime = psspy.dsrval('TIME', 0)
            psspy.run(0, runtime + 5.0, nprt, nplt, crtplt)
            os.chdir(os.path.dirname(main_python))
            psspy.stoprecording()
            delete_first_line('%s.py' % new_sav_file[0:-4])
            psspy.close_powerflow()
            psspy.pssehalt_2()
    if 'y' in delete_option.lower():
        dirCreateClean(output_files_dir, ["*.dll", "*.txt"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir, Test, Project_info, wfbase_MW)
    if overlay != '':
        copy(output_files_dir, overlay_dir, [".out"])

# =======================================================
def threeph_fault_line_CB_Fail(output_files_dir, output_filename_suffix, case_file, Model_dyr, add_dyrs, from_bus, to_bus, ckt, remote_fault_time, CB_fail_time, arc_time):
    psspy.psseinit(80000)
    psspy.case(os.path.join(NEM_input_files_dir, case_file))
    network_modifications(case_file)
    ierr, V_base = psspy.busdat(from_bus, 'BASE')
    ierr, Fr_bus = psspy.notona(from_bus)
    ierr, To_bus = psspy.notona(to_bus)
    arc = ''
    if arc_time > 0.01:
        arc = '_ARC'
    output_filename = '%sThreePhFlt_CBF%s_%s_%1.0fkV_Fr_%s_To_%s_%s' % (output_filename_suffix,arc, case_file[0:-4], V_base, Fr_bus[:12].strip(), To_bus[:12].strip(),ckt)
    dynamic_intialise(output_files_dir, output_filename, Model_dyr, add_dyrs)
    psspy.ltap(from_bus, to_bus, ckt, 0.2, 888, 'FLT_BUS', _f)
    psspy.dist_bus_fault(888, 3, _f, [0, -0.2E-10])
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + remote_fault_time, NPRT, 10, 0)
    psspy.branch_data(888, to_bus, ckt, [0, _i, _i, _i, _i, _i],
                      [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
    psspy.run(0, simtime + CB_fail_time, NPRT, 10, 0)
    psspy.branch_data(888, from_bus, ckt, [0, _i, _i, _i, _i, _i],
                      [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
    psspy.dscn(888)
    psspy.dscn(from_bus)
    psspy.dist_clear_fault(1)
    if arc_time > 0.01:                                     #ARC_time: AutoReClosing time
        ierr, simtime = psspy.dsrval('TIME', 0)
        psspy.run(0, simtime + arc_time, nprt, nplt, crtplt)
        psspy.recn(888)
        psspy.recn(from_bus)
        psspy.dist_bus_fault(888, 1, _f, [2, -0.2E+10])
        psspy.branch_data(888, from_bus, ckt, [1, _i, _i, _i, _i, _i],
                          [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
        psspy.branch_data(888, to_bus, ckt, [1, _i, _i, _i, _i, _i],
                          [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])

        psspy.run(0, simtime + arc_time + remote_fault_time, NPRT, 10, 0)
        psspy.branch_data(888, to_bus, ckt, [0, _i, _i, _i, _i, _i],
                      [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
        psspy.run(0, simtime + arc_time + CB_fail_time, nprt, nplt, crtplt)
        psspy.dscn(888)
        psspy.dscn(from_bus)
        psspy.branch_data(888, from_bus, ckt, [0, _i, _i, _i, _i, _i],
                          [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + 10.0, nprt, nplt, crtplt)
    # Close PSS/E
    psspy.close_powerflow()
    psspy.pssehalt_2()
# =======================================================
def LLG_fault_line_CB_Fail(output_files_dir, output_filename_suffix, case_file, Model_dyr, add_dyrs, from_bus, to_bus, ckt, remote_fault_time, CB_fail_time, arc_time):
    psspy.psseinit(80000)
    psspy.case(os.path.join(NEM_input_files_dir, case_file))
    network_modifications(case_file)
    ierr, V_base = psspy.busdat(from_bus, 'BASE')
    ierr, Fr_bus = psspy.notona(from_bus)
    ierr, To_bus = psspy.notona(to_bus)
    arc = ''
    if arc_time > 0.01:
        arc = '_ARC'
    output_filename = '%sLLG_fault_line_CBF%s_%s_%1.0fkV_Fr_%s_To_%s_%s' % (output_filename_suffix,arc, case_file[0:-4], V_base, Fr_bus[:12].strip(), To_bus[:12].strip(),ckt)
    dynamic_intialise(output_files_dir, output_filename, Model_dyr, add_dyrs)
    psspy.ltap(from_bus, to_bus, ckt, 0.2, 888, 'FLT_BUS', _f)
    psspy.dist_scmu_fault([0, 0, 2, 888], [0.0, 0.0, 0.0, 0.0])  #OPTIONS(3) = 2: line-to-line or line-to-line-to-ground fault
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + remote_fault_time, NPRT, 10, 0)
    psspy.branch_data(888, to_bus, ckt, [0, _i, _i, _i, _i, _i],
                      [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
    psspy.run(0, simtime + CB_fail_time, NPRT, 10, 0)
    psspy.branch_data(888, from_bus, ckt, [0, _i, _i, _i, _i, _i],
                      [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
    psspy.dscn(888)
    psspy.dscn(from_bus)
    psspy.dist_clear_fault(1)
    if arc_time > 0.01:                                     #ARC_time: AutoReClosing time
        ierr, simtime = psspy.dsrval('TIME', 0)
        psspy.run(0, simtime + arc_time, nprt, nplt, crtplt)
        psspy.recn(888)
        psspy.recn(from_bus)
        psspy.dist_scmu_fault([0, 0, 2, 888], [0.0, 0.0, 0.0, 0.0])  #OPTIONS(3) = 2: line-to-line or line-to-line-to-ground fault
        psspy.branch_data(888, from_bus, ckt, [1, _i, _i, _i, _i, _i],
                          [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
        psspy.branch_data(888, to_bus, ckt, [1, _i, _i, _i, _i, _i],
                          [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])

        psspy.run(0, simtime + arc_time + remote_fault_time, NPRT, 10, 0)
        psspy.branch_data(888, to_bus, ckt, [0, _i, _i, _i, _i, _i],
                      [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
        psspy.run(0, simtime + arc_time + CB_fail_time, nprt, nplt, crtplt)
        psspy.dscn(888)
        psspy.dscn(from_bus)
        psspy.branch_data(888, from_bus, ckt, [0, _i, _i, _i, _i, _i],
                          [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + 10.0, nprt, nplt, crtplt)
    # Close PSS/E
    psspy.close_powerflow()
    psspy.pssehalt_2()
# =======================================================
def LG_fault_line_CB_Fail(output_files_dir, output_filename_suffix, case_file, Model_dyr, add_dyrs, from_bus, to_bus, ckt, remote_fault_time, CB_fail_time, arc_time):
    psspy.psseinit(80000)
    psspy.case(os.path.join(NEM_input_files_dir, case_file))
    network_modifications(case_file)
    ierr, V_base = psspy.busdat(from_bus, 'BASE')
    ierr, Fr_bus = psspy.notona(from_bus)
    ierr, To_bus = psspy.notona(to_bus)
    arc = ''
    if arc_time > 0.01:
        arc = '_ARC'
    output_filename = '%sLG_fault_line_CBF%s_%s_%1.0fkV_Fr_%s_To_%s_%s' % (output_filename_suffix, arc, case_file[0:-4], V_base, Fr_bus[:12].strip(), To_bus[:12].strip(),ckt)
    dynamic_intialise(output_files_dir, output_filename, Model_dyr, add_dyrs)
    psspy.ltap(from_bus, to_bus, ckt, 0.2, 888, 'FLT_BUS', _f)
    psspy.dist_scmu_fault([0, 0, 1, 888], [0.0, 0.0, 0.0, 0.0])  #OPTIONS(3) = 1 line-to-ground fault
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + remote_fault_time, NPRT, 10, 0)
    psspy.branch_data(888, to_bus, ckt, [0, _i, _i, _i, _i, _i],
                      [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
    psspy.run(0, simtime + CB_fail_time, NPRT, 10, 0)
    psspy.branch_data(888, from_bus, ckt, [0, _i, _i, _i, _i, _i],
                      [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
    psspy.dscn(888)
    psspy.dscn(from_bus)
    psspy.dist_clear_fault(1)
    if arc_time > 0.01:                                     #ARC_time: AutoReClosing time
        ierr, simtime = psspy.dsrval('TIME', 0)
        psspy.run(0, simtime + arc_time, nprt, nplt, crtplt)
        psspy.recn(888)
        psspy.recn(from_bus)
        psspy.dist_scmu_fault([0, 0, 1, 888], [0.0, 0.0, 0.0, 0.0])  #OPTIONS(3) = 1 line-to-ground fault
        psspy.branch_data(888, from_bus, ckt, [1, _i, _i, _i, _i, _i],
                          [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
        psspy.branch_data(888, to_bus, ckt, [1, _i, _i, _i, _i, _i],
                          [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])

        psspy.run(0, simtime + arc_time + remote_fault_time, NPRT, 10, 0)
        psspy.branch_data(888, to_bus, ckt, [0, _i, _i, _i, _i, _i],
                      [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
        psspy.run(0, simtime + arc_time + CB_fail_time, nprt, nplt, crtplt)
        psspy.dscn(888)
        psspy.dscn(from_bus)
        psspy.branch_data(888, from_bus, ckt, [0, _i, _i, _i, _i, _i],
                          [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f])
    ierr, simtime = psspyoutput_filename.dsrval('TIME', 0)
    psspy.run(0, simtime + 10.0, nprt, nplt, crtplt)
   # Close PSS/E
    psspy.close_powerflow()
    psspy.pssehalt_2()
# =======================================================
def threeph_fault_trans(output_files_dir,output_filename_suffix, case_file, Model_dyr, add_dyrs, from_bus, to_bus, ckt, fault_time):
    psspy.psseinit(80000)
    psspy.case(os.path.join(NEM_input_files_dir, case_file))
    network_modifications(case_file)
    ierr, V_base = psspy.busdat(from_bus, 'BASE')
    ierr, Fr_bus = psspy.notona(from_bus)
    ierr, To_bus = psspy.notona(to_bus)
    output_filename = '%sThreeph_fault_trans_%s_%1.0fkV_Fr_%s_To_%s_%s' % (output_filename_suffix, case_file[0:-4],V_base,Fr_bus[:12].strip(),To_bus[:12].strip(),ckt)
    dynamic_intialise(output_files_dir, output_filename, Model_dyr, add_dyrs)
    psspy.dist_bus_fault(to_bus, 1, _f, [0.0, -0.2E+10])
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + fault_time, NPRT, 10, 0)
    psspy.two_winding_data_3(from_bus, to_bus, ckt, [0, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i],[_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f,_f, _f], _s)
    psspy.dist_clear_fault(1)
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + 10.0, NPRT, 10, 0)
    # Close PSS/E
    psspy.close_powerflow()
    psspy.pssehalt_2()
# =======================================================
def LLG_fault_trans(output_files_dir,output_filename_suffix, case_file, Model_dyr, add_dyrs, from_bus, to_bus, ckt, fault_time):
    psspy.psseinit(80000)
    psspy.case(os.path.join(NEM_input_files_dir, case_file))
    network_modifications(case_file)
    ierr, V_base = psspy.busdat(from_bus, 'BASE')
    ierr, Fr_bus = psspy.notona(from_bus)
    ierr, To_bus = psspy.notona(to_bus)
    output_filename = '%sLLG_fault_line_trans_%s_%1.0fkV_Fr_%s_To_%s_%s' % (output_filename_suffix, case_file[0:-4],V_base,Fr_bus[:12].strip(),To_bus[:12].strip(),ckt)
    dynamic_intialise(output_files_dir, output_filename, Model_dyr, add_dyrs)
    psspy.dist_scmu_fault([0, 0, 2, from_bus], [0.0, 0.0, 0.0, 0.0])  #OPTIONS(3) = 2 line-to-line or line-to-line-to-ground fault
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + fault_time, NPRT, 10, 0)
    psspy.two_winding_data_3(from_bus, to_bus, ckt, [0, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i],
                             [_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f,
                              _f, _f], _s)
    psspy.dist_clear_fault(1)
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + 10.0, nprt, nplt, crtplt)
    # Close PSS/E
    psspy.close_powerflow()
    psspy.pssehalt_2()
# =======================================================
def LG_fault_trans(output_files_dir,output_filename_suffix,case_file, Model_dyr, add_dyrs, from_bus, to_bus, ckt, fault_time):
    psspy.psseinit(80000)
    psspy.case(os.path.join(NEM_input_files_dir, case_file))
    network_modifications(case_file)
    ierr, V_base = psspy.busdat(from_bus, 'BASE')
    ierr, Fr_bus = psspy.notona(from_bus)
    ierr, To_bus = psspy.notona(to_bus)
    output_filename = '%sLG_fault_trans_%s_%1.0fkV_Fr_%s_To_%s_%s' % (output_filename_suffix,case_file[0:-4],V_base,Fr_bus[:12].strip(),To_bus[:12].strip(),ckt)
    dynamic_intialise(output_files_dir, output_filename, Model_dyr, add_dyrs)
    psspy.dist_scmu_fault([0, 0, 1, from_bus], [0.0, 0.0, 0.0, 0.0]) #OPTIONS(3) = 1 line-to-ground fault
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + fault_time, NPRT, 10, 0)
    psspy.two_winding_data_3(from_bus, to_bus, ckt, [0, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i, _i],[_f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f, _f,_f, _f], _s)
    psspy.dist_clear_fault(1)
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + 10.0, nprt, nplt, crtplt)
    # Close PSS/E
    psspy.close_powerflow()
    psspy.pssehalt_2()
# =======================================================
def threeph_fault_3wnd_trans(output_files_dir,output_filename_suffix, case_file, Model_dyr, add_dyrs, from_bus, to_bus,to_bus2, ckt, fault_time):
    psspy.psseinit(80000)
    psspy.case(os.path.join(NEM_input_files_dir, case_file))
    network_modifications(case_file)
    ierr, V_base = psspy.busdat(from_bus, 'BASE')
    ierr, Fr_bus = psspy.notona(from_bus)
    ierr, To_bus = psspy.notona(to_bus)
    output_filename = '%sThreeph_fault_3wnd_trans_%s_%1.0fkV_Fr_%s_To_%s_%s' % (output_filename_suffix, case_file[0:-4],V_base,Fr_bus[:12].strip(),To_bus[:12].strip(),ckt)
    dynamic_intialise(output_files_dir, output_filename, Model_dyr, add_dyrs)
    psspy.dist_bus_fault(from_bus, 1, _f, [0.0, -0.2E+4])
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + fault_time, NPRT, 10, 0)
    psspy.three_wnd_imped_data_4(from_bus, to_bus, to_bus2, ckt, [_i, _i, _i, _i, _i, _i, _i, 0, _i, _i, _i, _i, _i])
    psspy.dist_clear_fault(1)
    ierr, simtime = psspy.dsrval('TIME', 0)
    psspy.run(0, simtime + 10.0, nprt, nplt, crtplt)
    # Close PSS/E
    psspy.close_powerflow()
    psspy.pssehalt_2()

# =======================================================
