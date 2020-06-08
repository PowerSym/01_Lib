import os, sys, math, re, glob
import shutil
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
MisMatch_Tol = 0.001
cwd = os.getcwd()
Plotting_tools = os.path.join(os.path.dirname(cwd),'Plotting_tools')
sys.path.append(Plotting_tools)
import Plotting as pl
PSCAD_PATH = r"C:\Program Files\Python37\Lib\site-packages\mhrc"
sys.path.append(PSCAD_PATH)
import usercode37 as uc
import win32com.client
from win32com.client.gencache import EnsureDispatch as Dispatch
from automation.utilities.word import Word
from automation.utilities.file import File
from automation.utilities.mail import Mail
import automation.controller
import automation.certificate
import automation.pscad
# =======================================================
def Initial(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'Initial'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,Vref_pu_ID,p_ppc_ref_pu_ID, Vsmib_pu_ID = Component_IDs
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom,wfbase_MW, test_cases['SCR'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
            uc.dirCreateClean(overlay_dir, [".pdf"])
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=20, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(V_1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(P_1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            ws.create_simulation_set(prj_name)
            ss = ws.simulation_set(prj_name)
            ss.add_tasks(prj_name)
            if 'y' in save_option.lower():
                project.save_as(output_filename)
                shutil.copy(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename),output_files_dir)
                uc.dirCreateClean(os.path.dirname(PSCAD_input_file), [".bakx", ".psmx"])
                # os.remove(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename))
            ss.run()
            pscad.quit()
            uc.copy(src_folder, output_files_dir, [".out", ".inf"])
            if overlay != '':
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def APT(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs,Ref_ID):
    Test = 'APT'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,Vref_pu_ID,p_ppc_ref_pu_ID, Vsmib_pu_ID = Component_IDs
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            refs = []
            for ref in (re.split('[ :,;]', test_cases['Ref 5'][i])):
                refs.append(float(ref))
            time_duration = max(refs[0:10])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom,wfbase_MW, test_cases['SCR'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
            uc.dirCreateClean(overlay_dir, [".pdf"])
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(V_1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(P_1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            # Main Program:
            pt_1,pt_2,pt_3,pt_4,pt_5,pt_6,pt_7,pt_8,pt_9,pt_10,P_1,P_2,P_3,P_4,P_5,P_6,P_7,P_8,P_9,P_10 = refs
            Main.user_cmp(Ref_ID).set_parameters(OutMod=2,Pnum=int(10),pt_1=pt_1,pt_2=pt_2,pt_3=pt_3,pt_4=pt_4,pt_5=pt_5,pt_6=pt_6,pt_7=pt_7,pt_8=pt_8,pt_9=pt_9,pt_10=pt_10,
                                                     P_1=P_1,P_2=P_2,P_3=P_3,P_4=P_4,P_5=P_5,P_6=P_6,P_7=P_7,P_8=P_8,P_9=P_9,P_10=P_10)
            ws.create_simulation_set(prj_name)
            ss = ws.simulation_set(prj_name)
            ss.add_tasks(prj_name)
            if 'y' in save_option.lower():
                project.save_as(output_filename)
                shutil.copy(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename),output_files_dir)
                uc.dirCreateClean(os.path.dirname(PSCAD_input_file), [".bakx", ".psmx"])
                # os.remove(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename))
            ss.run()
            pscad.quit()
            uc.copy(src_folder, output_files_dir, [".out", ".inf"])
            if overlay != '':
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def Fault_SMIB(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs,Ref_ID):
    Test = 'Fault_SMIB'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,Vref_pu_ID,p_ppc_ref_pu_ID, Vsmib_pu_ID = Component_IDs
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            time_duration = float(test_cases['Ref 8'][i])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom,wfbase_MW, test_cases['SCR'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
            uc.dirCreateClean(overlay_dir, [".pdf"])
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(V_1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(P_1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            # Main Program:
            Timed_Fault_Logic = int(Ref_ID[0])
            Vdip_ID = int(Ref_ID[1])
            Main.user_cmp(Timed_Fault_Logic).set_parameters(TF=float(test_cases['Ref 6'][i]), DF=float(test_cases['Ref 7'][i]))
            if float(test_cases['Ref 5'][i]) == 0:
                test_cases['Ref 5'][i] =0.01
            Main.user_cmp(Vdip_ID).set_parameters(Value=float(test_cases['Ref 5'][i]))
            ws.create_simulation_set(prj_name)
            ss = ws.simulation_set(prj_name)
            ss.add_tasks(prj_name)
            if 'y' in save_option.lower():
                project.save_as(output_filename)
                shutil.copy(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename),output_files_dir)
                uc.dirCreateClean(os.path.dirname(PSCAD_input_file), [".bakx", ".psmx"])
                # os.remove(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename))
            ss.run()
            pscad.quit()
            uc.copy(src_folder, output_files_dir, [".out", ".inf"])
            if overlay != '':
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def FCT_PLB(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs,FCT_PLB_ID,APT_ID):
    Test = 'FCT_PLB'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,Vref_pu_ID,p_ppc_ref_pu_ID, Vsmib_pu_ID = Component_IDs
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            ref1s = []
            for ref1 in (re.split('[ :,;]', test_cases['Ref 5'][i])):
                ref1s.append(float(ref1))
            time_duration = max(ref1s[0:10])
            if int(test_cases['Ref 6'][i]) == 3:
                Pctrl_mode = 'FSM'
            elif int(test_cases['Ref 6'][i]) == 1:
                Pctrl_mode = 'PCtrl'
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom,wfbase_MW, test_cases['SCR'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_%s_PSCAD' % (test_cases['Case No'][i],Test,Pctrl_mode)
            overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
            uc.dirCreateClean(overlay_dir, [".pdf"])
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            if time_duration >= 200:
                project.set_parameters(sample_step=50000)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(V_1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(P_1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            # Main Program:
            Main.user_cmp(PPC_ID).set_parameters(Pctrl=int(test_cases['Ref 6'][i]))
            X1,X2,X3,X4,X5,X6,X7,X8,X9,X10,Y1,Y2,Y3,Y4,Y5,Y6,Y7,Y8,Y9,Y10 = ref1s
            Main.user_cmp(FCT_PLB_ID).set_parameters(Mode=1, N=10,X1=X1,X2=X2,X3=X3,X4=X4,X5=X5,X6=X6,X7=X7,X8=X8,X9=X9,X10=X10,Y1=Y1,Y2=Y2,Y3=Y3,Y4=Y4,Y5=Y5,Y6=Y6,Y7=Y7,Y8=Y8,Y9=Y9,Y10=Y10)
            if test_cases['Ref 7'][i] != ' ':
                ref2s = []
                for ref2 in (re.split('[ :,;]', test_cases['Ref 7'][i])):
                    ref2s.append(float(ref2))
                pt_1,pt_2,pt_3,pt_4,pt_5,pt_6,pt_7,pt_8,pt_9,pt_10,P_1,P_2,P_3,P_4,P_5,P_6,P_7,P_8,P_9,P_10 = ref2s
                Main.user_cmp(APT_ID).set_parameters(OutMod=2,Pnum=int(10),pt_1=pt_1,pt_2=pt_2,pt_3=pt_3,pt_4=pt_4,pt_5=pt_5,pt_6=pt_6,pt_7=pt_7,pt_8=pt_8,pt_9=pt_9,pt_10=pt_10,
                                                         P_1=P_1,P_2=P_2,P_3=P_3,P_4=P_4,P_5=P_5,P_6=P_6,P_7=P_7,P_8=P_8,P_9=P_9,P_10=P_10)
            ws.create_simulation_set(prj_name)
            ss = ws.simulation_set(prj_name)
            ss.add_tasks(prj_name)
            if 'y' in save_option.lower():
                project.save_as(output_filename)
                shutil.copy(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename),output_files_dir)
                uc.dirCreateClean(os.path.dirname(PSCAD_input_file), [".bakx", ".psmx"])
                # os.remove(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename))
            ss.run()
            pscad.quit()
            uc.copy(src_folder, output_files_dir, [".out", ".inf"])
            if overlay != '':
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def VCT(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs,Ref_ID):
    Test = 'VCT'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,Vref_pu_ID,p_ppc_ref_pu_ID, Vsmib_pu_ID = Component_IDs
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            refs = []
            for ref in (re.split('[ :,;]', test_cases['Ref 5'][i])):
                refs.append(float(ref))
            time_duration = max(refs[0:10])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom,wfbase_MW, test_cases['SCR'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
            uc.dirCreateClean(overlay_dir, [".pdf"])
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(V_1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(P_1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            # Main Program:
            vt_1,vt_2,vt_3,vt_4,vt_5,vt_6,vt_7,vt_8,vt_9,vt_10,V_1,V_2,V_3,V_4,V_5,V_6,V_7,V_8,V_9,V_10 = refs
            Main.user_cmp(Ref_ID).set_parameters(OutMod=2,Vnum=int(10),vt_1=vt_1,vt_2=vt_2,vt_3=vt_3,vt_4=vt_4,vt_5=vt_5,vt_6=vt_6,vt_7=vt_7,vt_8=vt_8,vt_9=vt_9,vt_10=vt_10,
                                                     V_1=V_1,V_2=V_2,V_3=V_3,V_4=V_4,V_5=V_5,V_6=V_6,V_7=V_7,V_8=V_8,V_9=V_9,V_10=V_10)
            ws.create_simulation_set(prj_name)
            ss = ws.simulation_set(prj_name)
            ss.add_tasks(prj_name)
            if 'y' in save_option.lower():
                project.save_as(output_filename)
                shutil.copy(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename),output_files_dir)
                uc.dirCreateClean(os.path.dirname(PSCAD_input_file), [".bakx", ".psmx"])
                # os.remove(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename))
            ss.run()
            pscad.quit()
            uc.copy(src_folder, output_files_dir, [".out", ".inf"])
            if overlay != '':
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def VDT(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs,Ref_ID):
    Test = 'VDT'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,Vref_pu_ID,p_ppc_ref_pu_ID, Vsmib_pu_ID = Component_IDs
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            refs = []
            for ref in (re.split('[ :,;]', test_cases['Ref 5'][i])):
                refs.append(float(ref))
            time_duration = max(refs[0:10])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom,wfbase_MW, test_cases['SCR'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
            uc.dirCreateClean(overlay_dir, [".pdf"])
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(V_1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(P_1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            # Main Program:
            X1,X2,X3,X4,X5,X6,X7,X8,X9,X10,Y1,Y2,Y3,Y4,Y5,Y6,Y7,Y8,Y9,Y10 = refs
            Main.user_cmp(Ref_ID).set_parameters(Mode=1, N=int(10),X1=X1,X2=X2,X3=X3,X4=X4,X5=X5,X6=X6,X7=X7,X8=X8,X9=X9,X10=X10,Y1=Y1,Y2=Y2,Y3=Y3,Y4=Y4,Y5=Y5,Y6=Y6,Y7=Y7,Y8=Y8,Y9=Y9,Y10=Y10)
            ws.create_simulation_set(prj_name)
            ss = ws.simulation_set(prj_name)
            ss.add_tasks(prj_name)
            if 'y' in save_option.lower():
                project.save_as(output_filename)
                shutil.copy(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename),output_files_dir)
                uc.dirCreateClean(os.path.dirname(PSCAD_input_file), [".bakx", ".psmx"])
                # os.remove(os.path.join(os.path.dirname(PSCAD_input_file), '%s.pscx'%output_filename))
            ss.run()
            pscad.quit()
            uc.copy(src_folder, output_files_dir, [".out", ".inf"])
            if overlay != '':
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version):
    controller = automation.controller.Controller()
    pscad = controller.launch(pscad_ver=pscad_ver, options={'silence': silence},settings={'cl_use_advanced': cl_use_advanced,'cl_auto_renewal': cl_auto_renewal,'fortran_version': fortran_version})
    ws = pscad.workspace()
    pscad.load(PSCAD_input_file)
    projects = ws.projects()
    prj_name = projects[1]['name']
    project = pscad.project(prj_name)
    project.focus()
    return pscad,ws,project,prj_name
# =======================================================
def RL_calculation(VPCCnom,wfbase_MW,scr,xr):
    sc_pcc_pu_Sbase = (scr * wfbase_MW) / 100.0  # 100MVA System base
    Rgrid_pu = math.sqrt(((1.0 / sc_pcc_pu_Sbase) ** 2) / (xr ** 2 + 1.0))
    Xgrid_pu = Rgrid_pu * xr
    Rgrid_ohm = Rgrid_pu * (VPCCnom * VPCCnom / 100)
    Xgrid_ohm = Xgrid_pu * (VPCCnom * VPCCnom / 100)
    Lgrid_henry = Xgrid_ohm / (2 * 50 * math.pi)
    return Rgrid_ohm,Lgrid_henry
# =======================================================
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
    models = ['GSVW','GSPQ','GSLH','GSME','GSVF','VWPO','VWRE','WTG_GSV8','PPC_V8']
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
                            if isinstance(cell_obj, str):
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
                            if isinstance(cell_obj, str):
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

        xl_sheet = xl_workbook.sheet_by_name('USRCD1')
        dyrs = ['MSU', 'FCT_01', 'FCT_02', 'FCT_03', 'FCT_04', 'FCT_05', 'FCT_06', 'FCT_07', 'FCT_08']
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
                        if isinstance(cell_obj, str):
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
                        if isinstance(cell_obj, str):
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
            if (isinstance(cell_obj, str)):
                if len(re.findall(("\d+\.\d+"), cell_obj)) >0:
                    cell_obj = float(re.findall("\d+\.\d+", cell_obj)[0])
                elif len(re.findall(("\d+"), cell_obj)) > 0:
                    cell_obj = float(re.findall("\d+", cell_obj)[0])
                else:
                    print('This is not float type')
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
def copy_out_inf(src_folder,destination_fld):
    for inf_f in os.listdir(src_folder):
        if inf_f.endswith(".inf"):
            for out_f in os.listdir(src_folder):
                if out_f[0:-7] == inf_f[0:-4]:
                    shutil.copy(os.path.join(src_folder,inf_f),destination_fld)
                    shutil.copy(os.path.join(src_folder,out_f),destination_fld)
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
    try:
        os.mkdir(path)
    except OSError:
        pass
    filelist = os.listdir(path)
    for type in fileTypes:
        for f in filelist:
            if f.endswith(type):
                os.remove(os.path.join(path, f))
# ========================================================
def merge_pdf(dir_to_merge):
    copy_dir = os.path.dirname(dir_to_merge)
    input_streams = []
    input_files = [input_file for input_file in os.listdir(dir_to_merge) if (input_file.endswith(".pdf") and os.path.getsize(os.path.join(dir_to_merge, input_file))>2000)] #only merge the pdf file with more than 1K size
    if input_files != []:
        try:
            for input_file in input_files:
                input_streams.append(open(os.path.join(dir_to_merge, input_file), 'rb'))
            writer = PdfFileWriter()
            for reader in map(PdfFileReader, input_streams):
                for n in range(reader.getNumPages()):
                    writer.addPage(reader.getPage(n))
            output_stream = dir_to_merge+'\\'+ dir_to_merge.split(os.sep)[-1] +'_Plots.pdf'
            output_stream = open(output_stream, 'w+b')
            writer.write(output_stream)
        finally:
            for f in input_streams:
                f.close()
    for file in os.listdir(dir_to_merge):
        if file.endswith(".pdf") and (not file.endswith(dir_to_merge.split(os.sep)[-1] +'_Plots.pdf')):
            os.remove(os.path.join(dir_to_merge, file))
    pdf_file = dir_to_merge + '\\' + dir_to_merge.split(os.sep)[-1] + '_Plots.pdf'
    copy_pdf_file = copy_dir + '\\' + dir_to_merge.split(os.sep)[-1] + '_Plots.pdf'
    # shutil.copy2(pdf_file, copy_dir)
    # command = "xcopy \"" + pdf_file + "\" " + "\"" + copy_pdf_file + "*\""
    # os.system(command)

# ========================================================
def read_output_channels(excel_file, sheet_name):
    import xlrd
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
    chandata = {}
    for col_idx in range(len(row_title)):
        col_idx_data = []
        for row_idx in range(1, xl_sheet.nrows):
            cell_obj = (xl_sheet.cell(row_idx, col_idx)).value
            col_idx_data.append(cell_obj)
        chanid[col_idx] = row_title[col_idx].value
        chandata[row_title[col_idx].value] = col_idx_data
    IDs = chandata['Instance']
    Chans = chandata['Symbol']
    return IDs,Chans
# ========================================================
def plot_all_in_dir(dir_to_plot,Test,Project_info,wfbase_MW):
    dirCreateClean(dir_to_plot, [".pdf"])
    filenames = glob.glob(os.path.join(dir_to_plot,'*.inf'))
    plot_Vdroop, plot_Qdroop = Droop_parameters(Project_info, wfbase_MW)
    for filename in filenames:
        if 'APT' in Test:
            plot = pl.Plotting()
            plot.read_data(filename)
            plot.subplot_spec(0, (0, 'V_RMS_WTG1'), title='WTG Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(0, (0, 'V_RMS_WTG2'))
            plot.subplot_spec(1, (0, 'V_PCC_RMS_meas'), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(1, (0, 'Vref'), scale=1.0)
            plot.subplot_spec(2, (0, 'P_WTG1'), title='Active Power at WTG', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(2, (0, 'P_WTG2'), scale=1.0)
            plot.subplot_spec(3, (0, 'P_PCC_meas'), title='Active Power at POC', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(3, (0, 'Pref'),scale=wfbase_MW, offset=0.0)
            plot.subplot_spec(4, (0, 'Q_WTG1'), title='Reactive Power at WTG', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
            plot.subplot_spec(4, (0, 'Q_WTG2'), scale=1.0)
            plot.subplot_spec(5, (0, 'Q_PCC_meas'), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
        else:
            plot = pl.Plotting()
            plot.read_data(filename)
            plot.subplot_spec(0, (0, 'V_RMS_WTG1'), title='WTG Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(0, (0, 'V_RMS_WTG2'))
            plot.subplot_spec(1, (0, 'V_PCC_RMS_meas'), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0) # plot_Vdroop = [26, 15, 0.04, 72.2, 0.005] [POC_V_ref_Channel_id,Q_POC_Channel_id, QDROOP, wfbase_MW]
            plot.subplot_spec(1, (0, 'Vref'), scale=1.0)
            plot.subplot_spec(2, (0, 'P_WTG1'), title='Active Power at WTG', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(2, (0, 'P_WTG2'), scale=1.0)
            plot.subplot_spec(3, (0, 'P_PCC_meas'), title='Active Power at POC', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(3, (0, 'Pref'),scale=wfbase_MW, offset=0.0)
            plot.subplot_spec(4, (0, 'Q_WTG1'), title='Reactive Power at WTG', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
            plot.subplot_spec(4, (0, 'Q_WTG2'), scale=1.0)
            plot.subplot_spec(5, (0, 'Q_PCC_meas'), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
        plot.plot(figname=os.path.splitext(filename)[0], show=0)
    merge_pdf(dir_to_plot)
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