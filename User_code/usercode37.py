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
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            output_filename = '%s' % (os.path.basename(PSCAD_input_file)[0:-9])
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=20, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(V_1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(P_1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) ==0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName = 0, Name = New_Symbols[j])
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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def APT(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'APT'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            N, refs, time_duration = get_reference(test_cases['Ref 6'][i])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(Y1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(Y1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            Main.user_cmp(int(PPC_ID)).set_parameters(Pctrl=1)
            # Main Program:
            Main.user_cmp(p_ppc_ref_pu_ID).set_parameters(N=N,X1=refs[0][0],X2=refs[1][0],X3=refs[2][0],X4=refs[3][0],X5=refs[4][0],X6=refs[5][0],X7=refs[6][0],X8=refs[7][0],X9=refs[8][0],X10=refs[9][0],
                                                          Y1=refs[0][1],Y2=refs[1][1],Y3=refs[2][1],Y4=refs[3][1],Y5=refs[4][1],Y6=refs[5][1],Y7=refs[6][1],Y8=refs[7][1],Y9=refs[8][1],Y10=refs[9][1])
            if 'interpolate' in test_cases['Ref 7'][i].lower():
                Main.user_cmp(p_ppc_ref_pu_ID).set_parameters(Mode=0)
            else:
                Main.user_cmp(p_ppc_ref_pu_ID).set_parameters(Mode=1)
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) == 0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName=0, Name=New_Symbols[j])
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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def Fault_SMIB(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'Fault_SMIB'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            time_duration = float(test_cases['Ref 10'][i])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(Y1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(Y1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            # Main Program:
            Timed_Fault_Logic = int(Fault_IDs[0])
            Vdip_ID = int(Fault_IDs[1])
            Fault_type_ID = int(Fault_IDs[2])
            Main.user_cmp(Timed_Fault_Logic).set_parameters(TF=float(test_cases['Ref 7'][i]), DF=float(test_cases['Ref 8'][i]))
            if float(test_cases['Ref 6'][i]) == 0:
                test_cases['Ref 6'][i] =0.01
            Main.user_cmp(Vdip_ID).set_parameters(Value=float(test_cases['Ref 6'][i]))
            if 'a' in test_cases['Ref 9'][i].lower():
                Main.user_cmp(Fault_type_ID).set_parameters(A=1)
            else:
                Main.user_cmp(Fault_type_ID).set_parameters(A=0)
            if 'b' in test_cases['Ref 9'][i].lower():
                Main.user_cmp(Fault_type_ID).set_parameters(B=1)
            else:
                Main.user_cmp(Fault_type_ID).set_parameters(B=0)
            if 'c' in test_cases['Ref 9'][i].lower():
                Main.user_cmp(Fault_type_ID).set_parameters(C=1)
            else:
                Main.user_cmp(Fault_type_ID).set_parameters(C=0)
            if 'g' in test_cases['Ref 9'][i].lower():
                Main.user_cmp(Fault_type_ID).set_parameters(G=1)
            else:
                Main.user_cmp(Fault_type_ID).set_parameters(G=0)
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) == 0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName=0, Name=New_Symbols[j])
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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def FCT_PLB(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'FCT_PLB'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            N, refs, time_duration = get_reference(test_cases['Ref 6'][i])
            Pctrl_mode = test_cases['Ref 8'][i]
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_%s_PSCAD' % (test_cases['Case No'][i],Test,Pctrl_mode)
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
            if 'pctrl' in Pctrl_mode.lower():
                Main.user_cmp(PPC_ID).set_parameters(Pctrl=1)
            elif 'fsm1' in Pctrl_mode.lower():
                Main.user_cmp(PPC_ID).set_parameters(Pctrl=2)
            else:
                Main.user_cmp(PPC_ID).set_parameters(Pctrl=3)
            Main.user_cmp(FCT_PLB_ID).set_parameters(N=N,X1=refs[0][0],X2=refs[1][0],X3=refs[2][0],X4=refs[3][0],X5=refs[4][0],X6=refs[5][0],X7=refs[6][0],X8=refs[7][0],X9=refs[8][0],X10=refs[9][0],
                                                          Y1=refs[0][1],Y2=refs[1][1],Y3=refs[2][1],Y4=refs[3][1],Y5=refs[4][1],Y6=refs[5][1],Y7=refs[6][1],Y8=refs[7][1],Y9=refs[8][1],Y10=refs[9][1])
            if 'interpolate' in test_cases['Ref 7'][i].lower():
                Main.user_cmp(FCT_PLB_ID).set_parameters(Mode=0)
            else:
                Main.user_cmp(FCT_PLB_ID).set_parameters(Mode=1)
            if test_cases['Ref 9'][i] != '':
                N2, ref2s, time_duration2 = get_reference(test_cases['Ref 9'][i])
                Main.user_cmp(Fallback_ID[0]).set_parameters(N=N2, Mode=1, X1=ref2s[0][0], X2=ref2s[1][0], X3=ref2s[2][0],
                                                         X4=ref2s[3][0], X5=ref2s[4][0], X6=ref2s[5][0], X7=ref2s[6][0],
                                                         X8=ref2s[7][0], X9=ref2s[8][0], X10=ref2s[9][0],
                                                         Y1=ref2s[0][1], Y2=ref2s[1][1], Y3=ref2s[2][1], Y4=ref2s[3][1],
                                                         Y5=ref2s[4][1], Y6=ref2s[5][1], Y7=ref2s[6][1], Y8=ref2s[7][1],
                                                         Y9=ref2s[8][1], Y10=ref2s[9][1])
                Main.slider(Fallback_ID[1],Fallback_ID[2]).value(float(test_cases['Ref 10'][i]))
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) == 0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName=0, Name=New_Symbols[j])
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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def VCT(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'VCT'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            N, refs, time_duration = get_reference(test_cases['Ref 6'][i])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(Y1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(Y1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            # Main Program:
            Main.user_cmp(Vref_pu_ID).set_parameters(N=N,X1=refs[0][0],X2=refs[1][0],X3=refs[2][0],X4=refs[3][0],X5=refs[4][0],X6=refs[5][0],X7=refs[6][0],X8=refs[7][0],X9=refs[8][0],X10=refs[9][0],
                                                          Y1=refs[0][1],Y2=refs[1][1],Y3=refs[2][1],Y4=refs[3][1],Y5=refs[4][1],Y6=refs[5][1],Y7=refs[6][1],Y8=refs[7][1],Y9=refs[8][1],Y10=refs[9][1])
            if 'interpolate' in test_cases['Ref 7'][i].lower():
                Main.user_cmp(Vref_pu_ID).set_parameters(Mode=0)
            else:
                Main.user_cmp(Vref_pu_ID).set_parameters(Mode=1)
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) == 0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName=0, Name=New_Symbols[j])

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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def VDT(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'VDT'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            N, refs, time_duration = get_reference(test_cases['Ref 6'][i])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(V_1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(P_1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            # Main Program:
            Main.user_cmp(Vsmib_pu_ID).set_parameters(N=N,X1=refs[0][0],X2=refs[1][0],X3=refs[2][0],X4=refs[3][0],X5=refs[4][0],X6=refs[5][0],X7=refs[6][0],X8=refs[7][0],X9=refs[8][0],X10=refs[9][0],
                                                          Y1=refs[0][1],Y2=refs[1][1],Y3=refs[2][1],Y4=refs[3][1],Y5=refs[4][1],Y6=refs[5][1],Y7=refs[6][1],Y8=refs[7][1],Y9=refs[8][1],Y10=refs[9][1])
            if 'interpolate' in test_cases['Ref 7'][i].lower():
                Main.user_cmp(Vsmib_pu_ID).set_parameters(Mode=0)
            else:
                Main.user_cmp(Vsmib_pu_ID).set_parameters(Mode=1)
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) == 0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName=0, Name=New_Symbols[j])
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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def FRB(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'FRB'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            N, refs, time_duration = get_reference(test_cases['Ref 6'][i])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
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
            Main.user_cmp(Fallback_ID[0]).set_parameters(N=N,X1=refs[0][0],X2=refs[1][0],X3=refs[2][0],X4=refs[3][0],X5=refs[4][0],X6=refs[5][0],X7=refs[6][0],X8=refs[7][0],X9=refs[8][0],X10=refs[9][0],
                                                          Y1=refs[0][1],Y2=refs[1][1],Y3=refs[2][1],Y4=refs[3][1],Y5=refs[4][1],Y6=refs[5][1],Y7=refs[6][1],Y8=refs[7][1],Y9=refs[8][1],Y10=refs[9][1])
            if 'interpolate' in test_cases['Ref 7'][i].lower():
                Main.user_cmp(Fallback_ID[0]).set_parameters(Mode=0)
            else:
                Main.user_cmp(Fallback_ID[0]).set_parameters(Mode=1)
            Main.slider(Fallback_ID[1],Fallback_ID[2]).value(float(test_cases['Ref 8'][i]))
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) == 0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName=0, Name=New_Symbols[j])
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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def RPT(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'RPT'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            N, refs, time_duration = get_reference(test_cases['Ref 6'][i])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(Y1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(Y1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            Main.user_cmp(int(PPC_ID)).set_parameters(Pctrl=1,Qctrl = 1)
            # Main Program:
            Main.user_cmp(Q_ppc_ref_pu_ID).set_parameters(N=N,X1=refs[0][0],X2=refs[1][0],X3=refs[2][0],X4=refs[3][0],X5=refs[4][0],X6=refs[5][0],X7=refs[6][0],X8=refs[7][0],X9=refs[8][0],X10=refs[9][0],
                                                          Y1=refs[0][1],Y2=refs[1][1],Y3=refs[2][1],Y4=refs[3][1],Y5=refs[4][1],Y6=refs[5][1],Y7=refs[6][1],Y8=refs[7][1],Y9=refs[8][1],Y10=refs[9][1])
            if 'interpolate' in test_cases['Ref 7'][i].lower():
                Main.user_cmp(Q_ppc_ref_pu_ID).set_parameters(Mode=0)
            else:
                Main.user_cmp(Q_ppc_ref_pu_ID).set_parameters(Mode=1)
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) == 0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName=0, Name=New_Symbols[j])
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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def PFT(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'PFT'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            N, refs, time_duration = get_reference(test_cases['Ref 6'][i])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(Y1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(Y1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            Main.user_cmp(int(PPC_ID)).set_parameters(Pctrl=1,Qctrl = 2)
            # Main Program:
            Main.user_cmp(PF_ppc_ref_ID).set_parameters(N=N,X1=refs[0][0],X2=refs[1][0],X3=refs[2][0],X4=refs[3][0],X5=refs[4][0],X6=refs[5][0],X7=refs[6][0],X8=refs[7][0],X9=refs[8][0],X10=refs[9][0],
                                                          Y1=refs[0][1],Y2=refs[1][1],Y3=refs[2][1],Y4=refs[3][1],Y5=refs[4][1],Y6=refs[5][1],Y7=refs[6][1],Y8=refs[7][1],Y9=refs[8][1],Y10=refs[9][1])
            if 'interpolate' in test_cases['Ref 7'][i].lower():
                Main.user_cmp(PF_ppc_ref_ID).set_parameters(Mode=0)
            else:
                Main.user_cmp(PF_ppc_ref_ID).set_parameters(Mode=1)
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) == 0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName=0, Name=New_Symbols[j])
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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def P_curtail(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'P_curtail'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            N, refs, time_duration = get_reference(test_cases['Ref 6'][i])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(time_duration=time_duration, PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(Y1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(Y1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            Main.user_cmp(int(PPC_ID)).set_parameters(Pctrl=1)
            # Main Program:
            Main.user_cmp(P_curt_ref_ID).set_parameters(N=N,X1=refs[0][0],X2=refs[1][0],X3=refs[2][0],X4=refs[3][0],X5=refs[4][0],X6=refs[5][0],X7=refs[6][0],X8=refs[7][0],X9=refs[8][0],X10=refs[9][0],
                                                          Y1=refs[0][1],Y2=refs[1][1],Y3=refs[2][1],Y4=refs[3][1],Y5=refs[4][1],Y6=refs[5][1],Y7=refs[6][1],Y8=refs[7][1],Y9=refs[8][1],Y10=refs[9][1])
            if 'interpolate' in test_cases['Ref 7'][i].lower():
                Main.user_cmp(P_curt_ref_ID).set_parameters(Mode=0)
            else:
                Main.user_cmp(P_curt_ref_ID).set_parameters(Mode=1)
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) == 0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName=0, Name=New_Symbols[j])
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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
                uc.copy(src_folder, overlay_dir, [".out", ".inf"])
            uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    if 'y' in Plotting_option.lower():
        plot_all_in_dir(output_files_dir,Test,Project_info,wfbase_MW)
# =======================================================
def PQT(PSCAD_input_file,study_cases_file,python_file,results, Project_info, wfbase_MW,VPCCnom, settings,Component_IDs):
    Test = 'PQT'
    pscad_ver, silence, cl_use_advanced, cl_auto_renewal, fortran_version,time_step,sample_step,Plotting_option,save_option,overlay = settings
    PPC_ID, Rgrid_ID,Lgrid_ID,p_ppc_ref_pu_ID,F_ppc_ref_pu_ID,P_curt_ref_ID,Q_ppc_ref_pu_ID,PF_ppc_ref_ID,Vref_pu_ID,Vsmib_pu_ID,Fault_IDs,FCT_PLB_ID,Fallback_ID = Component_IDs
    IDs,Chans,Enables,Symbols,New_Symbols = read_output_channels(study_cases_file, 'Output_channels')
    output_files_dir = os.path.join(results, Test)
    src_folder = '%s.if15_x86' % PSCAD_input_file[0:-5]
    uc.dirCreateClean(output_files_dir, [".pdf"])
    uc.dirCreateClean(src_folder, [".out", ".inf", ".infx", ".pdf"])
    # uc.dirCreateClean(output_files_dir, [".out", ".inf", ".infx", ".pdf"])
    test_cases = uc.read_excel_sheet(study_cases_file,os.path.basename(python_file)[0:-3])
    for i in range(len(test_cases['Case No'])):
        if Test in str(test_cases['Name 1'][i]) and 'y' in str(test_cases['Execute'][i]).lower():
            N, refs, time_duration = get_reference(test_cases['Ref 6'][i])
            Rgrid_ohm, Lgrid_henry = uc.RL_calculation(VPCCnom, test_cases['sc_pcc_pu_Sbase'][i], test_cases['XR'][i])
            output_filename = 'Case_%02d_%s_PSCAD' % (test_cases['Case No'][i],Test)
            pscad,ws,project,prj_name = uc.PSCAD_start(PSCAD_input_file,pscad_ver,silence,cl_use_advanced,cl_auto_renewal,fortran_version)
            project.set_parameters(PlotType=1, output_filename=output_filename, sample_step=sample_step,time_step=time_step)
            if time_duration >= 200:
                project.set_parameters(sample_step=50000)
            Main = project.user_canvas('Main')
            Main.user_cmp(int(Rgrid_ID)).set_parameters(Value=Rgrid_ohm)
            Main.user_cmp(int(Lgrid_ID)).set_parameters(Value=Lgrid_henry)
            Main.user_cmp(int(Vref_pu_ID)).set_parameters(V_1=test_cases['Vref_PSCAD_droop (pu)'][i])
            Main.user_cmp(int(p_ppc_ref_pu_ID)).set_parameters(P_1=test_cases['Initial P PSCAD (pu)'][i])
            Main.user_cmp(int(Vsmib_pu_ID)).set_parameters(Y1=test_cases['Vsmib_PSCAD (pu)'][i])
            Main.user_cmp(int(PPC_ID)).set_parameters(Pctrl=1, Qctrl=1)
            # Main Program:
            Main.user_cmp(p_ppc_ref_pu_ID).set_parameters(N=N,X1=refs[0][0],X2=refs[1][0],X3=refs[2][0],X4=refs[3][0],X5=refs[4][0],X6=refs[5][0],X7=refs[6][0],X8=refs[7][0],X9=refs[8][0],X10=refs[9][0],
                                                          Y1=refs[0][1],Y2=refs[1][1],Y3=refs[2][1],Y4=refs[3][1],Y5=refs[4][1],Y6=refs[5][1],Y7=refs[6][1],Y8=refs[7][1],Y9=refs[8][1],Y10=refs[9][1])
            if 'interpolate' in test_cases['Ref 7'][i].lower():
                Main.user_cmp(p_ppc_ref_pu_ID).set_parameters(Mode=0)
            else:
                Main.user_cmp(p_ppc_ref_pu_ID).set_parameters(Mode=1)
            N2, ref2s, time_duration2 = get_reference(test_cases['Ref 8'][i])
            Main.user_cmp(Q_ppc_ref_pu_ID).set_parameters(N=N2,X1=ref2s[0][0],X2=ref2s[1][0],X3=ref2s[2][0],X4=ref2s[3][0],X5=ref2s[4][0],X6=ref2s[5][0],X7=ref2s[6][0],X8=ref2s[7][0],X9=ref2s[8][0],X10=ref2s[9][0],
                                                          Y1=ref2s[0][1],Y2=ref2s[1][1],Y3=ref2s[2][1],Y4=ref2s[3][1],Y5=ref2s[4][1],Y6=ref2s[5][1],Y7=ref2s[6][1],Y8=ref2s[7][1],Y9=ref2s[8][1],Y10=ref2s[9][1])
            if 'interpolate' in test_cases['Ref 9'][i].lower():
                Main.user_cmp(Q_ppc_ref_pu_ID).set_parameters(Mode=0)
            else:
                Main.user_cmp(Q_ppc_ref_pu_ID).set_parameters(Mode=1)

            time_duration = max(time_duration,time_duration2)
            project.set_parameters(time_duration=time_duration)
            project.create_layer('Disable')
            project.set_layer('Disable', 'disabled')
            for j, chan in enumerate(Chans):
                if type(Enables[j]) == float or type(Enables[j]) == int:
                    if int(Enables[j]) == 0:
                        Main.user_cmp(IDs[j]).add_to_layer('Disable')
                    elif (int(Enables[j]) == 1 and New_Symbols[j] != '' and New_Symbols[j] != Symbols[j]):
                        Main.user_cmp(IDs[j]).set_parameters(UseSignalName=0, Name=New_Symbols[j])
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
                overlay_dir = os.path.join(overlay, '%s' % output_filename[0:-6])
                uc.dirCreateClean(overlay_dir, [".pdf"])
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
def RL_calculation(VPCCnom,sc_pcc_pu_Sbase,xr):
    # sc_pcc_pu_Sbase = (scr * wfbase_MW) / 100.0  # 100MVA System base
    Rgrid_pu = math.sqrt(((1.0 / sc_pcc_pu_Sbase) ** 2) / (xr ** 2 + 1.0))
    Xgrid_pu = Rgrid_pu * xr
    Rgrid_ohm = Rgrid_pu * (VPCCnom * VPCCnom / 100)
    Xgrid_ohm = Xgrid_pu * (VPCCnom * VPCCnom / 100)
    Lgrid_henry = Xgrid_ohm / (2 * 50 * math.pi)
    return Rgrid_ohm,Lgrid_henry
# =======================================================
def find_scpu_xr(VPCCnom,Rgrid_ohm,Lgrid_henry):
    Xgrid_ohm = Lgrid_henry * (2 * 50 * math.pi)
    Xgrid_pu = Xgrid_ohm/(VPCCnom * VPCCnom / 100)
    Rgrid_pu = Rgrid_ohm/(VPCCnom * VPCCnom / 100)
    xr = Xgrid_pu/Rgrid_pu
    sc_pcc_pu_Sbase = 1/(math.sqrt(Rgrid_pu ** 2 + Xgrid_pu ** 2))
    return sc_pcc_pu_Sbase,xr
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
    xl_sheet = xl_workbook.sheet_by_name('PPC_PSSE_VWRE')
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
        # ---------------------------------------
        if sheet_name == 'WTG_PSCAD_Parfile':
            data = read_excel_sheet(excel_file, sheet_name)
            Project_info[sheet_name] = data
        # ---------------------------------------
        if sheet_name == 'PPC_PSCAD_Parfile':
            data = read_excel_sheet(excel_file, sheet_name)
            Project_info[sheet_name] = data
    return Project_info
# =======================================================
def create_WTG_PSCAD_Parfile(Project_info,PSCAD_version, WTG_PSCAD_Parfile):
    Parfile = open(WTG_PSCAD_Parfile, "w")
    for i in range(len(Project_info['WTG_PSCAD_Parfile']['Enable'])):
        if Project_info['WTG_PSCAD_Parfile']['Enable'][i] == 1:
            if type(Project_info['WTG_PSCAD_Parfile'][PSCAD_version][i]) == str:
                Project_info['WTG_PSCAD_Parfile'][PSCAD_version][i] = float(Project_info['WTG_PSCAD_Parfile'][PSCAD_version][i])
            Parfile.write("%g\t" %Project_info['WTG_PSCAD_Parfile'][PSCAD_version][i])
            Parfile.write("%s\t" %Project_info['WTG_PSCAD_Parfile']['Name'][i])
            Parfile.write('%d' % Project_info['WTG_PSCAD_Parfile']['Index'][i])
            Parfile.write("\n")
    Parfile.close()
# =======================================================
def create_PPC_PSCAD_Parfile(Project_info,PSCAD_version, PPC_PSCAD_Parfile):
    Parfile = open(PPC_PSCAD_Parfile, "w")
    Parfile.write('!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! \n')
    Parfile.write('! Parameters always user defined from PSCAD tool interface BEGIN\n')
    Parfile.write('! Site_P_Nominal_kW\n')
    Parfile.write('! Site_V_Nominal_KVolt\n')
    Parfile.write('! P_FRB_TSO_kW\n')
    Parfile.write('! UseSetpointLimits (1/0)\n')
    Parfile.write('! MaxSetpointSumCAP_kVAR\n')
    Parfile.write('! MaxSetpointSumIND_kVAR\n')
    Parfile.write('! QBase_kVAR\n')
    Parfile.write('! UseRefQLimits  (1/0)\n')
    Parfile.write('! RefQLimCAP_kVAR=Site_P_Nominal_kW*(TAN(ACOS(0.95)))\n')
    Parfile.write('! P_ControlMode(1=Active_power_control,2=Frequency_control_type_1,3=Frequency_control_type_2)\n')
    Parfile.write('! Q_ControlMode(1=Reactive_power_control,2=Power_factor_control,3=Voltage_control,4=Voltage_droop_control,5=VoltageQ_Slope_Control)\n')
    Parfile.write('! Enable_P_FRB_TSO\n')
    Parfile.write('! Parameters always user defined from PSCAD tool interface END\n')
    Parfile.write('! The above values set from PSCAD tool interface will overwrite the below corresponding default values \n')
    Parfile.write('!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n')
    Parfile.write('\n')
    Parfile.write('!PPC PM Start ===========================================================\n')
    for i in range(len(Project_info['PPC_PSCAD_Parfile']['Index'])):
        if type(Project_info['PPC_PSCAD_Parfile']['Index'][i]) == str:
            if 'PPC PM Start' in Project_info['PPC_PSCAD_Parfile']['Index'][i]:
                PM_Start = i
            elif 'PPC PM End' in Project_info['PPC_PSCAD_Parfile']['Index'][i]:
                PM_End = i
            elif 'PPC P Start' in Project_info['PPC_PSCAD_Parfile']['Index'][i]:
                P_Start = i
            elif 'PPC P End' in Project_info['PPC_PSCAD_Parfile']['Index'][i]:
                P_End = i
            elif 'PPC Q Start' in Project_info['PPC_PSCAD_Parfile']['Index'][i]:
                Q_Start = i
            elif 'PPC Q End' in Project_info['PPC_PSCAD_Parfile']['Index'][i]:
                Q_End = i
    for i in range(PM_Start,PM_End):
        if Project_info['PPC_PSCAD_Parfile']['Enable'][i] == 1:
            if type(Project_info['PPC_PSCAD_Parfile'][PSCAD_version][i]) == str:
                Project_info['PPC_PSCAD_Parfile'][PSCAD_version][i] = float(Project_info['PPC_PSCAD_Parfile'][PSCAD_version][i])
            Parfile.write("%g\t" %float(Project_info['PPC_PSCAD_Parfile'][PSCAD_version][i]))
            Parfile.write("%s\t" %Project_info['PPC_PSCAD_Parfile']['Description'][i])
            Parfile.write('%d'% Project_info['PPC_PSCAD_Parfile']['Index'][i])
            Parfile.write("\n")
    Parfile.write('!PPC PM End ===========================================================\n')
    Parfile.write('!PPC P Start ===========================================================\n')
    for i in range(P_Start,P_End):
        if Project_info['PPC_PSCAD_Parfile']['Enable'][i] == 1:
            if type(Project_info['PPC_PSCAD_Parfile'][PSCAD_version][i]) == str:
                Project_info['PPC_PSCAD_Parfile'][PSCAD_version][i] = float(Project_info['PPC_PSCAD_Parfile'][PSCAD_version][i])
            Parfile.write("%s\t" %Project_info['PPC_PSCAD_Parfile']['Description'][i])
            Parfile.write('%d'% Project_info['PPC_PSCAD_Parfile']['Index'][i])
            Parfile.write("\n")
    Parfile.write('!PPC P End ===========================================================\n')
    Parfile.write('!PPC Q Start ===========================================================\n')
    for i in range(Q_Start,Q_End):
        if Project_info['PPC_PSCAD_Parfile']['Enable'][i] == 1:
            if type(Project_info['PPC_PSCAD_Parfile'][PSCAD_version][i]) == str:
                Project_info['PPC_PSCAD_Parfile'][PSCAD_version][i] = float(Project_info['PPC_PSCAD_Parfile'][PSCAD_version][i])
            Parfile.write("%s\t" %Project_info['PPC_PSCAD_Parfile']['Description'][i])
            Parfile.write('%d'% Project_info['PPC_PSCAD_Parfile']['Index'][i])
            Parfile.write("\n")
    Parfile.close()
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
    short_title = os.path.spliEnables(filename)[0]
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
def get_reference(list):
    refs = []
    for ref in (re.split('[;]', list)):
        ref = ref.strip('()[]').split(',')
        float_ref = ref
        for i in range(len(ref)):
            try:
                float_ref[i] = float(ref[i])
            except:
                float_ref[i] = (ref[i])
        refs.append(float_ref)
    N = int(len(refs))
    while len(refs) <10:
        refs.append(refs[-1])
    time_duration = 0
    for ref in refs:
        if ref[0] > time_duration:
            time_duration = ref[0]
    return N, refs, time_duration
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
    Enables = chandata['Enable/Disable']
    Symbols = chandata['Symbol']
    New_Symbols = chandata['New_Symbol']
    return IDs,Chans,Enables,Symbols,New_Symbols
# ========================================================
def plot_all_in_dir(dir_to_plot,Test,Project_info,wfbase_MW):
    dirCreateClean(dir_to_plot, [".pdf"])
    filenames = glob.glob(os.path.join(dir_to_plot,'*.inf'))
    plot_Vdroop, plot_Qdroop = PSCAD_Droop_parameters(Project_info, wfbase_MW)
    for filename in filenames:
        if 'FCT' in Test:
            plot = pl.Plotting()
            plot.read_data(filename)
            plot.subplot_spec(0, (0, 'F_PCC_MEAS_HZ'), title='Frequency', ylabel='Frequency (Hz)', scale=1.0, offset=0.0)
            plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(1, (0, 'V_WTG01_MEAS_PU'))
            plot.subplot_spec(1, (0, 'V_WTG02_MEAS_PU'))
            plot.subplot_spec(1, (0, 'V_PCC_REF_PU'), scale=1.0)
            plot.subplot_spec(2, (0, 'P_WTG01_MEAS_MW'), title='Active Power at WTG', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(2, (0, 'P_WTG02_MEAS_MW'), scale=1.0)
            plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(3, (0, 'P_PCC_REF_MW'),scale=1, offset=0.0)
            plot.subplot_spec(4, (0, 'Q_WTG01_MEAS_MVAR'), title='Reactive Power at WTG', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
            plot.subplot_spec(4, (0, 'Q_WTG02_MEAS_MVAR'), scale=1.0)
            plot.subplot_spec(4, (0, 'Q_MSU1_MEAS_MVAR'), scale=1.0)
            plot.subplot_spec(4, (0, 'Q_MSU2_MEAS_MVAR'), scale=1.0)
            plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0, plot_Qdroop=plot_Qdroop)
        else:
            plot = pl.Plotting()
            plot.read_data(filename)
            plot.subplot_spec(0, (0, 'V_WTG01_MEAS_PU'), title='WTG Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(0, (0, 'V_WTG02_MEAS_PU'))
            plot.subplot_spec(1, (0, 'V_PCC_MEAS_PU'), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(1, (0, 'V_PCC_REF_PU'), scale=1.0)
            plot.subplot_spec(2, (0, 'P_WTG01_MEAS_MW'), title='Active Power at WTG', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(2, (0, 'P_WTG02_MEAS_MW'), scale=1.0)
            plot.subplot_spec(3, (0, 'P_PCC_MEAS_MW'), title='Active Power at POC', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(3, (0, 'P_PCC_REF_MW'),scale=1, offset=0.0)
            plot.subplot_spec(4, (0, 'Q_WTG01_MEAS_MVAR'), title='Reactive Power at WTG', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
            plot.subplot_spec(4, (0, 'Q_WTG02_MEAS_MVAR'), scale=1.0)
            plot.subplot_spec(4, (0, 'Q_MSU1_MEAS_MVAR'), scale=1.0)
            plot.subplot_spec(4, (0, 'Q_MSU2_MEAS_MVAR'), scale=1.0)
            plot.subplot_spec(5, (0, 'Q_PCC_MEAS_MVAR'), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0, plot_Qdroop=plot_Qdroop)
        plot.plot(figname=os.path.splitext(filename)[0], show=0)
    merge_pdf(dir_to_plot)
# =======================================================
def PSCAD_Droop_parameters(Project_info,wfbase_MW):
    if 'VWRE_J' in Project_info:
        POC_V_ref_Channel_id = 'V_PCC_REF_PU'
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
        POC_V_ref_Channel_id = 'V_PCC_REF_PU'
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