import glob, os, sys, math, csv, time, logging, traceback, os.path, shutil
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
cwd = os.getcwd()
plotting_dir = r"""C:\Users\qdanu\OneDrive - Vestas Wind Systems A S\Projects\Lib\Plotting_tools"""
sys.path.append(plotting_dir)
import PSSE_PSCAD_plot as PSSE_PSCAD_plt
User_code = r"""C:\Users\qdanu\OneDrive - Vestas Wind Systems A S\Projects\Lib\User_code"""
sys.path.append(User_code)
import usercode27 as uc

xoffset = 0
step_boundary = [0,0]
dirs = next(os.walk('.'))[1]
# ========================================================
def dirCreateClean(path,fileTypes):
    global cwd
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
def plot2out(dir_to_plot,xlims,ylimit):
    channels_PSSE = [['F_DEV_PCC_MEAS_PU'],['V_PCC_MEAS_PU'], ['P_PCC_MEAS_MW'], ['Q_PCC_MEAS_MVAR']]
    channels_PSCAD = [['Sys_Freq'],['V_PCC_RMS_meas'], ['P_PCC_meas'], ['Q_PCC_meas']]
    channels_PSSE_adj = [[[50, 50]],[[1, 0]],  [[1, 0]], [[1, 0]]]
    channels_PSCAD_adj = [[[1, 0]],[[1, 0]], [[1, 0]], [[1, 0]]]
    ylabel = ['Frequency(Hz)', 'POC_Voltage (pu)','P_POC (MW)', 'QPOC (MVAr)']
    title = ['Frequency (Hz)','POC_Voltage (pu)','P_POC (MW)', 'QPOC (MVAr)']
    PSSE_PSCAD_plt.plot_all_out_files(dir_to_plot, [channels_PSSE, channels_PSCAD],[channels_PSSE_adj, channels_PSCAD_adj], title=title, xlims=xlims, ylims=ylimit, ylabel=ylabel,xoffset = xoffset,step_boundary=step_boundary)
# ========================================================
dirs = next(os.walk('.'))[1]
for dir in dirs:
    output_files_dir = os.path.join(cwd, dir)
    dirCreateClean(output_files_dir,['*.pdf'])
    if 'Case_00' in dir:
        xlims = [5, 20]
        ylimit = ''
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_01' in dir:
        xlims = [5,140]
        ylimit = ""
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_02' in dir:
        xlims = [5, 15]
        ylimit = ''
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_03' in dir:
        xlims = [5, 15]
        ylimit = ''
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_04' in dir:
        xlims = [5, 15]
        ylimit = ''
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_05' in dir:
        xlims = [9.5, 11]
        ylimit = ''
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_06' in dir:
        xlims = [10, 150]
        ylimit = ''
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_07' in dir:
        xlims = [10, 150]
        ylimit = ''
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_08' in dir:
        xlims = [10, 700]
        ylimit = ''
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_09' in dir:
        xlims = [10, 700]
        ylimit = [[46,56],[0.9,1.1],[200,240],[-15,25]]
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_10' in dir:
        xlims = [5, 50]
        ylimit = [[49.5,50.5],[0.95,1.05],[200,240],[-70,70]]
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_11' in dir:
        xlims = [5, 50]
        ylimit = [[49.5,50.5],[0.9,1.1],[200,240],[-100,80]]
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_12' in dir:
        xlims = [5, 50]
        ylimit = [[49.5,50.5],[0.88,1.02],[200,240],[40,80]]
        plot2out(output_files_dir,xlims,ylimit)
    elif 'Case_13' in dir:
        xlims = [5, 50]
        ylimit = [[49.5,50.5],[0.95,1.15],[200,240],[-100,-60]]
        plot2out(output_files_dir,xlims,ylimit)
    uc.merge_pdf(output_files_dir)
