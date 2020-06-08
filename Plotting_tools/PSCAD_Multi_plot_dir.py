import time
import os,sys
import shutil
import glob
from PyPDF2 import PdfFileMerger
cwd = os.getcwd()
output_files_dir = cwd
copy_dir = os.path.dirname(cwd)
plotting_dir = cwd
sys.path.append(plotting_dir)
import PSCAD_Multi_plot as PSCAD_Multi_plot
xlims = [4,6]
ylimit = ""
step_boundary = [4.9,5.1]
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
    dirCreateClean(dir_to_plot,['*.pdf'])
    import PSCAD_Multi_plot as PSCAD_Multi_plot
    channels_PRE = [['V_RMS_WTG1'], ['V_PCC_RMS_meas'], ['P_WTG2'], ['P_PCC_meas'], ['Q_WTG2'], ['Q_PCC_meas']]
    channels_POST = [['V_RMS_WTG1'], ['V_PCC_RMS_meas'], ['P_WTG2'], ['P_PCC_meas'], ['Q_WTG2'], ['Q_PCC_meas']]
    channels_PRE_adj = [[[1, 0]],  [[1, 0]], [[1, 0]], [[1, 0]],  [[1, 0]], [[1, 0]]] # pu * 100/37 (37 is number of inverter in the feeder 1)
    channels_POST_adj = [[[1, 0]], [[1, 0]], [[1, 0]], [[1, 0]],  [[1, 0]], [[1, 0]]]

    ylabel = ['WTG_Voltage', 'POC_Voltage', 'P_WTG','P_POC', 'Q_WTG', 'Q_POC']
    title = ['WTG_Voltage (pu)', 'POC_Voltage(pu)', 'P_WTG (MW)','P_POC (MW)', 'Q_WTG (MVAr)', 'Q_POC (MVAr)']
    PSCAD_Multi_plot.plot_all_out_files(dir_to_plot, [channels_PRE, channels_POST],[channels_PRE_adj, channels_POST_adj],  ['_old_pll','_new_pll'],title=title, xlims=xlims, ylims=ylimit, ylabel=ylabel, xoffset = 0, step_boundary = step_boundary)

# ========================================================
def merge_pdf(dir_to_merge):
    global cwd
    os.chdir(dir_to_merge)
    x = [a for a in os.listdir(dir_to_merge) if a.endswith(".pdf")]
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
        new_dst_file = dir_to_merge+'\\'+ dir_to_merge.split(os.sep)[-1] +'_'+ file
        os.rename(dst_file, new_dst_file)
        #shutil.copy(new_dst_file,copy_dir)
    os.chdir(cwd)
    
output_files_dir = cwd
plot2out(output_files_dir,xlims,ylimit)
merge_pdf(output_files_dir)
