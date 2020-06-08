import time
import os,sys
import shutil
import glob
from PyPDF2 import PdfFileMerger
timestr = time.strftime("%Y%m%d_%H%M%S")
cwd = os.getcwd()
output_files_dir = cwd
copy_dir = os.path.dirname(cwd)
plotting_dir = r'C:\Users\nguye\Dropbox\Vestas\Lib\PSCAD PSSE Plotting'
sys.path.append(plotting_dir)
import PSSE_PSCAD_plot as PSSE_PSCAD_plt
xlims=[7, 20]
ylimit = ""
t_start_stop = [10,15]
dirs = next(os.walk('.'))[1]
Pmax = 83.6
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
    channels_PSSE = [['POC_VOLTAGE'], ['P_POC'], ['Q_POC']]
    channels_PSCAD = [['PCC_Vmeas'], ['PCC_Pmeas'], ['PCC_Qmeas']]
    channels_PSSE_adj = [[[1, 0]],  [[1, 0]], [[1, 0]]] # pu * 100/37 (37 is number of inverter in the feeder 1)
    channels_PQM_adj = [[[1, 0]], [[1, 0]], [[1, 0]]]
    ylabel = ['POC_Voltage (pu)','P_POC (MW)', 'QPOC (MVAr)']
    title = ['POC_Voltage (pu)','P_POC (MW)', 'QPOC (MVAr)']
    PSSE_PSCAD_plt.plot_all_out_files(dir_to_plot, [channels_PSSE, channels_PSCAD],[channels_PSSE_adj, channels_PQM_adj], title=title, xlims=xlims, ylims=ylimit, ylabel=ylabel,xoffset = 0,step_boundary=[7.9,8.2])
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

dirs = next(os.walk('.'))[1]
for dir in dirs:
    output_files_dir = os.path.join(cwd, dir)
    dirCreateClean(output_files_dir,['*.pdf'])
    plot2out(output_files_dir,xlims,ylimit)
    merge_pdf(output_files_dir)


