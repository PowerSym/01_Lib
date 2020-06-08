import time
import os,sys
import shutil
import glob
from PyPDF2 import PdfFileMerger

timestr = time.strftime("%Y%m%d_%H%M%S")
cwd = os.getcwd()
output_files_dir = cwd
copy_dir = os.path.dirname(cwd)
plotting_dir = cwd
sys.path.append(plotting_dir)

delay = 1.5
xoffset = 46560
xoffset += delay

xlims = [46560, 48500]
ylimit = [[1.025,1.06],[19.5,20.5],[-5,10]]
ylimit = ""

step_boundary = [46560, 48499]
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
    import PSSE_PQM_plot as PSSE_PQM_plt
    channels_PSSE = [['POC_VOLTAGE'],  ['P_POC'], ['Q_POC']]
    channels_PQM = [['V12'], ['Active Power'], ['Reactive Power']]
    channels_PSSE_adj = [[[1, 0]],  [[1, 0]], [[1, 0]]]
    channels_PQM_adj = [[[0.000015151515, 0]], [[0.000001, 0]], [[0.000001, 0]]]
    ylabel = ['POC_Voltage (pu)','P_POC (MW)', 'QPOC (MVAr)']
    title = ['POC_Voltage (pu)','P_POC (MW)', 'QPOC (MVAr)']
    PSSE_PQM_plt.plot_all_out_files(dir_to_plot, [channels_PSSE, channels_PQM],[channels_PSSE_adj, channels_PQM_adj],['_PSSE','_PQM'], title=title, xlims=xlims, ylims=ylimit, ylabel=ylabel,xoffset = xoffset,
                                    step_boundary=step_boundary,)

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
dirCreateClean(output_files_dir,['*.pdf'])
plot2out(output_files_dir,xlims,ylimit)
merge_pdf(output_files_dir)
