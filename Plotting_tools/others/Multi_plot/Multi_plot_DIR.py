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
import Multi_plot as Multi_plt
xlims= ""
ylimit = ""

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
    import Multi_plot as Multi_plt
    channels_PRE = [['UUT_VOLTAGE_1101', 4], [7], [2, 5], [8], [3, 6], [9]]
    channels_POST = [[1, 4], [7], [2, 5], [8], [3, 6], [9]]
    title = ['UUT_Voltage (pu)', 'POC Voltage (pu)', 'PELEC (p.u)','P_POC (MW)', 'QELEC (pu)', 'QPOC (MVAr)']
    ylabel = ['UUT_Voltage (pu)', 'POC Voltage (pu)', 'PELEC (p.u)','P_POC (MW)', 'QELEC (pu)', 'QPOC (MVAr)']
    Multi_plt.plot_all_out_files(dir_to_plot, [channels_PRE, channels_POST],['_pre','_post'], title=title, xlims=xlims, ylims=ylimit, ylabel=ylabel)

    #Outfile channels 0: {1: 'UUT_VOLTAGE_1101', 2: 'UUT_PELEC_1101', 3: 'UUT_QELEC_1101', 4: 'UUT_VOLTAGE_1102', 5: 'UUT_PELEC_1102', 6: 'UUT_QELEC_1102', 7: 'POC_VOLTAGE', 8: 'P_POC', 9: 'Q_POC', 10: 'SYS FREQ DEVIATION', 11: 'SMIB_VOLTAGE', 12: 'MVA_POC', 'time': 'Time(s)'}

# ========================================================
def merge_pdf(dir_to_merge):
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
        new_dst_file = dir_to_merge+'\\'+ dir_to_merge.split(os.sep)[-1] +'_'+file
        os.rename(dst_file, new_dst_file)
        #shutil.copy(new_dst_file,copy_dir)
plot2out(output_files_dir,xlims,ylimit)
merge_pdf(output_files_dir)
