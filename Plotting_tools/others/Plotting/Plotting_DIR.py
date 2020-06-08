from __future__ import division
import time
import os,sys
import shutil
import glob
from PyPDF2 import PdfFileMerger
timestr = time.strftime("%Y%m%d_%H%M%S")
cwd = os.getcwd()
output_files_dir = cwd
copy_dir = cwd
dirs = next(os.walk('.'))[1]
import Plotting as pl
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
def plot_standard_files(dir_to_plot):
    dirCreateClean(dir_to_plot, ["*.pdf"])
    os.chdir(dir_to_plot)
    filenames = glob.glob(os.path.join(dir_to_plot, '*.*'))
    xtype = 'hh:mm'
    for filename in filenames:
        if filename.endswith('.xlsx') and 'pqm' in filename.lower():
            if 'poc' in filename.lower():
                loc = 'POC'
                Vscale = 1 / 66000
            elif '33kv' in filename.lower():
                loc = '33KV'
                Vscale = 1 / 33000
            else:
                loc = 'Inverter terminal'
                Vscale = 1 / 575
            plot = pl.Plotting()
            xlimit = plot.read_data2(os.path.join(dir_to_plot, filename))
            plot.subplot_spec(0, (0, '[V] RMS - Fundamental V12 (Auto) Average'), title='%s voltage' %loc, ylabel='Voltage (pu)', scale= Vscale, offset=0.0, xlimit =xlimit,xtype = xtype)
            plot.legends[0] = ['%s_VOLTAGE_Act' %loc]
            plot.subplot_spec(1, (0, '[W] Active Power - Fundamental Total (Auto) Average'), title='Active Power at %s' %loc, ylabel='Active Power(MW)', scale=0.000001, offset=0.0)
            plot.legends[1] = ['P_%s_Act' %loc]
            plot.subplot_spec(2, (0, '[VAr] Reactive Power - Fundamental Total (Auto) Average'), title='Reactive Power at %s' %loc, ylabel='Reactive power(MVar)', scale=0.000001, offset=0.0)
            plot.legends[2] = ['Q_%s_Act' %loc]
            # plot.subplot_spec(3, (0, 0), title='Power Factor', ylabel='Power Factor', scale=1.0, offset=0.0, plot_PF=['[W] Active Power', '[VAr] Reactive Power', 'none', 'none'])
            # plot.legends[3] = ['Power Factor at %s' %loc]
            plot.subplot_spec(3, (0, 2), title='Systems Frequency', ylabel='Systems Frequency (Hz)', scale=1.0, offset=0.0)
            plot.legends[3] = ['Systems Frequency']
            plot.plot(figname=os.path.splitext(filename)[0], show=0)
        elif filename.endswith('.out'):
            inf_files = glob.glob(os.path.join(dir_to_plot, '*.inf'))
            infx_files = glob.glob(os.path.join(dir_to_plot, '*.infx'))
            print filename
            print inf_files[0]
            if any(filename[0:-7] == s[0:-4] for s in inf_files) or any(filename[0:-7] == s[0:-4] for s in infx_files):
                pass
            else:
                plot = pl.Plotting()
                plot.read_data(os.path.join(dir_to_plot, filename))
                plot.subplot_spec(0, (0, 'UUT_VOLTAGE_1101'), title='Inverter voltage', ylabel='Voltage (pu)', offset=0.0)
                plot.subplot_spec(0, (0, 'UUT_VOLTAGE_1102'))
                plot.subplot_spec(1, (0, 'POC_VOLTAGE'), title='POC voltage', ylabel='Voltage (pu)', offset=0.0)
                plot.subplot_spec(2, (0, 'UUT_PELEC_1101'), title='Active Power at Inverter Terminal', scale = 100, ylabel='Voltage (pu)', offset=0.0)
                plot.subplot_spec(2, (0, 'UUT_PELEC_1102'),scale = 100)
                plot.subplot_spec(3, (0, 'P_POC'),title='Active Power at POC', ylabel='Active Power(MW) at POC', scale= 1, offset=0.0)
                plot.subplot_spec(4, (0, 'UUT_QELEC_1101'), title='Reactive Power at Inverter Terminal', ylabel='Reactive power(MVar)', scale=100,offset=0.0)
                plot.subplot_spec(4, (0, 'UUT_QELEC_1102'), scale =100)
                plot.subplot_spec(5, (0, 'Q_POC'),title='Reactive Power at POC', ylabel='Reactive power(MVar) at POC', scale= 1,offset=0.0)
                # plot.subplot_spec(5, (0, 'SYS FREQ DEVIATION'), title='SYSTEMS FREQUENCY', ylabel='Systems Frequency (Hz)', scale=50.0,offset=50.0)
                plot.plot(figname=os.path.splitext(filename)[0], show=0)
        elif filename.endswith('.inf') or filename.endswith('.infx'):
            plot = pl.Plotting()
            plot.read_data(filename)
            plot.subplot_spec(0, (0, 'UUT_Voltage_1101'), title='Inverter voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0, timeoffset=0)
            plot.subplot_spec(0, (0, 'UUT_Voltage_1102'))
            plot.subplot_spec(1, (0, 'POC_Voltage'), title='POC voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(2, (0, 'UUT_Pelec_1101'), title='Active Power at Inverter terminal',ylabel='Active Power (MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(2, (0, 'UUT_Pelec_1102'), scale=1.0)
            plot.subplot_spec(3, (0, 'P_POC'), title='Active Power', ylabel='Active Power(MW) at POC', scale=1.0, offset=0.0)
            plot.subplot_spec(4, (0, 'UUT_Qelec_1101'), title='Reactive power at Inverter terminal',ylabel='Reactive Power(MVAr)', scale=1.0, offset=0.0)
            plot.subplot_spec(4, (0, 'UUT_Qelec_1102'))
            plot.subplot_spec(5, (0, 'Q_POC'), title='Reactive Power', ylabel='Reactive power(MVar) at POC', scale=1.0,offset=0.0)
            plot.plot(figname=os.path.splitext(filename)[0], show=0)
        else:
            pass
    os.chdir(cwd)
# ========================================================
def merge_pdf(dir_to_merge):
    x = [a for a in os.listdir(dir_to_merge) if a.endswith(".pdf")]
    merger = PdfFileMerger()
    for pdf in x:
        merger.append(open(os.path.join(dir_to_merge,pdf), 'rb'))
    with open(os.path.join(dir_to_merge,"Plots.pdf"), "wb") as fout:
        merger.write(fout)
    for file in os.listdir(dir_to_merge):
        if file.endswith(".pdf") and (not file.endswith("Plots.pdf")):
            os.remove(os.path.join(dir_to_merge, file))
    for file in os.listdir(dir_to_merge):
        if (not file.endswith(".pdf")):
            continue    
        dst_file = dir_to_merge+'\\'+file
        new_dst_file = dir_to_merge+'\\'+ dir_to_merge.split(os.sep)[-1] +'_'+os.path.basename(file)
        os.rename(dst_file, new_dst_file)
        # shutil.copy(new_dst_file,copy_dir)

# dirs = next(os.walk('.'))[1]
# for dir in dirs:
#     output_files_dir = os.path.join(cwd, dir)
#     plot_standard_files(output_files_dir)
#     #merge_pdf(output_files_dir)
output_files_dir = cwd
plot_standard_files(output_files_dir)
merge_pdf(output_files_dir)
