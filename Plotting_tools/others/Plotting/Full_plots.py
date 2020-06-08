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
    # dirCreateClean(dir_to_plot, ["*.pdf"])
    #os.chdir(dir_to_plot)
    filenames = glob.glob(os.path.join(dir_to_plot, '*.xlsx'))
    for filename in filenames:
        if 'VCT' in os.path.basename(filename):
            plot = pl.Plotting()
            xlimit = plot.read_data2(os.path.join(dir_to_plot, filename))
            xtype = 'hh:mm'
            if 'poc' in filename.lower():
                loc = 'POC'
                Vscale = 1 / 66000
                ylimitP = [0,70]
                ylimitQ = [-40, 10]
                #plot_setpoint = ['SP_From_PPC_ALL_04062019.xlsx', 'PPC_Processing.U_Ctrl_Loop_SP.Value', 1 ,0.0,  0]
                #plot_setpointP = ['SP_From_PPC_ALL_04062019.xlsx', 'PPC_Processing.P_Lim_Supply_UTI_target',0.001 ,0, 0]
                plot_setpoint = ''
                plot_setpointP = ''
            elif '33kv' in filename.lower():
                loc = '33KV'
                Vscale = 1 / 33000
                ylimitP = [0, 50]
                plot_setpoint = ''
            else:
                loc = 'Inverter terminal'
                Vscale = 1 / 575
                ylimitP = [0, 2]
                plot_setpoint = ''
            plot.subplot_spec(0, (0, 6), title='%s voltage' %loc, ylabel='Voltage (pu)', scale= Vscale, offset=0.0, xlimit =xlimit,xtype = xtype, plot_setpoint=plot_setpoint)
            # plot.subplot_spec(0, (0, 6),scale= Vscale)
            # plot.subplot_spec(0, (0, 7),scale= Vscale)
            plot.legends[0] = ['POC_Votlage', 'POC_Votlage_setpoint']
            plot.subplot_spec(1, (0, '[W] Active Power - Fundamental Total (Auto) Average'), title='Active Power at %s' %loc, ylabel='Active Power(MW)', scale=0.000001, offset=0.0,ylimit = ylimitP, plot_setpoint = plot_setpointP)
            plot.legends[1] = ['P_%s_Act' %loc]
            plot.subplot_spec(2, (0, '[VAr] Reactive Power - Fundamental Total (Auto) Average'), title='Reactive Power at %s' %loc, ylabel='Reactive power(MVar)', ylimit = ylimitQ, scale=0.000001, offset=0.0)
            plot.legends[2] = ['Q_%s_Act' %loc]
            plot.subplot_spec(3, (0, 0), title='Power Factor', ylabel='Power Factor', scale=1.0, offset=0.0, plot_PF=['Active Power', 'Reactive Power', 'none', 'none'])#)
            plot.legends[3] = ['Power Factor at %s' %loc]
            # plot.subplot_spec(3, (0, 'Frequency'), title='Systems Frequency', ylabel='Systems Frequency (Hz)', scale=1.0, offset=0.0)
            # plot.legends[3] = ['Systems Frequency']
            plot.plot(figname=os.path.splitext(filename)[0]+ '_Full', show=0)

        elif 'PFCT' in os.path.basename(filename):
            plot = pl.Plotting()
            xlimit = plot.read_data2(os.path.join(dir_to_plot, filename))
            xtype = 'hh:mm'
            if 'poc' in filename.lower():
                loc = 'POC'
                Vscale = 1 / 66000
                ylimitP = [0, 70]
                ylimitQ = [-40, 10]
                if 'cap' in filename.lower():
                    plot_setpoint = ['SP_From_PPC_PFCT_06062019.xlsx', 'PPC!PPC_Processing.IN_PPC_IN.Angle_Ctrl_Loop_SP.Value', 1,0, 0]
                    ylimit = [0.999,1.003]
                elif 'ind' in filename.lower():
                    plot_setpoint = ['SP_From_PPC_PFCT_05062019.xlsx','PPC!PPC_Processing.IN_PPC_IN.Angle_Ctrl_Loop_SP.Value', 1,0, 0]
                    ylimit = ''
                # plot_setpoint = ''
            elif '33kv' in filename.lower():
                loc = '33KV'
                Vscale = 1 / 33000
                plot_setpoint = ''
            else:
                loc = 'Inverter terminal'
                Vscale = 1 / 575
                plot_setpoint = ''
            plot.subplot_spec(0, (0, '[V] RMS - Fundamental V31 (Auto) Average'), title='%s voltage' % loc,ylabel='Voltage (pu)', scale=Vscale, offset=0.0, xlimit=xlimit, xtype=xtype,)
            plot.legends[0] = ['%s_VOLTAGE_Act' % loc]
            plot.subplot_spec(1, (0, '[W] Active Power - Fundamental Total (Auto) Average'),title='Active Power at %s' % loc, ylabel='Active Power(MW)',ylimit = ylimitP, scale=0.000001, offset=0.0)
            plot.legends[1] = ['P_%s_Act' % loc]
            plot.subplot_spec(2, (0, '[VAr] Reactive Power - Fundamental Total (Auto) Average'),title='Reactive Power at %s' % loc, ylabel='Reactive power(MVar)', ylimit = ylimitQ, scale=0.000001,offset=0.0)
            plot.legends[2] = ['Q_%s_Act' % loc]
            plot.subplot_spec(3, (0, 0), title='Power Factor', ylabel='Power Factor', scale=1.0, offset=0.0, plot_PF=['Active Power', 'Reactive Power', 'none', 'none'],plot_setpoint = plot_setpoint, ylimit = ylimit)#)
            plot.legends[3] = ['Power Factor at %s' %loc,'Power Factor setpoint']
            # plot.subplot_spec(3, (0, 'Frequency'), title='Systems Frequency', ylabel='Systems Frequency (Hz)',scale=1.0, offset=0.0)
            # plot.legends[3] = ['Systems Frequency']
            plot.plot(figname=os.path.splitext(filename)[0]+ '_Full', show=0)

        elif 'TCT' in os.path.basename(filename):
            plot = pl.Plotting()
            xlimit = plot.read_data2(os.path.join(dir_to_plot, filename))
            xtype = 'hh:mm'
            if 'poc' in filename.lower():
                loc = 'POC'
                Vscale = 1 / 66000
                ylimitP = [0, 70]
                ylimitQ = [-40, 10]
                if 'trans1' in filename.lower():
                    plot_setpoint = ['SP_From_PPC_ALL_04062019.xlsx', 'PPC_Processing.P_Lim_Supply_UTI_target', 0.001,0, 0]
                elif 'trans2' in filename.lower():
                    plot_setpoint = ['SP_From_PPC_ALL_06062019.xlsx', 'PPC_Processing.P_Lim_Supply_UTI_target', 0.001, 0, 0]
            elif '33kv' in filename.lower():
                loc = '33KV'
                Vscale = 1 / 33000
                ylimitP = [0, 40]
                ylimitQ = [-20, 0]
                plot_setpoint = ''
            else:
                loc = 'Inverter terminal'
                Vscale = 1 / 575
                plot_setpoint = ''
            plot.subplot_spec(0, (0, '[V] RMS - Fundamental V31 (Auto) Average'), title='%s voltage' % loc,ylabel='Voltage (pu)', scale=Vscale, offset=0.0, xlimit=xlimit, xtype=xtype)
            plot.legends[0] = ['%s_VOLTAGE_Act' % loc]
            plot.subplot_spec(1, (0, '[W] Active Power - Fundamental Total (Auto) Average'),title='Active Power at %s' % loc, ylabel='Active Power(MW)',ylimit = ylimitP, scale=0.000001,
                              offset=0.0, plot_setpoint= plot_setpoint)
            plot.legends[1] = ['P_%s_Act' % loc, 'P_%s_Setpoint' % loc]
            plot.subplot_spec(2, (0, '[VAr] Reactive Power - Fundamental Total (Auto) Average'),title='Reactive Power at %s' % loc, ylabel='Reactive power(MVar)', ylimit = ylimitQ, scale=0.000001,offset=0.0)
            plot.legends[2] = ['Q_%s_Act' % loc]
            plot.subplot_spec(3, (0, 0), title='Power Factor', ylabel='Power Factor', scale=1.0, offset=0.0, plot_PF=['Active Power', 'Reactive Power', 'none', 'none'])#)
            plot.legends[3] = ['Power Factor at %s' %loc]
            # plot.subplot_spec(3, (0, 'Frequency'), title='Systems Frequency', ylabel='Systems Frequency (Hz)',scale=1.0, offset=0.0)
            # plot.legends[3] = ['Systems Frequency']
            plot.plot(figname=os.path.splitext(filename)[0]+ '_Full', show=0)

        elif 'APD' in os.path.basename(filename):
            plot = pl.Plotting()
            xlimit = plot.read_data2(os.path.join(dir_to_plot, filename))
            xtype = 'hh:mm'
            if 'poc' in filename.lower():
                loc = 'POC'
                Vscale = 1 / 66000
                ylimitP = [0, 70]
                ylimitQ = [-40, 10]
                plot_setpoint = ['SP_From_PPC_ALL_06062019.xlsx', 'PPC_Processing.P_Lim_Supply_UTI_target', 0.001 ,0,  0]
            elif '33kv' in filename.lower():
                loc = '33KV'
                Vscale = 1 / 33000
                plot_setpoint = ''
            else:
                loc = 'Inverter terminal'
                Vscale = 1 / 575
                plot_setpoint = ''
            plot.subplot_spec(0, (0, '[V] RMS - Fundamental V31 (Auto) Average'), title='%s voltage' % loc,ylabel='Voltage (pu)', scale=Vscale, offset=0.0, xlimit=xlimit, xtype=xtype)
            plot.legends[0] = ['%s_VOLTAGE_Act' % loc]
            plot.subplot_spec(1, (0, '[W] Active Power - Fundamental Total (Auto) Average'),title='Active Power at %s' % loc, ylabel='Active Power(MW)',ylimit = ylimitP, scale=0.000001, offset=0.0, plot_setpoint= plot_setpoint)
            plot.legends[1] = ['P_%s_Act' % loc, 'P_%s_setpoint' %loc]
            plot.subplot_spec(2, (0, '[VAr] Reactive Power - Fundamental Total (Auto) Average'),title='Reactive Power at %s' % loc, ylabel='Reactive power(MVar)', ylimit = ylimitQ, scale=0.000001,offset=0.0)
            plot.legends[2] = ['Q_%s_Act' % loc]
            plot.subplot_spec(3, (0, 0), title='Power Factor', ylabel='Power Factor', scale=1.0, offset=0.0, plot_PF=['Active Power', 'Reactive Power', 'none', 'none'])#)
            plot.legends[3] = ['Power Factor at %s' %loc]
            # plot.subplot_spec(3, (0, 'Frequency'), title='Systems Frequency', ylabel='Systems Frequency (Hz)',scale=1.0, offset=0.0)
            # plot.legends[3] = ['Systems Frequency']
            plot.plot(figname=os.path.splitext(filename)[0]+ '_Full', show=0)

        elif 'RCT' in os.path.basename(filename):
            plot = pl.Plotting()
            xlimit = plot.read_data2(os.path.join(dir_to_plot, filename))
            xtype = 'hh:mm'
            if 'poc' in filename.lower():
                loc = 'POC'
                Vscale = 1 / 66000
                ylimitP = [0, 70]
                ylimitQ = [-40, 10]
                if 'cap' in filename.lower():
                    plot_setpoint = ['SP_From_PPC_ALL_06062019.xlsx', 'PPC_Processing.Q_Ctrl_Loop_SP.Value', 0.001 , 0, 0]
                    plot_setpointP = ['SP_From_PPC_ALL_06062019.xlsx', 'PPC_Processing.P_Lim_Supply_UTI_target', 0.001 , 0, 0]
                else:
                    plot_setpoint = ['SP_From_PPC_ALL_02062019.xlsx', 'PPC_Processing.Q_Ctrl_Loop_SP.Value', 0.001 , 0, 0]
                    plot_setpointP = ['SP_From_PPC_ALL_02062019.xlsx', 'PPC_Processing.P_Lim_Supply_UTI_target', 0.001 , 0, 0]

            elif '33kv' in filename.lower():
                loc = '33KV'
                Vscale = 1 / 33000
                plot_setpoint = ''
            else:
                loc = 'Inverter terminal'
                Vscale = 1 / 575
                plot_setpoint = ''
            plot.subplot_spec(0, (0, '[V] RMS - Fundamental V31 (Auto) Average'), title='%s voltage' % loc,ylabel='Voltage (pu)', scale=Vscale, offset=0.0, xlimit=xlimit, xtype=xtype)
            plot.legends[0] = ['%s_VOLTAGE_Act' % loc]
            plot.subplot_spec(1, (0, '[W] Active Power - Fundamental Total (Auto) Average'),title='Active Power at %s' % loc, ylabel='Active Power(MW)',ylimit = ylimitP, scale=0.000001, offset=0.0, plot_setpoint= plot_setpointP)
            plot.legends[1] = ['P_%s_Act' % loc]
            plot.subplot_spec(2, (0, '[VAr] Reactive Power - Fundamental Total (Auto) Average'),title='Reactive Power at %s' % loc, ylabel='Reactive power(MVar)', ylimit = ylimitQ, scale=0.000001,offset=0.0, plot_setpoint= plot_setpoint)
            plot.legends[2] = ['Q_%s_Act' % loc,'Q_%s_Setpoint' % loc]
            plot.subplot_spec(3, (0, 0), title='Power Factor', ylabel='Power Factor', scale=1.0, offset=0.0, plot_PF=['Active Power', 'Reactive Power', 'none', 'none'])#)
            plot.legends[3] = ['Power Factor at %s' %loc]
            # plot.subplot_spec(3, (0, 'Frequency'), title='Systems Frequency', ylabel='Systems Frequency (Hz)',scale=1.0, offset=0.0)
            # plot.legends[3] = ['Systems Frequency']
            plot.plot(figname=os.path.splitext(filename)[0]+ '_Full', show=0)

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
        shutil.copy(new_dst_file,copy_dir)

dirs = next(os.walk('.'))[1]
dirs = ['J:\Grid Connection\Numurkah Solar Farm\Hold Points\HP2\HP2 results\TCT']
for dir in dirs:
    output_files_dir = os.path.join(cwd, dir)
    plot_standard_files(output_files_dir)
    #merge_pdf(output_files_dir)
# output_files_dir = cwd
# plot_standard_files(output_files_dir)
# merge_pdf(output_files_dir)
