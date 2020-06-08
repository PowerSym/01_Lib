from __future__ import division
import glob, os, sys, math, csv, time, logging, traceback, exceptions, os.path
import psse34
import psspy
import redirect
import shutil
import dyntools
from win32com import client
from matplotlib import pyplot as plt
import numpy as np
from PyPDF2 import PdfFileMerger

cwd = os.getcwd()
plotting_dir = os.path.join(os.path.dirname(cwd),'Lib\\Plotting_tools')
sys.path.append(plotting_dir)
import Plotting as pl
# ========================================================
def get_files_name(path, fileTypes):
    global cwd
    cases = []
    for type in fileTypes:
        for file in os.listdir(path):
            if (not file.endswith(type)):
                continue
            else:
                cases.append(file)
    return(cases)
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
def plot_all_in_dir(dir_to_plot,Test):
    dirCreateClean(dir_to_plot,["*.pdf"])
    os.chdir(dir_to_plot)
    filenames = glob.glob(os.path.join(dir_to_plot,'*.out'))
    for filename in filenames:
        plot = pl.Plotting()
        plot.read_data(os.path.join(dir_to_plot, filename))
        if 'APT' in Test:
            plot.subplot_spec(0, (0, 'POC F'), title='Frequency', ylabel='Frequency (Hz)', scale=50.0, offset=50.0)
            plot.subplot_spec(1, (0, 'POC V'), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(2, (0, 'POWR 50[ELAINE_WTG1 0.6500]1'), title='Active Power at WTG', ylabel='Active Power(MW)', scale=100.0, offset=0.0)
            plot.subplot_spec(2, (0, 'POWR 52[ELAINE_WTG3 0.6500]1'), scale=100.0)
            plot.subplot_spec(3, (0, 'P FLOW FROM POC'), title='Active Power at POC', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(3, (0, 'PSET IN PU'),scale=83.6)
            plot.subplot_spec(4, (0, 'VARS 50[ELAINE_WTG1 0.6500]1'), title='Reactive Power at WTG', ylabel='Reactive power(MVar)', scale=100.0,offset=0.0)
            plot.subplot_spec(4, (0, 'VARS 52[ELAINE_WTG3 0.6500]1'), scale=100.0)
            plot.subplot_spec(5, (0, 'Q FLOW FROM POC'), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
        elif 'RPT' in Test:
            plot.subplot_spec(0, (0, 'POC F'), title='Frequency', ylabel='Frequency (Hz)', scale=50.0, offset=50.0)
            plot.subplot_spec(1, (0, 'POC V'), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(2, (0, 'POWR 50[ELAINE_WTG1 0.6500]1'), title='Active Power at WTG', ylabel='Active Power(MW)', scale=100.0, offset=0.0)
            plot.subplot_spec(2, (0, 'POWR 52[ELAINE_WTG3 0.6500]1'), scale=100.0)
            plot.subplot_spec(3, (0, 'P FLOW FROM POC'), title='Active Power at POC', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(4, (0, 'VARS 50[ELAINE_WTG1 0.6500]1'), title='Reactive Power at WTG', ylabel='Reactive power(MVar)', scale=100.0,offset=0.0)
            plot.subplot_spec(4, (0, 'VARS 52[ELAINE_WTG3 0.6500]1'), scale=100.0)
            plot.subplot_spec(5, (0, 'Q FLOW FROM POC'), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
            plot.subplot_spec(5, (0, 'QSET IN PU'), scale=83.6)
            plot.subplot_spec(5, (0, 'Qref after limiter'), scale=83.6)
        elif 'VCT' in Test or 'VDT' in Test:
            # plot.subplot_spec(0, (0, 'POC F'), title='Frequency', ylabel='Frequency (Hz)', scale=50.0, offset=50.0)
            plot.subplot_spec(0, (0, 'VSET IN PU'), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(1, (0, 'POC V'), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0,rstime = [50,59,'',''])#, plot_Vdroop = [32, 10, 0.05, 33.022, 0.005]) # plot_Vdroop = POC_V_ref_Channel_id,Q_POC_Channel_id, QDROOP, SBASE, deadband
            # plot.subplot_spec(0, (0, 'VSET IN PU'))
            plot.subplot_spec(2, (0, 'POWR 50[ELAINE_WTG1 0.6500]1'), title='Active Power at WTG', ylabel='Active Power(MW)', scale=100.0, offset=0.0)
            plot.subplot_spec(2, (0, 'POWR 52[ELAINE_WTG3 0.6500]1'), scale=100.0)
            plot.subplot_spec(3, (0, 'P FLOW FROM POC'), title='Active Power at POC', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(4, (0, 'VARS 50[ELAINE_WTG1 0.6500]1'), title='Reactive Power at WTG', ylabel='Reactive power(MVar)', scale=100.0,offset=0.0)
            plot.subplot_spec(4, (0, 'VARS 52[ELAINE_WTG3 0.6500]1'), scale=100.0)
            plot.subplot_spec(5, (0, 'Q FLOW FROM POC'), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
            # plot.subplot_spec(5, (0, 'QSET(L) IN PU'), scale=83.6, offset=0.0 )
            # plot.subplot_spec(5, (0, 'QSET(L+5) IN PU'), scale=83.6, offset=0.0)
            # plot.subplot_spec(5, (0, 'QSET(L+9) IN PU'), scale=83.6, offset=0.0)
            # plot.subplot_spec(5, (0, 'QSET(L+25) IN PU'), scale=83.6, offset=0.0)
            # plot.subplot_spec(5, (0, 'QSET(L+79) IN PU'), scale=83.6, offset=0.0)
            # plot.subplot_spec(5, (0, 'QSET(L+81) IN PU'), scale=83.6, offset=0.0)
            # plot.subplot_spec(5, (0, 'QSET(L+94) IN PU'), scale=83.6, offset=0.0)
            # plot.subplot_spec(5, (0, 'QSET(L+96) IN PU'), scale=83.6, offset=0.0)
        elif 'FCT' in Test:
            plot.subplot_spec(0, (0, 'SYSTEM FREQUENCY'), title='Frequency', ylabel='Frequency (Hz)', scale=50.0, offset=50.0)
            plot.subplot_spec(1, (0, 'POC_VOLTAGE'), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(2, (0, 'WTG_PELEC_101'), title='Active Power at WTG', ylabel='Active Power(MW)', scale=100.0, offset=0.0)
            plot.subplot_spec(2, (0, 'WTG_PELEC_102'), scale=100.0)
            plot.subplot_spec(3, (0, 'P_POC'), title='Active Power at POC', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(4, (0, 'WTG_QELEC_101'), title='Reactive Power at WTG', ylabel='Reactive power(MVar)', scale=100.0,offset=0.0)
            plot.subplot_spec(4, (0, 'WTG_QELEC_102'), scale=100.0)
            plot.subplot_spec(5, (0, 'Q_POC'), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
        else:
            plot.subplot_spec(0, (0, 'POC F'), title='Frequency', ylabel='Frequency (Hz)', scale=50.0, offset=50.0)
            plot.subplot_spec(1, (0, 'POC V'), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(2, (0, 'POWR 50[ELAINE_WTG1 0.6500]1'), title='Active Power at WTG', ylabel='Active Power(MW)', scale=100.0, offset=0.0)
            plot.subplot_spec(2, (0, 'POWR 52[ELAINE_WTG3 0.6500]1'), scale=100.0)
            plot.subplot_spec(3, (0, 'P FLOW FROM POC'), title='Active Power at POC', ylabel='Active Power(MW)', scale=1.0, offset=0.0)
            plot.subplot_spec(4, (0, 'VARS 50[ELAINE_WTG1 0.6500]1'), title='Reactive Power at WTG', ylabel='Reactive power(MVar)', scale=100.0,offset=0.0)
            plot.subplot_spec(4, (0, 'VARS 52[ELAINE_WTG3 0.6500]1'), scale=100.0)
            plot.subplot_spec(5, (0, 'Q FLOW FROM POC'), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
        plot.plot(figname=os.path.splitext(filename)[0], show=0)
    os.chdir(cwd)
# ========================================================
def merge_pdf(dir_to_merge):
    import os
    global cwd
    copy_dir = os.path.dirname(dir_to_merge)
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
        new_dst_file = dir_to_merge+'\\'+ dir_to_merge.split(os.sep)[-1] +'_' + file
        os.rename(dst_file, new_dst_file)
        shutil.copy(new_dst_file,copy_dir)
    os.chdir(cwd)
# ========================================================
def APT():
    Test = 'APT'
    output_files_dir = os.path.join(cwd, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf"])
    plot_all_in_dir(output_files_dir,Test)
    merge_pdf(output_files_dir)
# ========================================================
def CAP_Manual():
    Test = 'CAP_Manual'
    output_files_dir = os.path.join(cwd, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf"])
    plot_all_in_dir(output_files_dir,Test)
    merge_pdf(output_files_dir)
# ========================================================
def RPT():
    Test = 'RPT'
    output_files_dir = os.path.join(cwd, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf"])
    plot_all_in_dir(output_files_dir,Test)
    merge_pdf(output_files_dir)
# ========================================================
def Tx_manual():
    Test = 'Tx_manual'
    output_files_dir = os.path.join(cwd, Test)
    dirCreateClean(output_files_dir, ["*.out", "*.pdf"])
    plot_all_in_dir(output_files_dir,Test)
    merge_pdf(output_files_dir)
# ========================================================
def VCT():
    Test = 'VCT'
    output_files_dir = os.path.join(cwd, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf"])
    plot_all_in_dir(output_files_dir,Test)
    merge_pdf(output_files_dir)
# ========================================================
def VDT():
    Test = 'VDT'
    output_files_dir = os.path.join(cwd, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf"])
    plot_all_in_dir(output_files_dir,Test)
    merge_pdf(output_files_dir)
def FCT():
    Test = 'FCT'
    output_files_dir = os.path.join(cwd, Test)
    # dirCreateClean(output_files_dir, ["*.out", "*.pdf"])
    plot_all_in_dir(output_files_dir,Test)
    merge_pdf(output_files_dir)
# =======================================================
def main():
    FCT()
    # APT()
    # CAP_Manual()
    # RPT()
    # Tx_manual()
    # VCT()
    # VDT()
# ============================================
if __name__ == '__main__':
    main()
