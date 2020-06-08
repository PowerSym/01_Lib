import time
import os,sys
import shutil
import Plotting as pl
import glob
from PyPDF2 import PdfFileMerger
timestr = time.strftime("%Y%m%d_%H%M%S")
cwd = os.getcwd()
output_files_dir = cwd
copy_dir = os.path.dirname(cwd)

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
    dirCreateClean(dir_to_plot,["*.pdf"])
    os.chdir(dir_to_plot)
    #filenames = glob.glob(os.path.join(dir_to_plot,'*.out'))    #pssefile = glob.glob(os.path.join(dir_to_plot,'*.inf'))
    #filenames = glob.glob(os.path.join(dir_to_plot,'*.inf'))
    #filenames = glob.glob(os.path.join(dir_to_plot,'*.xlsx'))
    filenames = glob.glob(os.path.join(dir_to_plot,'*.*'))
    for filename in filenames:
        filename = os.path.basename(filename)
        plot = pl.Plotting()
        #PSSE Channels and Excel PSSE standard
        if (filename[-4:].lower() in ('.out','xlsx')) and ('pscad' not in filename.lower()):
            plot.read_data(filename)
            #plot.read_data2(filename)
            plot.subplot_spec(0, (0, 'UUT_VOLTAGE'), title='Inverter Voltage', ylabel='Voltage (pu)',scale=1.0, offset = 0.0,timeoffset = 0.0)
            plot.subplot_spec(0, (0, 4))
            plot.subplot_spec(1, (0, 0), title='Power Factor', ylabel='Voltage (pu)',scale = 1.0, offset = 0.0, plot_PF = [5,6,12])
            plot.legends[1] = ['Power factor', 'Power Factor ref']
            plot.subplot_spec(2, (0, 2), title='Active Power at Interter terminal', ylabel='Active Power (MW)', scale=100.0, offset = 0.0)
            plot.subplot_spec(3, (0, 5), title='Active Power at POC', ylabel='Actice Power(MW)',scale = 1.0, offset = 0.0)#,P_Recovery = [5,2,1,2])
            plot.subplot_spec(4, (0, 3), title='Reactive power at Interter terminal', ylabel='Reactive Power(MVAr)', scale = 100.0, offset = 0.0)#,rstime = [2.0])
            plot.subplot_spec(5, (0, 6), title='Reactive Power at POC', ylabel='Reactive power(MVar)',scale = 0.01, offset = 0.0)#, plot_iq = [1,3,4,6,1,1.5,0])

            plot.plot(figname=os.path.splitext(filename)[0], show=0)
            # plot_iq = [V_INV_Channel_id,Q_INV_Channel_id,V_POC_Channel_id,Q_POC_Channel_id, tFault,dProp_IqCap,dProp_IqInd]
            # P_Recovery = [P__POC_Channel_id,tFault,clearnear,clearremote]
            # rstime = [tFault]
        #PSCAD Channel
        elif ('.inf' in filename.lower()):
            plot.read_data(filename)
            plot.subplot_spec(0, (0, 5), title='Inverter voltage', ylabel='Voltage (pu)',scale = 1.0, offset = 0.0,timeoffset = 0,xlimit = [0,20])
            plot.subplot_spec(1, (0, 4), title='POC voltage', ylabel='Voltage (pu)',scale = 1.0, offset = 0.0)
            plot.subplot_spec(2, (0, 10), title='Active Power at Interter terminal', ylabel='Active Power (MW)', scale=100.0, offset = 0.0)
            plot.subplot_spec(3, (0, 8), title='Active Power at POC', ylabel='Actice Power(MW)',scale = 1.0, offset = 0.0)
            plot.subplot_spec(4, (0, 7), title='Reactive power at Interter terminal', ylabel='Reactive Power(MVAr)', scale = 100.0, offset = 0.0)
            plot.subplot_spec(5, (0, 6), title='Reactive Power at POC', ylabel='Reactive power(MVar)',scale = 1.0, offset = 0.0)
            plot.legends[0] = ['Inverter voltage']
            plot.plot(figname=os.path.splitext(filename)[0], show=0)
        else:
            pass
    os.chdir(cwd)

# ========================================================
def plot_all_in_dir(dir_to_plot,Test):
    dirCreateClean(dir_to_plot,["*.pdf"])
    os.chdir(dir_to_plot)
    filenames = glob.glob(os.path.join(dir_to_plot,'*.out'))
    for filename in filenames:
        plot = pl.Plotting()
        plot.read_data(os.path.join(dir_to_plot, filename))

        if 'POC_Vdroop' in Test:
            plot.subplot_spec(0, (0, 7), title='POC voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0, plot_Vdroops_Hybrid = [14,22,9,0.0375,112],rstime = [250,260])  #droop = 0.0375/1.12
            plot.subplot_spec(0, (0, 14))
            #plot.legends[0] = ['POC_VOLTAGE', 'POC_V_REF', 'Vdroops_ref_Hybrids_Adjusted']

        elif Test == 'POC_PF_Setpoint_Step_Test_HP3':
            plot.subplot_spec(0, (0, 7), title='POC voltage', ylabel='Voltage (pu)', scale=1.0,offset=0.0, rstime=[8100, 12100])

        elif ('Capbank_Test' in Test) or ('Tap_Test' in Test):
            plot.subplot_spec(0, (0, 7), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(0, (0, 1))
            plot.subplot_spec(0, (0, 4))
            plot.subplot_spec(0, (0, 17))
            plot.subplot_spec(0, (0, 18))

        elif ('Reactive_Capability_Test' in Test) or ('POC_Qref_Step_Test' in Test):
            plot.subplot_spec(0, (0, 7), title='Voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)
            plot.subplot_spec(0, (0, 1))
            plot.subplot_spec(0, (0, 4))
            plot.subplot_spec(1, (0, 2), scale=100.0)
            plot.subplot_spec(1, (0, 5), scale=100.0)

        else:
            plot.subplot_spec(0, (0, 7), title='POC voltage', ylabel='Voltage (pu)', scale=1.0, offset=0.0)

        if Test == 'POC_Pref_Step_Test' or Test == 'POC_Pref_Step_Test_HP3':
            plot.subplot_spec(1, (0, 8), title='Active Power at POC', ylabel='Actice Power(MW)', scale=1.0, offset=0.0,ylimit=[0, 75])
            plot.subplot_spec(1, (0, 12),scale = 112.0)
        else:
            plot.subplot_spec(1, (0, 8), title='Active Power at POC', ylabel='Actice Power(MW)', scale=1.0, offset=0.0,ylimit=[0, 75])

        if Test == 'POC_Qref_Step_Test_HP3':
            plot.subplot_spec(2, (0, 9), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0, rstime = [490,520],xlimit=[0,700])
            plot.subplot_spec(2, (0, 13),scale = 112.0)

        elif Test =='Reactive_Capability_Test_HP3':
            plot.subplot_spec(2, (0, 9), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
            plot.subplot_spec(1, (0, 12),scale = 112.0)
            plot.subplot_spec(2, (0, 13), scale=112.0)

        elif 'POC_Vdroop' in Test:
            plot.subplot_spec(2, (0, 9), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0, rstime=[250, 260])
        elif Test == 'POC_PF_Setpoint_Step_Test_HP3':
            plot.subplot_spec(2, (0, 9), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0, rstime=[8100, 12100])
        else:
            plot.subplot_spec(2, (0, 9), title='Reactive Power at POC', ylabel='Reactive power(MVar)', scale=1.0,offset=0.0)
        plot.plot(figname=os.path.splitext(filename)[0], show=0)
    os.chdir(cwd)
# ========================================================    
def plot_R2_excel_files(dir_to_plot):
    os.chdir(dir_to_plot)
    dirCreateClean(dir_to_plot,["*.pdf"])
    filenames = glob.glob(os.path.join(dir_to_plot,'*.*'))
    for filename in filenames:
        filename = os.path.basename(filename)
        plot = pl.Plotting()
        #PSSE Channels and Excel PSSE standard
        if (filename[-4:].lower() in ('.out','xlsx')) and ('pscad' not in filename.lower()):
            plot.read_data2(filename)
            plot.subplot_spec(0, (0, 1), title='Active Power at POC', ylabel='Actice Power(MW)',scale = 1e-6, ylimit = [0,65])
            plot.subplot_spec(1, (0, 2), title='Reactive Power at POC', ylabel='Reactive power(MVar)',scale = 1e-6)
            plot.subplot_spec(2, (0, 3), title='Line-Line Voltages', ylabel='Voltage (kV)',scale = 1e-3/132)
            plot.subplot_spec(2, (0, 4),scale = 1e-3/132)
            plot.subplot_spec(2, (0, 5),scale = 1e-3/132)
            plot.legends[2] = ['V12','V23','V31']
            plot.subplot_spec(3, (0, 7), title='Current', ylabel='Current (Amp)')#,P_Recovery = [5,2,1,2])
            plot.subplot_spec(3, (0, 8))
            plot.subplot_spec(3, (0, 9))
            plot.legends[3] = ['I1','I2','I3']
            plot.subplot_spec(4, (0, 10), title='Frequency', ylabel='Frquency(Hz)',ylimit = [0,55])#,rstime = [2.0])
            plot.subplot_spec(5, (0, 6), title='Power Factor', ylabel='Power Factor')#, plot_iq = [1,3,4,6,1,1.5,0])
            plot.plot(figname=os.path.splitext(filename)[0], show=0)
    os.chdir(cwd)
    
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
        new_dst_file = dir_to_merge+'\\'+ dir_to_merge.split(os.sep)[-1] +'_'+os.path.splitext(os.path.basename(file))[0] +'_'+timestr+os.path.splitext(os.path.basename(file))[1]
        os.rename(dst_file, new_dst_file)
        #shutil.copy(new_dst_file,copy_dir)
plot_standard_files(output_files_dir)
#Test = 'POC_Vdroop_Hybrid_HP3'
#plot_all_in_dir(output_files_dir,Test)
#plot_R2_excel_files(output_files_dir)
merge_pdf(output_files_dir)
