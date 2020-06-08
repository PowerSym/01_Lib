import numpy as np
import time

class Plotting(object):
    def __init__(self):
        self.plotspec = [[] for _ in range(6)]
        self.timearrays = []
        self.dataarrays = []
        self.filenames = []
        self.offsets = []
        self.timeoffset = [0]
        self.scales = []
        self.xlimit = []
        self.ylimits = [[] for _ in range(6)]
        self.channel_names = []
        self.xlabels = [[] for _ in range(6)]
        self.ylabels = [[] for _ in range(6)]
        self.titles = [[] for _ in range(6)]
        self.legends = [[] for _ in range(6)]
        self.rstimes = [[] for _ in range(6)]
        self.plot_iqs = [[] for _ in range(6)]
        self.plot_Vdroops = [[] for _ in range(6)]
        self.plot_Vdroops_Hybrids = [[] for _ in range(6)]
        self.P_Recoverys = [[] for _ in range(6)]
        self.plot_PFs = [[] for _ in range(6)]

    def PF_plot(self,tDataArray,P_POC, Q_POC, PF_set):
        import matplotlib.pyplot as plt
        import math
        t = tDataArray
        tanpi = PF_set
        S_POC = []
        PF_ref = []
        PF = []
        for i in range(len(t)):
            Si = ((P_POC[i]**2)+(Q_POC[i]**2))**0.5
            if Si == 0:
                Si = 0.001
            S_POC.append(Si)
        for i in range(len(t)):
            PFi =P_POC[i]/S_POC[i] * np.sign(Q_POC[i])# math.copysign(P_pre[i]/S_pre[i],Q_pre[i])
            PF.append(PFi)
            PF_ref_i = math.cos(math.atan(tanpi[i]))* np.sign(Q_POC[i])
            PF_ref.append(PF_ref_i)
        plt.plot(t,PF,label="POC_PF",linewidth = 3.5) #change ax1. to plt.
        plt.plot(t,PF_ref,label="POC_PF_ref",linestyle="dashed", dashes=(4, 3),linewidth = 3.5)
        plt.autoscale(enable=True, axis='both', tight=False)
    def rise_settle_TimeCalc(self,tDataArray,outDataArray,tFault,Tfin):
        import matplotlib.pyplot as plt
        dataPTS = len(tDataArray)
        timeSubSet = tDataArray[:dataPTS/1]
        #QSubSet = outDataArray[:dataPTS/1]

        timeSubSet = tDataArray[:dataPTS]
        QSubSet = outDataArray[:dataPTS]

        
        QSubSetIndex = 0
        time = 0
        while time <= tFault:
            QSubSetIndex += 1
            time = timeSubSet[QSubSetIndex]
        Qini = QSubSet[QSubSetIndex-1]
        while time <= Tfin:
            QSubSetIndex += 1
            time = timeSubSet[QSubSetIndex]
        Qfin = QSubSet[QSubSetIndex-1]
        QSubSetIndexFin = QSubSetIndex
        Qchange = abs(Qfin-Qini)

        QSubSetIndex = 0
        x = 0

        #while x <= 0.005*Qchange:    # --> x <= 0.001*Qchange/Qchange
        #    QSubSetIndex += 1
        #    x = abs(QSubSet[QSubSetIndex]-Qini)
        #T_ini = timeSubSet[QSubSetIndex]
        T_ini = tFault
        # if (isinstance(tFault, float)) or (isinstance(tFault, list)):
        #     T_ini = tFault# Make it easy, the time that we change reference signal.
        # elif (isinstance(tFault, str)):
        #     while x <= 0.005*Qchange:    # --> x <= 0.001*Qchange/Qchange
        #         QSubSetIndex += 1
        #         x = abs(QSubSet[QSubSetIndex]-Qini)
        #     T_ini = timeSubSet[QSubSetIndex]

        QSubSetIndex = 0
        time = 0
        while time <= T_ini:
            time = timeSubSet[QSubSetIndex]
            QSubSetIndex += 1
        x = abs(QSubSet[QSubSetIndex] - Qini)

        while x <= 0.1*Qchange:
            QSubSetIndex += 1
            x = abs(QSubSet[QSubSetIndex]-Qini)
        T_10pc = timeSubSet[QSubSetIndex]

        while x <= 0.9*Qchange:
            QSubSetIndex += 1
            x = abs(QSubSet[QSubSetIndex]-Qini)
        T_90pc = timeSubSet[QSubSetIndex]

        QSubSetIndex = QSubSetIndexFin
        x = 0
        while x < 0.1*Qchange:
            QSubSetIndex -= 1
            x = abs(QSubSet[QSubSetIndex]-Qfin)
        T_setl = timeSubSet[QSubSetIndex]

        ymin, ymax = plt.ylim()
        xmin, xmax = plt.xlim()

        plt.plot([T_10pc,T_10pc],[ymin,ymax],linestyle='-', color='k', linewidth=1.2)
        plt.plot([T_90pc,T_90pc],[ymin,ymax],linestyle='-', color='k', linewidth=1.2)

        plt.plot([T_ini,T_ini],[ymin,ymax],linestyle='--', color='r', linewidth=1.2)
        plt.plot([T_setl,T_setl],[ymin,ymax],linestyle='--', color='r', linewidth=1.2)

        plt.autoscale(enable=True, axis='both', tight=False)
    
        riseTime = abs(T_90pc-T_10pc)
        setlTime = abs(T_setl-T_ini)

        ymax = np.amax(QSubSet)
        ymin = np.amin(QSubSet)
        
        xtext = xmin + 0.5 *(xmax-xmin)
        ytext = ymin + 0.5 *(ymax-ymin)
        ytext2 = ymin + 0.35 *(ymax-ymin)
        plt.text(xtext, ytext, r'Rise time:  %.3f s, Settling time: %.3f s'%(riseTime,setlTime),color='blue',fontsize=14)
        plt.text(xtext, ytext2, r'Change of %1.3f from %1.3f to %1.3f' %(Qchange,Qini,Qfin),color='blue',fontsize=14)

    def Iq_plot(self,tDataArray,V_INV,Q_INV, V_POC, Q_POC,tFault,dProp_IqCap,dProp_IqInd):
        import matplotlib.pyplot as plt
        INV_VolData = V_POC 
        INV_QData = Q_POC
        INV_QData = np.divide(INV_QData,100) # INV_QData here is in MVAr, remove it if INV_QData in pu.
        
        #INV_VolData = V_INV # this value can be change to change the access point, between V_INV and V_POC, in  pu.
        #INV_QData = Q_INV # this value can be change to change the access point, between Q_INV and Q_POC in Mvar.

        dNER_IqCap = 4 #Automatic Standard value: 4% change of Iq for 1% change in V: this one to plot what the Iq should be according to V during the faults
        dProp_IqCap = dProp_IqCap #The Propose value in GPS: 1.5% change for LVRT as per proposed in GPS
        dNER_IqInd = 6 #Automatic Standard value: 6% for HVRT ?
        dProp_IqInd = dProp_IqInd #The Propose value in GPS: 0% change for HVRT as per proposed in GPS
        csv_RI = -1
        tFault = tFault   #time when you apply fault, important.
        IqCHL = np.divide(INV_QData,INV_VolData)
        plt.plot(tDataArray,IqCHL,'--',label="Actual_Iq")
        plt.autoscale(enable=True, axis='both', tight=False)
        simt = 0
        tDataArrayIndex = 0
        while simt < tFault:
            simt = tDataArray[tDataArrayIndex]
            tDataArrayIndex += 1
        #Reading INV V and Q 5 data points before tFault
        INV_V_PreFault = INV_VolData[tDataArrayIndex-5]
        INV_Q_PreFault = INV_QData[tDataArrayIndex-5]
        INV_Iq_PreFault = INV_Q_PreFault/INV_V_PreFault

        #Find Iq requirement during the fault as per NER
        Iq_NER = []
        Iq_Proposed = []
        tDataArrayIndex = 0
        for t in tDataArray:
            INV_V = INV_VolData[tDataArrayIndex]
            if INV_V < 0.9:
                Iq_ExpNER = INV_Iq_PreFault+(dNER_IqCap*(INV_V_PreFault-INV_V))
                Iq_ExpProp = INV_Iq_PreFault+(dProp_IqCap*(INV_V_PreFault-INV_V))
                if abs(Iq_ExpNER)>1:
                    Iq_ExpNER = Iq_ExpNER/abs(Iq_ExpNER)
                if abs(Iq_ExpProp)>1:
                    Iq_ExpProp = Iq_ExpProp/abs(Iq_ExpProp)
                Iq_NER.append(Iq_ExpNER)
                Iq_Proposed.append(Iq_ExpProp)
            elif INV_V > 1.1:
                Iq_ExpNER = INV_Iq_PreFault+(dNER_IqInd*(INV_V_PreFault-INV_V))
                Iq_ExpProp = INV_Iq_PreFault+(dProp_IqInd*(INV_V_PreFault-INV_V))
                if abs(Iq_ExpNER)>1:
                    Iq_ExpNER = Iq_ExpNER/abs(Iq_ExpNER)
                if abs(Iq_ExpProp)>1:
                    Iq_ExpProp = Iq_ExpProp/abs(Iq_ExpProp)
                Iq_NER.append(Iq_ExpNER)
                Iq_Proposed.append(Iq_ExpProp) # Have to update this one
            else:
                Iq_NER.append(INV_Iq_PreFault)
                Iq_Proposed.append(INV_Iq_PreFault)
            tDataArrayIndex += 1
        plt.plot(tDataArray,Iq_NER,':',label="NER_Auto_Iq")
        plt.plot(tDataArray,Iq_Proposed,'k:',label="Proposed_Iq")
        plt.autoscale(enable=True, axis='both', tight=False)

    def Vdroop_adjusted_plot(self,tDataArray,POC_V_ref,Q_POC, QDROOP, SBASE):
        import matplotlib.pyplot as plt
        V_Droop_error = np.divide((np.multiply(Q_POC,QDROOP)),SBASE)
        V_Droop_Adjusted = np.subtract(POC_V_ref,V_Droop_error)
        plt.plot(tDataArray,V_Droop_Adjusted,'--',label="V_Droop_Adjusted",linewidth = 3.5)
        plt.autoscale(enable=True, axis='both', tight=False)

    def Vdroop_Hybrid_adjusted_plot(self,tDataArray,POC_V_ref,POC_V_err,Q_POC, QDROOP, SBASE):
        import matplotlib.pyplot as plt
        V_Droop_error = np.divide((np.multiply(Q_POC,QDROOP)),SBASE)
        V_Adjusted = np.subtract(POC_V_ref,POC_V_err)
        V_Droop_Hybrid_Adjusted = np.subtract(V_Adjusted, V_Droop_error)
        plt.plot(tDataArray,V_Droop_Hybrid_Adjusted,'--',label="V_Droop_Hybrid_Adjusted",linewidth = 3.5)
        plt.autoscale(enable=True, axis='both', tight=False)

    def P_Recovery_Calc(self,tDataArray,INV_PData,P__POC_Channel_id,tFault,clearnear,clearremote):
        #from 100 milliseconds after clearance of the fault, active power of at least 95% of the level existing just prior to the fault.
        import matplotlib.pyplot as plt
        dataPTS = len(tDataArray)
        tFault = tFault #time to apply the fault
        Pmax = max(INV_PData)
        Pmin = min(INV_PData)
        #Find pre disturbance P
        simt = 0
        tDataArrayIndex = 0
        while simt < tFault:
            simt = tDataArray[tDataArrayIndex]
            tDataArrayIndex += 1
        #Reading INV V and Q 5 data points before tFault
        INV_P_PreFault = INV_PData[tDataArrayIndex-5]
        plt.plot([tDataArray[0],tDataArray[dataPTS-1]],[INV_P_PreFault*0.95,INV_P_PreFault*0.95],'r--',label="0.95xP_PreFAULT") #plot a horizental line at 0.95P pre fault
        plt.autoscale(enable=True, axis='both', tight=False)
        ymin, ymax = plt.ylim()
        tFClearN = clearnear + tFault
        tFClearR = clearremote + tFault
        tFClear = max(tFClearN,tFClearR)
        plt.plot([(tFClear+0.1),(tFClear+0.1)],[ymin,ymax],'r:') #plot vertical line at 100ms after clerance fault
        plt.autoscale(enable=True, axis='both', tight=False)
        plt.ylim(ymin,ymax) # This was inserted to reset ylim after previous plotting commands
        while simt < tFClear+0.1: #after 100ms following clearing the fault
            simt = tDataArray[tDataArrayIndex]
            tDataArrayIndex += 1
        #Reading P output after 100ms following clearing the fault
        INV_P_PostFault = INV_PData[tDataArrayIndex]    #P value after 100ms fault

        #if INV_P_PostFault >= INV_P_PreFault*0.95:
        #    plt.text((simt+0.6), (ymax+ymin)/2, r'P_recovery > 0.95xP_pre-flt',color='blue',fontsize=9)
        #else:
        #    plt.text((simt+0.6), (ymax+ymin)/2, r'P_recovery < 0.95xP_pre-flt',color='red',fontsize=9)

        simt = 0
        tDataArrayIndex = 0
        
        while simt < tFClear: #after 100ms following clearing the fault
            simt = tDataArray[tDataArrayIndex]
            tDataArrayIndex += 1

        x = 0
        while x <= INV_P_PreFault*0.95:
            x = INV_PData[tDataArrayIndex]
            tDataArrayIndex += 1
        T_95 = tDataArray[tDataArrayIndex]
        T_recovery = (T_95 - tFClear)*1000

        if INV_P_PostFault >= INV_P_PreFault*0.95:
            plt.text((simt+0.6), (ymax+ymin)/2, r'P_recovery > 0.95xP_pre-flt',color='blue',fontsize=9)
            plt.text((simt+0.6), (ymax+ymin)/4, r'Recovery time = %5.2f (ms)'%T_recovery,color='blue',fontsize=9)
        else:
            plt.text((simt+0.6), (ymax+ymin)/2, r'P_recovery < 0.95xP_pre-flt',color='red',fontsize=9)
            plt.text((simt+0.6), (ymax+ymin)/4, r'Recovery time = %5.2f (ms)'%T_recovery,color='red',fontsize=9)
    
    def read_data(self, filename):
        # Read PSSE data

        if (filename[-4:].lower() in '.out') and ('pscad' not in filename.lower()):
            print 'Plotting file: ' + str(filename)
            self.filenames.append(filename)
            # PSSE output fileslen(time)
            import os, psse34
            import dyntools

            outfile_data = dyntools.CHNF(os.path.basename(filename))
            short_title, chanid, chandata = outfile_data.get_data()
            
            time = np.array(chandata['time'])
            data = np.zeros((len(chandata) - 1, len(time)))
            
            chan_ids = []
            for key in chandata.keys()[:-1]:
                print key, chanid[key]
                data[key - 1] = chandata[key]
                chan_ids.append(chanid[key])
            self.timearrays.append(time)
            self.dataarrays.append(data)
            self.scales.append(np.ones(len(data)+1))
            self.offsets.append(np.zeros(len(data)+1))
            self.channel_names = chan_ids
        #Read PSCAD file
        if filename[-4:].lower() in ('*.inf'):
            print 'Plotting file: ' + str(filename)
            # PSCAD output files
            import os, glob, re
            cwd = os.getcwd() 
            outfiles = []
            for pscadfile in glob.glob(cwd + "/" + filename[0:-4] + "_*.out"):
                outfiles.append(pscadfile)
            #
            chan_ids = []
            f = open(filename, 'r')
            lines = f.readlines()
            f.close()
            for idx, rawline in enumerate(lines):
                line = re.split(r'\s{1,}', rawline)
                chan_ids.append(line[2][6:-1])
            #
            # Setup time array using last outfile
            time = []
            f = open(outfiles[0], 'r')
            lines = f.readlines()
            f.close()
            for l_idx, rawline in enumerate(lines[1:]):
                line = re.split(r'\s{1,}', rawline)
                time.append(float(line[1]))
            time = np.array(time)
            #
            # Setup empty data array using known number of channels and time steps
            data = np.zeros((len(chan_ids), len(time)))
            #
            # Populate data array using all outfiles
            for f_idx, outfile in enumerate(outfiles):
                f = open(outfile, 'r')
                lines = f.readlines()
                f.close()
                for n, rawline in enumerate(lines[1:]):
                    line = re.split(r'\s{1,}', rawline)
                    for m, value in enumerate(line[2:-1]):
                        data[(f_idx*10) + m][n] = value
            for n, chan_id in enumerate(chan_ids):
                print n+1, chan_id
                
            # Add data to arrays in class
            self.timearrays.append(time)
            self.dataarrays.append(data)
            self.scales.append(np.ones(len(data)+1))
            self.offsets.append(np.zeros(len(data)+1))
            self.channel_names = chan_ids


        #Read Excel PSSE standard file
        if filename[-5:].lower() in ('*.xlsx'):
            print 'Plotting file: ' + str(filename)
            import os
            from os.path import join, dirname, abspath, isfile
            from collections import Counter
            import xlrd
            from xlrd.sheet import ctype_text
            cwd = os.getcwd()
            import psse34
            import dyntools
            def read_excel_tools(filename,sheet_index):
                xl_workbook = xlrd.open_workbook(filename)
                xl_sheet = xl_workbook.sheet_by_index(sheet_index)
                return xl_sheet
            def get_data(xl_sheet):
                chanid = {}
                chandata = {}
                row = xl_sheet.row(2)
                short_title = str(xl_sheet.row(1)[0].value) + '\n' + str(xl_sheet.row(2)[0].value)

                row_chan = xl_sheet.row(3)
                row_title = xl_sheet.row(4)
                for chan_idx in range(len(row_chan)):
                    chan_idx_data = []
                    for row_idx in range(5, xl_sheet.nrows):
                        cell_obj = xl_sheet.cell(row_idx, chan_idx)
                        chan_idx_data.append(cell_obj.value)
                    if chan_idx ==0:
                        chanid['time'] = str(row_title[chan_idx].value)
                        chandata['time'] = chan_idx_data
                    else:
                        chanid[int(row_chan[chan_idx].value)] = str(row_title[chan_idx].value)
                        chandata[int(row_chan[chan_idx].value)] = chan_idx_data
                return short_title, chanid, chandata

            xl_sheet = read_excel_tools(filename,sheet_index = 0)
            short_title, chanid, chandata = get_data(xl_sheet)
            time = np.array(chandata['time'])
            data = np.zeros((len(chandata) - 1, len(time)))
            chan_ids = []
            for key in chandata.keys()[:-1]:
                print key, chanid[key]
                data[key - 1] = chandata[key]
                chan_ids.append(chanid[key])
            print '\n'
            self.timearrays.append(time)
            self.dataarrays.append(data)
            self.scales.append(np.ones(len(data)+1))
            self.offsets.append(np.zeros(len(data)+1))
            self.channel_names = chan_ids

    def read_data2(self, filename):
        #Read Excel file
        if filename[-5:].lower() in ('*.xlsx'):
            print 'Plotting file: ' + str(filename)
            import os
            from os.path import join, dirname, abspath, isfile
            from collections import Counter
            import xlrd
            from xlrd.sheet import ctype_text
            cwd = os.getcwd()
            import psse34
            import dyntools
            def read_excel_tools(filename,sheet_index):
                xl_workbook = xlrd.open_workbook(filename)
                xl_sheet = xl_workbook.sheet_by_index(sheet_index)
                return xl_sheet
            def get_data(xl_sheet):
                import re
                chanid = {}
                chandata = {}
                title_row = 1
                title_col = 0
                chan_number_row = 0
                title_row = 0
                short_title = os.path.splitext(filename)[0]
                row_chan = xl_sheet.row(chan_number_row)
                row_title = xl_sheet.row(title_row)
                for chan_idx in range(len(row_chan)):
                    chan_idx_data = []
                    for row_idx in range(1, xl_sheet.nrows):
                        cell_obj = (xl_sheet.cell(row_idx, chan_idx)).value
                        if (isinstance(cell_obj, unicode)):
                            if len(re.findall(("\d+\.\d+"), cell_obj)) >0:
                                cell_obj = float(re.findall("\d+\.\d+", cell_obj)[0])
                            elif len(re.findall(("\d+"), cell_obj)) > 0:
                                cell_obj = float(re.findall("\d+", cell_obj)[0])
                            else:
                                cell_obj = 0
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

            xl_sheet = read_excel_tools(filename,sheet_index = 1)
            short_title, chanid, chandata = get_data(xl_sheet)
            time = np.array(chandata['time'])
            data = np.zeros((len(chandata) - 1, len(time)))
            chan_ids = []
            print(chandata[1])

            for key in chandata.keys()[:-1]:
                print key, chanid[key]
                data[key - 1] = chandata[key]
                chan_ids.append(chanid[key])
            print '\n'
            self.timearrays.append(time)
            self.dataarrays.append(data)
            self.scales.append(np.ones(len(data)+1))
            self.offsets.append(np.zeros(len(data)+1))
            self.channel_names = chan_ids

    def subplot_spec(self, subplot, plot_arrays, title = '', ylabel = '',xlimit='', ylimit ='', scale = 1.0, offset = 0.0, rstime ='',plot_iq ='', P_Recovery='', timeoffset='',plot_Vdroop ='',plot_Vdroops_Hybrid ='',plot_PF =''):
        """
        Specify which channels are to be plotted

        Args:
            subplot: Number of subplot for which we are specifying files 
            and channels
            
            plot_arrays: The plot_arrays input should be specified as a two element tuple, 
            with the first element being the input file and the second being the channel 
            of that file to be plotted. The second element can be specified either 
            as the channel number (as an int), or the name as shown in 
            self.channel_names (as a str). If a string is provided, only the first
            few letters need to be specified sufficient to be unique from the other
            channel_names. If the string is not unique then the first match will be 
            plotted.

            title: Subplot title (optional)
            
            ylabel: Y axis title (optional)
            
            scale: Constant multiplier for this trace (optional)
            
            offset: Y axis offset for this trace (optional)

        Returns:
            Nothing

        Raises:
            Nothing

        """
        if plot_arrays[0] < len(self.dataarrays):
            if title != '':
                self.titles[subplot] = title
            if ylabel != '':
                self.ylabels[subplot] = ylabel
            if ylimit != '':
                self.ylimits[subplot] = ylimit
            if xlimit != '':
                self.xlimit = xlimit
            if rstime != '':
                self.rstimes[subplot] = rstime
            if plot_iq != '':
                self.plot_iqs[subplot] = plot_iq
            if plot_Vdroop != '':
                self.plot_Vdroops[subplot] = plot_Vdroop

            if plot_Vdroops_Hybrid != '':
                self.plot_Vdroops_Hybrids[subplot] = plot_Vdroops_Hybrid

            if P_Recovery != '':
                self.P_Recoverys[subplot] = P_Recovery

            if timeoffset != '':
                self.timeoffset = [timeoffset]

            if plot_PF != '':
                self.plot_PFs[subplot] = plot_PF

            #
            # if type(plot_arrays[1]) is str:  #self.channel_names
            #     chars = len(plot_arrays[1])
            #     for idx, name in enumerate(self.channel_names[plot_arrays[0]]):
            #         if name[:chars] == plot_arrays[1]:
            #             plot_channel = idx + 1  # Channel numbers start at 1
            #             break

            if type(plot_arrays[1]) is str:  #self.channel_names
                for n, chan_id in enumerate(self.channel_names):
                    if plot_arrays[1].lower() == self.channel_names[n].lower():
                        plot_channel = n+1 # Channel numbers start at 1
                        break
            else:
                plot_channel = plot_arrays[1]
            #
            self.plotspec[subplot].append((plot_arrays[0], plot_channel))
            self.scales[plot_arrays[0]][plot_channel] = scale
            self.offsets[plot_arrays[0]][plot_channel] = offset
        else:
            print 'Specified file number {:d} is not in memory'.format(plot_arrays[0])

            
    def plot(self, figname = '', show = 0):
        '''
        Create plot of channels and files as specified in self.plotspec.\n
        Manually specify the following plot variables as required:\n
            self.legends[subplot]\n
            self.titles[subplot]\n
            self.ylabels[subplot]\n
            self.ylimit[subplot]\n
            self.xlimit
        '''
        import matplotlib.pyplot as plt

        if show:
            plt.ion() #Turn the interactive mode on.
            plt.pause(0.0001)

        from matplotlib.ticker import ScalarFormatter, FormatStrFormatter

        for idx, i in enumerate(self.plotspec):
            if i != []:
                subplots = idx + 1
        
        if subplots == 1:
            subplot_index = [111]
            nRows = 1
            nCols = 1
            #plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
            plt.figure(figsize=[(2 + 18 * nCols), (2 + 10 * nRows)])
            #plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98, left = 0.1, hspace = 0.25)
        if subplots == 2:
            subplot_index = [211, 212]
            nRows = 2
            nCols = 1
            #plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
            plt.figure(figsize=[(2 + 18 * nCols), (2 + 5 * nRows)])
            #plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98, left = 0.1, hspace = 0.28)
        if subplots == 3:
            subplot_index = [311, 312, 313]
            nRows = 3
            nCols = 1
            #plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
            plt.figure(figsize=[(2 + 24 * nCols), (2 + 5 * nRows)])
            #plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98,  left = 0.1, hspace = 0.32, wspace = 0.2)
        if subplots == 4:
            nRows = 2
            nCols = 2
            subplot_index = [221, 222, 223, 224]
            #plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
            plt.figure(figsize=[(2 + 12 * nCols), (2 + 6.5 * nRows)])
            #plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98, left = 0.1, hspace = 0.32, wspace = 0.2)
        if subplots == 5:
            subplot_index = [311, 323, 324, 325, 326]
            nRows = 3
            nCols = 2
            #plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
            plt.figure(figsize=[(2 + 12 * nCols), (2 + 5 * nRows)])
            #plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98, left = 0.1, hspace = 0.35, wspace = 0.2)
        if subplots == 6:
            nRows = 3
            nCols = 2
            subplot_index = [321, 322, 323, 324, 325, 326]
            #plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
            plt.figure(figsize=[(2 + 12 * nCols), (2 + 5 * nRows)])
            #plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98, left = 0.1, hspace = 0.37, wspace = 0.2)
        import os
        plot_title = os.path.basename(figname)
        plt.suptitle(plot_title,fontsize = 22,fontweight ='bold')
        
        #
        plt.rc('xtick', labelsize = 18)
        plt.rc('ytick', labelsize = 18)
        
        #print 'Number of subplots: ' + str(subplots)
        #print 'Detail plots:' + str(self.plotspec)
        
        for idx in range(subplots):
            plt.subplot(subplot_index[idx])
            ax = plt.gca()
            # ax.yaxis.set_major_formatter(FormatStrFormatter('%1.{:d}f'.format(2)))            
            ax.ticklabel_format(useOffset=False)

            legend_chan_names = []
            for trace in self.plotspec[idx]:
                datafile = trace[0]
                channel = trace[1]
                if channel != 0:
                    plt.plot(self.timearrays[datafile] + self.timeoffset[datafile],
                                                        (self.dataarrays[datafile][channel-1]
                                                        * self.scales[datafile][channel])
                                                        + self.offsets[datafile][channel], linewidth = 3.5)

                chan_name =  self.channel_names[channel-1]
                legend_chan_names.append(chan_name)

            #Plot raise time and settling time
            #trace = (self.plotspec)[idx][len(self.plotspec[idx])-1] #only plot settling time rise time at last
                

            #Plot Iq
            trace = (self.plotspec)[idx][0]  #only plot iq at first              
            datafile = trace[0]
            channel = trace[1]
            if self.plot_iqs[idx] != []:
                V_INV_Channel_id,Q_INV_Channel_id,V_POC_Channel_id,Q_POC_Channel_id, tFault,dProp_IqCap,dProp_IqInd = self.plot_iqs[idx]
                self.Iq_plot(self.timearrays[datafile] + self.timeoffset[datafile],self.dataarrays[datafile][V_INV_Channel_id-1],self.dataarrays[datafile][Q_INV_Channel_id-1],
                             self.dataarrays[datafile][V_POC_Channel_id-1],self.dataarrays[datafile][Q_POC_Channel_id-1],tFault,dProp_IqCap,dProp_IqInd)



            if self.plot_Vdroops[idx] != []:
                POC_V_ref_Channel_id,Q_POC_Channel_id, QDROOP, SBASE = self.plot_Vdroops[idx]
                self.Vdroop_adjusted_plot(self.timearrays[datafile] + self.timeoffset[datafile],
                                          self.dataarrays[datafile][POC_V_ref_Channel_id - 1],
                                          self.dataarrays[datafile][Q_POC_Channel_id - 1],QDROOP,SBASE)

            if self.plot_Vdroops_Hybrids[idx] != []:
                POC_V_ref_Channel_id,V_err_Channel_id,Q_POC_Channel_id, QDROOP, SBASE = self.plot_Vdroops_Hybrids[idx]
                self.Vdroop_Hybrid_adjusted_plot(self.timearrays[datafile] + self.timeoffset[datafile],
                                                 self.dataarrays[datafile][POC_V_ref_Channel_id - 1],
                                                 self.dataarrays[datafile][V_err_Channel_id - 1],
                                                 self.dataarrays[datafile][Q_POC_Channel_id - 1],QDROOP,SBASE)


            #Plot P Recovery
            trace = (self.plotspec)[idx][0]  #only plot iq at first              
            datafile = trace[0]
            channel = trace[1]            
            if self.P_Recoverys[idx] != []:
                P_POC_Channel_id,tFault,clearnear,clearremote = self.P_Recoverys[idx]
                self.P_Recovery_Calc(self.timearrays[datafile] + self.timeoffset[datafile],self.dataarrays[datafile][P__POC_Channel_id-1],P_POC_Channel_id,tFault,clearnear,clearremote)

            # Plot PF
            trace = (self.plotspec)[idx][0]  # only plot iq at first
            datafile = trace[0]
            channel = trace[1]
            if self.plot_PFs[idx] != []:
                P_POC_Channel_id, Q_POC_Channel_id, PF_set_Channel_id = self.plot_PFs[idx]
                self.PF_plot(self.timearrays[datafile] + self.timeoffset[datafile],
                             self.dataarrays[datafile][P_POC_Channel_id - 1],
                             self.dataarrays[datafile][Q_POC_Channel_id - 1],
                             self.dataarrays[datafile][PF_set_Channel_id - 1])

            trace = (self.plotspec)[idx][0]  # only plot settling time rise time at first
            datafile = trace[0]
            channel = trace[1]
            if self.rstimes[idx] != []:
                tFault = self.rstimes[idx][0]
                Tfin = self.rstimes[idx][1]
                self.rise_settle_TimeCalc(self.timearrays[datafile] + self.timeoffset[datafile],
                                          (self.dataarrays[datafile][channel - 1] * self.scales[datafile][channel]) +
                                          self.offsets[datafile][channel], tFault, Tfin)

            '''top = 0.9
            if self.titles[idx] == "":
                top = 0.95
            plt.subplots_adjust(right=0.9, hspace=0.3, wspace=0.3, top=top, left=0.1, bottom=0.12)'''

            #if self.legends[idx] != []:
            top = 0.9
            if self.titles[idx] == "":
                top = 0.95
            plt.subplots_adjust(right=0.9, hspace=0.3, wspace=0.3, top=top, left=0.1, bottom=0.12)
            #plt.legend(self.legends[idx], prop = {'size': 16}, loc = 'center right')

            if self.legends[idx] != []:
                legend = plt.legend(self.legends[idx],bbox_to_anchor=(0., -0.25, 1., .105), loc=2,
                                    ncol=2, mode="expand", prop={'size': 16}, borderaxespad=0.1,frameon =False)
                legend.get_title().set_fontsize('16')  # legend 'Title' fontsize
                plt.setp(plt.gca().get_legend().get_texts(), fontsize='16')  # legend 'list' fontsize
            else:
                if self.plot_Vdroops[idx] != []:
                    legend_chan_names.append('V_Droop_Ref_Adjusted')
                    legend = plt.legend(legend_chan_names,bbox_to_anchor=(0., -0.25, 1., .105), loc=2,
                                        ncol=2, mode="expand", prop={'size': 16}, borderaxespad=0.1,frameon =False)
                elif self.plot_Vdroops_Hybrids[idx] != []:
                    legend_chan_names.append('Vdroops_ref_Hybrids_Adjusted')
                    legend = plt.legend(legend_chan_names,bbox_to_anchor=(0., -0.25, 1., .105), loc=2,
                                        ncol=2, mode="expand", prop={'size': 16}, borderaxespad=0.1,frameon =False)

                else:
                    legend = plt.legend(legend_chan_names,bbox_to_anchor=(0., -0.25, 1., .105), loc=2,
                                        ncol=2, mode="expand", prop={'size': 16}, borderaxespad=0.1,frameon =False)

            plt.subplots_adjust(hspace=0.5)

            if self.titles[idx] != []:
                plt.title(self.titles[idx], fontsize = 20,fontweight ='bold')
            #
            if self.ylabels[idx] != []:
                plt.ylabel(self.ylabels[idx], fontsize = 18)
                plt.ylabel(self.ylabels[idx], labelpad=20)
            #
            if self.ylimits[idx] != []:
                plt.ylim(self.ylimits[idx])
            #
            if self.xlimit != []:
                plt.xlim(self.xlimit)
            else:
                xmin, xmax = ax.get_xlim()
                ticks = ax.get_xticks()
                plt.xlim([0, ticks[-2]])

                #'xtick', labelsize=SMALL_SIZE
            plt.xlabel('Time (sec)', fontsize = 16)
            plt.grid(1)
            
        if figname != '':
            #plt.tight_layout() #pad=0.4, w_pad=0.5, h_pad=1.0)
            #plt.savefig(figname + '.png', format = 'png')
            #plt.savefig(figname + '.pdf', format = 'pdf')
            plt.savefig(figname + '.pdf', format = 'pdf')
            plt.close()
            #plt.clf()
            #plt.savefig(figname + '.emf', format = 'emf')
            #plt.savefig(figname + '.svg', format = 'svg')
            #plt.savefig(figname + '.eps', format='eps')

        if show:
            plt.show(block=True)
