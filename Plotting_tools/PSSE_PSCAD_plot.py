# Ensure that the user is running v2.5 of python
# Defines the path of the PSSE/psspy library
import glob, os, sys, math, csv, time, logging, traceback, os.path
import numpy
from PyPDF2 import PdfFileMerger, PdfFileWriter, PdfFileReader
from matplotlib import pyplot as plt
from matplotlib import ticker

# Defines the path of the PSSE/psspy library
PSSE_PATH = r'C:\Program Files (x86)\PTI\PSSE34\PSSBIN'
sys.path.append(PSSE_PATH)

# Tells PSSE where it is itself. Will eventually cause error without
os.environ['PATH'] += ';' + PSSE_PATH

# Imports the psspy library
import psse34
import psspy
import dyntools

def restrict_y_axis(ylimit, data_range, max_data, min_data):
    """ Used to limit the y axis
    :param ylimit:
    :param data_range:
    :param max_data:
    :param min_data:
    :return:
    """
    new_y = [ylimit[0], ylimit[1]]

    if abs(ylimit[1] - ylimit[0]) < abs(0.02 * numpy.mean(ylimit)) and abs(numpy.mean(ylimit)) > 30:
        new_y[1], new_y[0] = 1.01 * numpy.mean(ylimit), 0.99 * numpy.mean(ylimit)
    elif abs(ylimit[1] - ylimit[0]) < abs(0.2 * numpy.mean(ylimit)) and abs(numpy.mean(ylimit)) > 1:
        new_y[1], new_y[0] = 1.1 * numpy.mean(ylimit), 0.9 * numpy.mean(ylimit)
    elif abs(numpy.mean(ylimit)) < 1 and abs(ylimit[1] - ylimit[0]) < 0.05:
        new_y[1], new_y[0] = numpy.mean(ylimit) + 0.25, numpy.mean(ylimit) - 0.25
    else:
        if new_y[1] < (0.125 * data_range + max_data):
            new_y[1] = 0.125 * data_range + max_data
        if new_y[0] > (-0.125 * data_range + min_data):
            new_y[0] = -0.125 * data_range + min_data
    return new_y

def find_row_col(count):
    if (count == 1):
        nRows = 1
        nCols = 1
    elif (count == 2):
        nRows = 2
        nCols = 1
    elif (count == 3):
        nRows = 3
        nCols = 1
    elif (count == 4):
        nRows = 2
        nCols = 2
    elif (count in [5, 6]):
        nRows = 3
        nCols = 2
    elif (count in [7, 8, 9]):
        nRows = 3
        nCols = 3
    elif (count in [10, 11, 12]):
        nRows = 4
        nCols = 3
    else:
        nRows = 4
        nCols = 3
    return nRows, nCols


def error_plot(out_folder, out_file, error_message):
    print("ERROR: UNREADABLE OUT FILE")
    fig = plt.figure()
    ax = fig.add_subplot(111)
    ax.text(5, 5, error_message, verticalalignment='center', horizontalalignment='center', fontsize=12)
    ax.axis([0, 10, 0, 10])
    out_file = "_UNREADABLE_OUT_FILE_" + out_file

    file_str = os.path.join(out_folder, out_file).replace('.out', '.png')

    plt.savefig(file_str, dpi=100)
    print(str(file_str + "Saved"))
    plt.close()
    return

# Documentation DONE
def plot(out_folder, out_file, out_chans, outchans_adj, plot_title=["NA"] * 12, plot_x_label=["Time (s)"] * 12,
         plot_y_label=[""] * 12,
         top_title="", xlims="", ylims="", angle=[0, 0, 0, 0, 0, 0, 0, 0, 0, 0],xoffset ="", step_boundary=""):
    '''if isinstance(out_file, basestring):
        out_file = [out_file]
    elif "_PSSE" in out_file[1].lower():
        out_file[0], out_file[1] = out_file[1], out_file[0]'''

    if out_chans is None:
        out_chans = []
    if outchans_adj is None:
        outchans_adj = []

    if not (ylims == "" or isinstance(ylims[0], list)):
        ylims = [ylims] * 12

    if not (xlims == "" or isinstance(xlims[0], list)):
        xlims = [xlims] * 12

    # for x in xrange(len(out_chans)):
    #     if not isinstance(out_chans, list):
    #         out_chans[x] = [out_chans[x]]

    _ch_id, _ch_data, _sh_tt, time = [], [], [], []
    try:
        chnfobj = dyntools.CHNF(out_file[0])
        sh_tt_0, ch_id_0, ch_data_0 = chnfobj.get_data()
        _ch_data.append(ch_data_0)
        _ch_id.append(ch_id_0)
        _sh_tt.append(sh_tt_0)
        #print(len(ch_data_0))
        for i in range(len(ch_data_0['time'])):
            ch_data_0['time'][i] += xoffset
        time.append(ch_data_0['time'])
    except OverflowError:
        error_plot(out_folder, out_file[0],
                   "The OUT file seems to have infinite data in it and cannot be read.\nThis case is considered a failed case.")
        return
    print('PSSE channels: ' + str(ch_id_0))
    #read PSCAD
    filename = out_file[1]
    # PSCAD output files
    import glob, re
    import os
    import numpy as np
    #
    outfiles = []
    for pscadfile in glob.glob(filename[0:-4] + "_*.OUT"):
        outfiles.append(pscadfile)
    #Find the channel ids name
    chan_ids = []
    f = open(filename, 'r')
    lines = f.readlines()
    f.close()
    for idx, rawline in enumerate(lines):
        line = re.split(r'\s{1,}', rawline)
        chan_ids.append(line[2][6:-1])

    # change format to similar to PSSE
    ch_id_1 = {}
    for i in range(len(chan_ids)):
        ch_id_1[i+1] = chan_ids[i]
        ch_id_1['time'] = 'Time(s)'
    

    time1 = []
    f = open(outfiles[-1], 'r')
    lines = f.readlines()
    f.close()
    for l_idx, rawline in enumerate(lines[1:]):
        line = re.split(r'\s{1,}', rawline)
        time1.append(float(line[1]))    
   
    #
    # Setup empty data array using known number of channels and time steps
    data1 = np.zeros((len(ch_id_1), len(time1)))
    #
    # Populate data array using all outfiles
    for f_idx, pscadout in enumerate(outfiles):
        f = open(pscadout, 'r')
        lines = f.readlines()
        f.close()
        for n, rawline in enumerate(lines[1:]):
            line = re.split(r'\s{1,}', rawline)
            for m, value in enumerate(line[2:-1]):
                data1[(f_idx * 10) + 1 + m][n] = value
            data1[0][n] = line[1]
    ch_data_1 = {}
    for i in range(len(data1)-1):
        ch_data_1[i+1] = data1[i+1].tolist()
    ch_data_1['time'] = data1[0].tolist()
    sh_ttl = sh_tt_0
    # Setup time array using data
    print('PSCAD channels: ' + str(ch_id_1))
    _ch_data.append(ch_data_1)
    _ch_id.append(ch_id_1)
    # _sh_tt.append(sh_tt_0)
    time.append(data1[0].tolist())

    # loop through each of the channels required and extract the data
    data = []
    angle_data = []
    for y in xrange(len(out_file)):  #len(outfile) = 2, psse and pscad inf file name
        j = 0
        data.append([])
        for k, chan in enumerate(out_chans[y]):
            if isinstance(chan, list):
                data[y].append([])
                for i in xrange(len(chan)):
                    scale = outchans_adj[y][k][i][0]
                    offset = outchans_adj[y][k][i][1]
                    if type(chan[i]) is str:
                        for key in _ch_id[y]:
                            if chan[i] == _ch_id[y][key]:
                                chan[i] = key
                                break
                    add = (np.array(_ch_data[y][chan[i]]) * scale + offset).tolist()
                    # data[y][j].append(_ch_data[y][chan[i]])
                    data[y][j].append(add)
            else:
                data[y].append([])
                data[y][j].append(_ch_data[y][chan])

            j += 1
        for angle_chan in angle:
            angle_data.append([])
            try:
                angle_data[y].append(_ch_data[y][angle_chan])
            except Exception:
                angle_data[y].append([None])

    if not isinstance(plot_title, list):
        plot_title = [plot_title]
    if not isinstance(plot_x_label, list):
        plot_x_label = [plot_x_label]
    if not isinstance(plot_y_label, list):
        plot_y_label = [plot_y_label]

    # # set up inital values for the plot
    # nRows, nCols = find_row_col(len(out_chans[y]))
    # fig = plt.figure(figsize=[(2 + 10 * nCols), (2 + 4 * nRows)])

    subplots = len(out_chans[y])
    if subplots == 1:
        subplot_index = [111]
        nRows = 1
        nCols = 1
        # plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
        plt.figure(figsize=[(2 + 18 * nCols), (2 + 10 * nRows)])
        # plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98, left = 0.1, hspace = 0.25)
    if subplots == 2:
        subplot_index = [211, 212]
        nRows = 2
        nCols = 1
        # plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
        plt.figure(figsize=[(2 + 18 * nCols), (2 + 5 * nRows)])
        # plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98, left = 0.1, hspace = 0.28)
    if subplots == 3:
        subplot_index = [311, 312, 313]
        nRows = 3
        nCols = 1
        # plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
        plt.figure(figsize=[(2 + 24 * nCols), (2 + 5 * nRows)])
        # plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98,  left = 0.1, hspace = 0.32, wspace = 0.2)
    if subplots == 4:
        nRows = 2
        nCols = 2
        subplot_index = [221, 222, 223, 224]
        # plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
        plt.figure(figsize=[(2 + 12 * nCols), (2 + 6.5 * nRows)])
        # plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98, left = 0.1, hspace = 0.32, wspace = 0.2)
    if subplots == 5:
        subplot_index = [311, 323, 324, 325, 326]
        nRows = 3
        nCols = 2
        # plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
        plt.figure(figsize=[(2 + 12 * nCols), (2 + 5 * nRows)])
        # plt.subplots_adjust(bottom = 0.05, top = 0.91, right = 0.98, left = 0.1, hspace = 0.35, wspace = 0.2)
    if subplots == 6:
        nRows = 3
        nCols = 2
        subplot_index = [321, 322, 323, 324, 325, 326]
        # plt.figure(figsize = (29.7/ 2.54, 21.0/ 2.54))
        plt.figure(figsize=[(2 + 12 * nCols), (2 + 5 * nRows)])

    plt.suptitle(plot_title, fontsize=22, fontweight='bold')
    plt.rc('xtick', labelsize=18)
    plt.rc('ytick', labelsize=18)

    # pre-initialize
    ax = []
    time_diff = 0
    num = len(data[0])
    for sub in xrange(num):
        # If no label given, default to blank labels to prevent error being thrown
        if sub > len(plot_title) - 1:
            plot_title.append("")
        if sub > len(plot_x_label) - 1:
            plot_x_label.append("Time (s)")
        if sub > len(plot_y_label) - 1:
            plot_y_label.append("")

        # The following code is plotting and plot options

        # Set it up so that all plots share the same x axis, this means that you can zoom and navigate the plot...
        #       and all the subplots will move together
        if xlims == "" or isinstance(xlims[0], int) or isinstance(xlims[0], float):
            if sub == 0:
                ax.append(plt.subplot(nRows, nCols, sub + 1))
            else:
                ax.append(plt.subplot(nRows, nCols, sub + 1, sharex=ax[0]))
        else:
            ax.append(plt.subplot(nRows, nCols, sub + 1))

        # Titles and labels
        if not isinstance(out_chans[0][sub], list):
            if plot_title[sub] == "NA":
                plt.title(_ch_id[0][out_chans[0][sub]], fontsize=18)
            else:
                plt.title(plot_title[sub], fontsize=18)
        elif plot_title[sub] == "NA":
            plt.title("", fontsize=18)
        else:
            plt.title(plot_title[sub], fontsize=18)
        plt.xlabel(plot_x_label[sub], fontsize=18)
        plt.ylabel(plot_y_label[sub], fontsize=18, labelpad=20)

        plt.suptitle(plot_title, fontsize=22, fontweight='bold')
        plt.rc('xtick', labelsize = 18)
        plt.rc('ytick', labelsize = 18)
        plt.grid(1)

        data_lims = []
        # If out data to plot, plot it
        ymin = data[0][sub][0][0]
        ymax = data[0][sub][0][0]
        for y in xrange(len(out_file)):
            if "_PSSE" in out_file[y]:
                psse_pscad = " Psse"
            else:
                psse_pscad = " Pscad"
            if isinstance(out_chans[0][sub], list):  
                for x in xrange(len(out_chans[0][sub])):
                    if type(out_chans[0][sub][x]) is str:
                        for key in _ch_id[y]:
                            if out_chans[0][sub][x] == _ch_id[y][key]:
                                out_chans[0][sub][x] = key
                                break
                    if data[y][sub][x] != -1:
                        '''Ts = 0.02 #from the start to the finish of the slope, 20 milliseconds                        
                        simt = 0
                        tDataArrayIndex_start = 0
                        tDataArrayIndex_stop = 0
                        data_110 = []
                        data_90 = []
                        time_bounce = []
                        while tDataArrayIndex_stop < len(time[y]):
                            simt = time[y][tDataArrayIndex_stop]-time[y][tDataArrayIndex_start]
                            while simt < Ts:
                                simt = time[y][tDataArrayIndex_stop]-time[y][tDataArrayIndex_start]
                                tDataArrayIndex_stop += 1
                                
                            #time[y][tDataArrayIndex_start]
                            #time[y][tDataArrayIndex_stop]
                            #tDataArrayIndex_start
                            #tDataArrayIndex_stop
                            #data[y][sub][x][tDataArrayIndex_start]
                            #data[y][sub][x][tDataArrayIndex_stop]
                            data_110.append(data[y][sub][x][tDataArrayIndex_start] + 0.1*(data[y][sub][x][tDataArrayIndex_stop]-data[y][sub][x][tDataArrayIndex_start]))
                            data_90.append(data[y][sub][x][tDataArrayIndex_start] - 0.1*(data[y][sub][x][tDataArrayIndex_stop]-data[y][sub][x][tDataArrayIndex_start]))
                            time_bounce.append(time[y][tDataArrayIndex_start])
                            tDataArrayIndex_start += 1
                            tDataArrayIndex_stop += 1'''

                        if xlims == "" or len(xlims) < sub + 1 or xlims[sub] == "":
                            xlimit = [min(min(time[0]), min(time[1])), max(max(time[0]), max(time[1]))]
                        else:
                            xlimit = xlims[sub]

                        t_start = step_boundary[0]
                        t_stop =  step_boundary[1]

                        t_start = xlimit[0]
                        t_stop =  xlimit[1]

                        tDataArrayIndex = 0
                        simt = time[y][tDataArrayIndex]
                        while simt < t_start:
                            simt = time[y][tDataArrayIndex]
                            tDataArrayIndex += 1
                        tDataArrayIndex_start = tDataArrayIndex
                        while simt < t_stop:
                            simt = time[y][tDataArrayIndex]
                            if data[y][sub][x][tDataArrayIndex] < ymin:
                                ymin = data[y][sub][x][tDataArrayIndex]
                            elif data[y][sub][x][tDataArrayIndex] > ymax:
                                ymax = data[y][sub][x][tDataArrayIndex]
                            tDataArrayIndex += 1
                        tDataArrayIndex_stop = tDataArrayIndex

                        data_110 = []
                        data_90 = []
                        time_bounce = []
                        # y_dev = 0.1* (max(data[y][sub][x])-min(data[y][sub][x]))
                        # y_dev = 0.1* ((data[y][sub][x][tDataArrayIndex_stop])-(data[y][sub][x][tDataArrayIndex_start]))
                        y_dev = 0.1 * (ymax - ymin)
                        for index in range(len(time[y])):
                            data_110.append(data[y][sub][x][index]+y_dev)
                            data_90.append(data[y][sub][x][index]-y_dev)
                            time_bounce.append(time[y][index])

                        # CUSTOM
                        if angle[sub] != 0:
                            try:
                                for d in xrange(len(data[y][sub][x])):
                                    data[y][sub][x][d] -= angle_data[y][sub][d]
                            except Exception:
                                print("ERROR -=-=-=-==-=-=-=-=-=-=------------=-=-==-=-=--==-=-=-=-=-=-=-=-=-")
                                # /CUSTOM
                        if psse_pscad == " Psse":
                            plt.plot(time[y], data[y][sub][x], label=_ch_id[y][out_chans[0][sub][x]] + psse_pscad,linewidth = 3.5)
                            plt.plot(time_bounce,data_90,'--',label="90% Boundary",color='black',linewidth = 1.5)
                            plt.plot(time_bounce,data_110,'--',label="110% Boundary",color='black',linewidth = 1.5)
                        else:
                            plt.plot(time[y], data[y][sub][x], linestyle='dashed', dashes=(4, 3),
                                     label=_ch_id[y][out_chans[1][sub][x]] + psse_pscad,linewidth = 3.5)
                            # plt.plot(time[y], data[y][sub][x],label=_ch_id[y][out_chans[1][sub][x]] + psse_pscad,linewidth = 3.5)
                            # plt.plot(time_bounce,data_90,':',label="90% Boundary",color='gray')
                            # plt.plot(time_bounce,data_110,':',label="110% Boundary",color='gray')
                        data_lims += data[y][sub][x]
            else:
                if type(out_chans[0][sub]) is str:
                    for key in _ch_id[y]:
                        if out_chans[0][sub] == _ch_id[y][key]:
                            out_chans[0][sub] = key
                            break
                if data[y][sub][0] != -1:
                    if angle[sub] != 0:
                        try:
                            for d in xrange(len(data[y][sub][0])):
                                data[y][sub][0][d] -= angle_data[y][sub][d]
                        except Exception:
                            print("ERROR -=-=-=-==-=-=-=-=-=-=------------=-=-==-=-=--==-=-=-=-=-=-=-=-=-")
                    # plt.plot(time, data[y][sub][0], label=_ch_id[y][out_chans[sub]] + psse_pscad)
                    if psse_pscad == " Psse":
                        plt.plot(time, data[y][sub][x], label=_ch_id[y][out_chans[0][sub][x]] + psse_pscad, linewidth = 3.5)
                        plt.plot(time[y],np.multiply(data[y][sub][x],0.9),'--',label="90% Boundary",linewidth = 1.5)
                        plt.plot(time[y],np.multiply(data[y][sub][x],1.1),'--',label="110% Boundary",linewidth = 1.5)
                    else:
                        plt.plot(time, data[y][sub][x], linestyle="dashed", dashes=(4, 3),
                                 label=_ch_id[y][out_chans[1][sub][x]] + psse_pscad, linewidth = 3.5)
                    data_lims += data[y][sub][0]

        plt.grid(True)  # turn on grid

        # Axis settings
        ax[sub].xaxis.set_minor_locator(ticker.AutoMinorLocator(2))  # automatically put a gridline in between tickers
        ax[sub].grid(b=True, which='minor', linestyle=':')  # gridline style is dots
        ax[sub].tick_params(axis='x', which='minor', bottom='on')  # display small ticks at bottom midway
        ax[sub].get_xaxis().get_major_formatter().set_useOffset(False)  # turn off the annoying offset values
        ax[sub].get_yaxis().get_major_formatter().set_useOffset(False)
        plt.tight_layout()
        if xlims == "" or len(xlims) < sub + 1 or xlims[sub] == "":
            xlimit = [min(min(time[0]),min(time[1])), max(max(time[0]),max(time[1]))]
            ax[sub].set_xlim(xlimit)
        else:
            xlimit = xlims[sub]
            ax[sub].set_xlim(xlimit)

        # Make sure that there is 15% padding on top and bottom
        max_data = max(data_lims)
        min_data = min(data_lims)
        max_data = ymax
        min_data = ymin

        data_range = max_data - min_data
        if ylims == "" or len(ylims) < sub + 1 or ylims[sub] == "":
            ylimit = ax[sub].get_ylim()
            ylimit = [ymin,ymax]
            new_y = restrict_y_axis(ylimit, data_range, max_data, min_data)
            ax[sub].set_ylim(new_y)
        else:
            # ylimit = ylims[sub]
            # new_y = restrict_y_axis(ylimit, data_range, max_data, min_data)
            # ax[sub].set_ylim(new_y)
            ax[sub].set_ylim(ylims[sub])

        # If there is a top title, make sure there is room for it
        top = 0.9
        if top_title == "":
            top = 0.95
        plt.subplots_adjust(right=0.9, hspace=0.3, wspace=0.3, top=top, left=0.1, bottom=0.12)
        # legend = plt.legend(bbox_to_anchor=(0., -0.28, 1., .105), loc=3,
        #                     ncol=2, mode="expand", prop={'size': 16}, borderaxespad=0.)
        legend = plt.legend(bbox_to_anchor=(0., -0.25, 1., .105), loc=2,
                            ncol=2, mode="expand", prop={'size': 16}, borderaxespad=0.1, frameon=False)

        legend.get_title().set_fontsize('16')  # legend 'Title' fontsize
        plt.setp(plt.gca().get_legend().get_texts(), fontsize='16')  # legend 'list' fontsize
        plt.subplots_adjust(hspace=0.5)
        # plt.legend(bbox_to_anchor=(0., -0.55, 1., .105), loc=3,
        #        ncol=4, mode="expand", borderaxespad=0.)

    # Plot the legend on the bottom, underneath everything
    # plt.legend(bbox_to_anchor=(0., -0.55, 1., .105), loc=3,
    #            ncol=4, mode="expand", borderaxespad=0.)

    # Save file

    file_str = os.path.join(out_folder, out_file[0]).replace('_PSSE.out', '.pdf')
    #file_str = file_str.split(".png", 1)
    #file_str = file_str[0] + ".png"
    plot_name = os.path.basename(file_str[0:-4])
    plt.suptitle(plot_name,fontsize = 20, fontweight = 'bold')
    plt.savefig(file_str, dpi=100)
    print(str(plot_name + "Saved"))
    plt.close()
def plot_all_out_files(folder=[''], chans=None,chans_adj=None, ylabel=[""] * 12, xlabel=["Time (s)"] * 12, title=["NA"] * 12, xlims="",
                       ylims="", top_title="", comp_angle_chan=[0, 0, 0, 0, 0, 0, 0, 0, 0],xoffset="",step_boundary=""):
    # path = folder + "/*.out"
    #print(str(glob.glob(folder)))
    for pssefile in glob.glob(folder + "/*_PSSE.out"):
        for pscadfile in glob.glob(folder + "/*_PSCAD.inf"):
            if pssefile[0:-9] == pscadfile[0:-10]:
                plot(folder, [pssefile, pscadfile], chans, chans_adj, plot_title=title, plot_x_label=xlabel, plot_y_label=ylabel,
                     top_title=top_title, xlims=xlims, ylims=ylims, angle=comp_angle_chan, xoffset = xoffset, step_boundary = step_boundary)
            else:
                pass
        #print pssefile + " done"




