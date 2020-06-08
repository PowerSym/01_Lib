import ctypes
import os
import glob
import psse34
import datetime
import sys
import numpy
from matplotlib import pyplot as plt
from matplotlib import ticker

# Ensure that the user is running v2.5 of python

# Defines the path of the PSSE/psspy library
PSSE_PATH = r'c:/program files (x86)/pti/psse32/pssbin'
sys.path.append(PSSE_PATH)

# Tells PSSE where it is itself. Will eventually cause error without
os.environ['PATH'] += ';' + PSSE_PATH

# Imports the psspy library
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
    print "ERROR: UNREADABLE OUT FILE"
    fig = plt.figure()
    ax = fig.add_subplot(111)
    ax.text(5, 5, error_message, verticalalignment='center', horizontalalignment='center', fontsize=12)
    ax.axis([0, 10, 0, 10])
    out_file = "_UNREADABLE_OUT_FILE_" + out_file

    file_str = os.path.join(out_folder, out_file).replace('.out', '.png')

    plt.savefig(file_str, dpi=100)
    print str(file_str + "Saved")
    plt.close()
    return

# Documentation DONE
def plot(out_folder, out_file, out_chans, filename, plot_title=["NA"] * 12, plot_x_label=["Time (s)"] * 12,
         plot_y_label=[""] * 12,
         top_title="", xlims="", ylims="", angle=[0, 0, 0, 0, 0, 0, 0, 0, 0, 0]):
    '''if isinstance(out_file, basestring):
        out_file = [out_file]
    elif "_pre" in out_file[1].lower():
        out_file[0], out_file[1] = out_file[1], out_file[0]'''

    if out_chans is None:
        out_chans = []

    if not (ylims == "" or isinstance(ylims[0], list)):
        ylims = [ylims] * 12

    if not (xlims == "" or isinstance(xlims[0], list)):
        xlims = [xlims] * 12

    # for x in xrange(len(out_chans)):
    #     if not isinstance(out_chans, list):
    #         out_chans[x] = [out_chans[x]]

    _ch_id, _ch_data, _sh_ttl, time = [], [], [], []
    for file in xrange(len(out_file)):
        try:
            chnfobj = dyntools.CHNF(os.path.join(out_folder, out_file[file]))
            sh_ttl, ch_id, ch_data = chnfobj.get_data()
            _ch_data.append(ch_data)
            _ch_id.append(ch_id)
            _sh_ttl.append(sh_ttl)

            time.append(ch_data['time'])
        except OverflowError:
            error_plot(out_folder, out_file[file],
                       "The OUT file seems to have infinite data in it and cannot be read.\nThis case is considered a failed case.")
            return

    # loop through each of the channels required and extract the data
    data = []
    angle_data = []
    for y in xrange(len(out_file)):
        j = 0
        data.append([])
        for chan in out_chans[y]:
            if isinstance(chan, list):
                data[y].append([])
                for i in xrange(len(chan)):
                    if type(chan[i]) is str:
                        for key in _ch_id[y]:
                            if chan[i] == _ch_id[y][key]:
                                chan[i] = key
                                break
                    data[y][j].append(_ch_data[y][chan[i]])
            else:
                print chan
                data[y].append([])
                if type(chan) is str:
                    for key in _ch_id[y]:
                        if chan == _ch_id[y][key]:
                            chan = key
                            break
                print chan
                data[y][j].append(_ch_data[y][chan])

            j += 1
        for angle_chan in angle:
            angle_data.append([])
            try:
                angle_data[y].append(_ch_data[y][angle_chan])
            except Exception:
                angle_data[y].append([None])
        print 'Outfile channels %d: %s' %(y, _ch_id[y])

    if not isinstance(plot_title, list):
        plot_title = [plot_title]
    if not isinstance(plot_x_label, list):
        plot_x_label = [plot_x_label]
    if not isinstance(plot_y_label, list):
        plot_y_label = [plot_y_label]

    # set up inital values for the plot
    nRows, nCols = find_row_col(len(out_chans[y]))
    fig = plt.figure(figsize=[(2 + 6 * nCols), (2 + 4 * nRows)])
    fig.suptitle(top_title, fontsize=20)

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
        # if xlims == "" or isinstance(xlims[0], int) or isinstance(xlims[0], float):
        #     if sub == 0:
        #         ax.append(plt.subplot(nRows, nCols, sub + 1))
        #     else:
        #         ax.append(plt.subplot(nRows, nCols, sub + 1, sharex=ax[0]))
        # else:

        ax.append(plt.subplot(nRows, nCols, sub + 1))

        # Titles and labels
        if not isinstance(out_chans[0][sub], list):
            if plot_title[sub] == "NA":
                plt.title(_ch_id[0][out_chans[0][sub]], fontsize=13)
            else:
                plt.title(plot_title[sub], fontsize=13)
        elif plot_title[sub] == "NA":
            plt.title("", fontsize=13)
        else:
            plt.title(plot_title[sub], fontsize=13)
        plt.xlabel(plot_x_label[sub])
        plt.ylabel(plot_y_label[sub], labelpad=20)
        data_lims = []
        # If out data to plot, plot it
        for y in xrange(len(out_file)):
            if filename[0] in out_file[y]:
                prepost = filename[0]
            else:
                prepost = filename[1]
            if isinstance(out_chans[0][sub], list):
                for x in xrange(len(out_chans[0][sub])):
                    if type(out_chans[0][sub][x]) is str:
                        for key in _ch_id[y]:
                            if out_chans[0][sub][x] == _ch_id[y][key]:
                                out_chans[0][sub][x] = key
                                break
                    if data[y][sub][x] != -1:
                        t_start = 10
                        t_stop = 15
                        tDataArrayIndex = 0
                        simt = time[y][tDataArrayIndex]
                        while simt < t_start:
                            simt = time[y][tDataArrayIndex]
                            tDataArrayIndex += 1
                        tDataArrayIndex_start = tDataArrayIndex
                        while simt < t_stop:
                            simt = time[y][tDataArrayIndex]
                            tDataArrayIndex += 1
                        tDataArrayIndex_stop = tDataArrayIndex
                        data_110 = []
                        data_90 = []
                        time_bounce = []
                        y_dev = 0.1 * (max(data[y][sub][x]) - min(data[y][sub][x]))
                        y_dev = 0.1 * (
                                    (data[y][sub][x][tDataArrayIndex_stop]) - (data[y][sub][x][tDataArrayIndex_start]))
                        for index in range(len(time[y])):
                            data_110.append(data[y][sub][x][index] + y_dev)
                            data_90.append(data[y][sub][x][index] - y_dev)
                            time_bounce.append(time[y][index])

                        # CUSTOM
                        if angle[sub] != 0:
                            try:
                                for d in xrange(len(data[y][sub][x])):
                                    data[y][sub][x][d] -= angle_data[y][sub][d]
                            except Exception:
                                print("ERROR -=-=-=-==-=-=-=-=-=-=------------=-=-==-=-=--==-=-=-=-=-=-=-=-=-")
                                # /CUSTOM
                        if prepost == filename[0]:
                            plt.plot(time[y], data[y][sub][x], label=_ch_id[y][out_chans[0][sub][x]] + prepost)
                            # plt.plot(time_bounce,data_90,':',label="90% Boundary",color='gray')
                            # plt.plot(time_bounce,data_110,':',label="110% Boundary",color='gray')
                        else:                           plt.plot(time[y], data[y][sub][x], linestyle='dashed', dashes=(4, 3),
                                     label=_ch_id[y][out_chans[1][sub][x]] + prepost)
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
                    t_start = 5
                    t_stop = 30

                    tDataArrayIndex = 0
                    simt = time[y][tDataArrayIndex]
                    while simt < t_start:
                        simt = time[y][tDataArrayIndex]
                        tDataArrayIndex += 1
                    tDataArrayIndex_start = tDataArrayIndex
                    while simt < t_stop:
                        simt = time[y][tDataArrayIndex]
                        tDataArrayIndex += 1
                    tDataArrayIndex_stop = tDataArrayIndex
                    data_110 = []
                    data_90 = []
                    time_bounce = []
                    y_dev = 0.1 * (max(data[y][sub][x]) - min(data[y][sub][x]))
                    y_dev = 0.1 * ((data[y][sub][x][tDataArrayIndex_stop]) - (data[y][sub][x][tDataArrayIndex_start]))
                    for index in range(len(time[y])):
                        data_110.append(data[y][sub][x][index] + y_dev)
                        data_90.append(data[y][sub][x][index] - y_dev)
                        time_bounce.append(time[y][index])


                    if angle[sub] != 0:
                        try:
                            for d in xrange(len(data[y][sub][0])):
                                data[y][sub][0][d] -= angle_data[y][sub][d]
                        except Exception:
                            print("ERROR -=-=-=-==-=-=-=-=-=-=------------=-=-==-=-=--==-=-=-=-=-=-=-=-=-")
                    # plt.plot(time, data[y][sub][0], label=_ch_id[y][out_chans[sub]] + prepost)
                    if prepost == filename[0]:
                        plt.plot(time[y], data[y][sub][x], label=_ch_id[y][out_chans[0][sub][x]] + prepost)
                        # plt.plot(time_bounce,data_90,':',label="90% Boundary",color='gray')
                        # plt.plot(time_bounce,data_110,':',label="110% Boundary",color='gray')
                    else:
                        plt.plot(time[y], data[y][sub][x], linestyle='dashed', dashes=(4, 3),
                                 label=_ch_id[y][out_chans[1][sub][x]] + prepost)
                        # plt.plot(time_bounce,data_90,':',label="90% Boundary",color='gray')
                        # plt.plot(time_bounce,data_110,':',label="110% Boundary",color='gray')
                    data_lims += data[y][sub][x]

        plt.grid(True)  # turn on grid

        # Axis settings
        ax[sub].xaxis.set_minor_locator(ticker.AutoMinorLocator(2))  # automatically put a gridline in between tickers
        ax[sub].grid(b=True, which='minor', linestyle=':')  # gridline style is dots
        ax[sub].tick_params(axis='x', which='minor', bottom='on')  # display small ticks at bottom midway
        ax[sub].get_xaxis().get_major_formatter().set_useOffset(False)  # turn off the annoying offset values
        ax[sub].get_yaxis().get_major_formatter().set_useOffset(False)
        plt.tight_layout()

        # Make sure that there is 15% padding on top and bottom
        max_data = max(data_lims)
        min_data = min(data_lims)
        data_range = max_data - min_data
        if ylims == "" or len(ylims) < sub + 1 or ylims[sub] == "":
            ylimit = ax[sub].get_ylim()
            new_y = restrict_y_axis(ylimit, data_range, max_data, min_data)
            ax[sub].set_ylim(new_y)
        else:
            ax[sub].set_ylim(ylims[sub])

        if xlims == "" or len(xlims) < sub + 1 or xlims[sub] == "":
            xlimit = [0, (max(time[0]))]  #time = [[time]] 
            #xlimit = ax[sub].get_xlim()
            ax[sub].set_xlim(xlimit)
        else:
            ax[sub].set_xlim(xlims[sub])

        # If there is a top title, make sure there is room for it
        top = 0.9
        if top_title == "":
            top = 0.95
        plt.subplots_adjust(right=0.9, hspace=0.3, wspace=0.3, top=top, left=0.1, bottom=0.12)
        legend = plt.legend(bbox_to_anchor=(0., -0.28, 1., .105), loc=3,
                            ncol=2, mode="expand", prop={'size': 6}, borderaxespad=0.)
        legend.get_title().set_fontsize('6')  # legend 'Title' fontsize
        plt.setp(plt.gca().get_legend().get_texts(), fontsize='10')  # legend 'list' fontsize
        plt.subplots_adjust(hspace=0.5)
        # plt.legend(bbox_to_anchor=(0., -0.55, 1., .105), loc=3,
        #        ncol=4, mode="expand", borderaxespad=0.)

    # Plot the legend on the bottom, underneath everything
    # plt.legend(bbox_to_anchor=(0., -0.55, 1., .105), loc=3,
    #            ncol=4, mode="expand", borderaxespad=0.)

    # Save file

    #file_str = os.path.join(out_folder, out_file[0]).replace('%s.out'%(filename[0]), '.pdf')
    file_str = os.path.join(out_folder, out_file[0]).replace('.out', '_VS' + filename[1]+'.pdf')
    #file_str = file_str.split(".png", 1)
    #file_str = file_str[0] + ".png"
    plot_name = os.path.basename(file_str[0:-4])
    plt.suptitle(plot_name,fontsize = 12,fontweight='bold')
    plt.savefig(file_str, dpi=100)
    print str(file_str + "Saved")
    plt.close()



def plot_all_out_files(folder=[''], chans=None,filename =[],ylabel=[""] * 12, xlabel=["Time (s)"] * 12, title=["NA"] * 12, xlims="",
                       ylims="", top_title="", comp_angle_chan=[0, 0, 0, 0, 0, 0, 0, 0, 0]):
    # path = folder + "/*.out"
    #print str(glob.glob(folder))
    for filepre in glob.glob(folder + "/*"+ filename[0] + "*"):
        for filepost in glob.glob(folder + "/*"+ filename[1] + "*"):
            if filepre[0:filepre.find(filename[0])] == filepost[0:filepost.find(filename[1])]:
                plot(folder, [filepre, filepost], chans,filename, plot_title=title, plot_x_label=xlabel, plot_y_label=ylabel,
                     top_title=top_title, xlims=xlims, ylims=ylims, angle=comp_angle_chan)
        print filepre + " done"
        


