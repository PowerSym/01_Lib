import sys
import os
# sys.path.append('C:\Program Files (x86)\PTI\PSSE34\PSSPY34')
# sys.path.append('C:\Program Files (x86)\PTI\PSSE34\PSSBIN')
# sys.path.append('C:\Program Files (x86)\PTI\PSSE34\PSSPY27')
import psse34
import psspy
import dyntools
WorkingFolder = os.getcwd()
def main():
    signals = []  # will print all signals, else you can indicate them like so   signals = [1,2,4,5]
    for outfile in os.listdir(WorkingFolder):
        if outfile.endswith(".out"):
            excel_file = outfile.split(".out")[0] + "_excel" # This line was added to show that xlsfile can be used to change name or even directory
            dyntools.CHNF(outfile).xlsout(channels=signals, show=False, outfile=outfile, xlsfile=excel_file)
if __name__ == '__main__':
    main()
