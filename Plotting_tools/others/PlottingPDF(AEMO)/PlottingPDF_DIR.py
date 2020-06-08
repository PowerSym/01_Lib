import time
import os,sys
import shutil
timestr = time.strftime("%Y%m%d_%H%M%S")
cwd = os.getcwd()

dir_to_plot = 'C:\Grid Connection\Beryl Solar Farm\Power Systems studies\PSSE\Benchmark studies\Results\Frequency_Disturbance_Test'
plot_style = "PlottingPDF_Gen_VPQ.csv"
copy_dir = os.path.dirname(dir_to_plot)

def plot_all_in_dir(dir_to_plot, plot_style):
    PlotScript = "\""+ cwd +'\\'+ "PlottingPDF.py" +"\""
    PlotConfig = "\""+ cwd +'\\'+ plot_style +"\""
    pathPlotResults = dir_to_plot
    for file in os.listdir(dir_to_plot):
        if file.endswith(".pdf"):
            os.remove(os.path.join(pathPlotResults, file))
    pathPlotResults = "\""+dir_to_plot+ "\""
    os.system("C:/Python27/python "+PlotScript+" "+PlotConfig+" "+pathPlotResults)
    pathPlotResults = dir_to_plot
    for file in os.listdir(dir_to_plot):
        if (not file.endswith(".pdf")):
            continue    
        dst_file = os.path.join(pathPlotResults, file)    
        new_dst_file = pathPlotResults+'\\'+ dir_to_plot.split(os.sep)[-1] +'_'+os.path.splitext(os.path.basename(file))[0] +'_'+timestr+os.path.splitext(os.path.basename(file))[1]
        os.rename(dst_file, new_dst_file)
        shutil.copy(new_dst_file,copy_dir)
        
plot_all_in_dir(dir_to_plot,plot_style)

