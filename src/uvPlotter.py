from tkinter import *
from PIL import ImageTk, Image
from tkinter import messagebox
from tkinter import filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
import numpy as np
import sys
from matplotlib import cm
from matplotlib.ticker import LinearLocator
from mpl_toolkits.mplot3d import axes3d
import webbrowser
from findAbs import *
import time
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.drawing.colors import ColorChoice

def open_web():
    webbrowser.open('https://wongzit.github.io/program/uvplotter', new = 1)

# Open File Function
def open_file_1():
    global fileName_1, shortName_1, save_path
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_1.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_1.get())):
        if fileName_1.get()[charNo] == '/':
            slashNo = charNo
    shortName_1.set(fileName_1.get()[slashNo+1:-4])
    save_path = fileName_1.get()[:slashNo]

def open_file_2():
    global fileName_2, shortName_2
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_2.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_2.get())):
        if fileName_2.get()[charNo] == '/':
            slashNo = charNo
    shortName_2.set(fileName_2.get()[slashNo+1:-4])

def open_file_3():
    global fileName_3, shortName_3
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_3.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_3.get())):
        if fileName_3.get()[charNo] == '/':
            slashNo = charNo
    shortName_3.set(fileName_3.get()[slashNo+1:-4])

def open_file_4():
    global fileName_4, shortName_4
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_4.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_4.get())):
        if fileName_4.get()[charNo] == '/':
            slashNo = charNo
    shortName_4.set(fileName_4.get()[slashNo+1:-4])

def open_file_5():
    global fileName_5, shortName_5
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_5.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_5.get())):
        if fileName_5.get()[charNo] == '/':
            slashNo = charNo
    shortName_5.set(fileName_5.get()[slashNo+1:-4])

def open_file_1():
    global fileName_1, shortName_1
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_1.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_1.get())):
        if fileName_1.get()[charNo] == '/':
            slashNo = charNo
    shortName_1.set(fileName_1.get()[slashNo+1:-4])

def open_file_6():
    global fileName_6, shortName_6
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_6.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_6.get())):
        if fileName_6.get()[charNo] == '/':
            slashNo = charNo
    shortName_6.set(fileName_6.get()[slashNo+1:-4])

def open_file_7():
    global fileName_7, shortName_7
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_7.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_7.get())):
        if fileName_7.get()[charNo] == '/':
            slashNo = charNo
    shortName_7.set(fileName_7.get()[slashNo+1:-4])

def open_file_8():
    global fileName_8, shortName_8
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_8.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_8.get())):
        if fileName_8.get()[charNo] == '/':
            slashNo = charNo
    shortName_8.set(fileName_8.get()[slashNo+1:-4])

def open_file_9():
    global fileName_9, shortName_9
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_9.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_9.get())):
        if fileName_9.get()[charNo] == '/':
            slashNo = charNo
    shortName_9.set(fileName_9.get()[slashNo+1:-4])

def open_file_10():
    global fileName_10, shortName_10
    fTyp = [('All files', '*.*')]
    iDir = '~/Desktop'         #iDir = '/home/'
    logfilepath = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir, title = 'Choose file')
    fileName_10.set(logfilepath)
    slashNo = 0
    for charNo in range(len(fileName_10.get())):
        if fileName_10.get()[charNo] == '/':
            slashNo = charNo
    shortName_10.set(fileName_10.get()[slashNo+1:-4])

# Read files
def read_file():
    global fileName_1, fileName_2, fileName_3, fileName_4, fileName_5, fileName_6, fileName_7, fileName_8, fileName_9, fileName_10
    global fileNameList, shortNameList, enerList, os_strList, colorCurList, colorStiList, titleSize, tickSize, leSize, usr_alpha
    global usr_lam_max, usr_lam_min
    
    fileNameList = []
    shortNameList = []
    fileNameList.append(fileName_1.get())
    fileNameList.append(fileName_2.get())
    fileNameList.append(fileName_3.get())
    fileNameList.append(fileName_4.get())
    fileNameList.append(fileName_5.get())
    fileNameList.append(fileName_6.get())
    fileNameList.append(fileName_7.get())
    fileNameList.append(fileName_8.get())
    fileNameList.append(fileName_9.get())
    fileNameList.append(fileName_10.get())
    shortNameList.append(shortName_1.get())
    shortNameList.append(shortName_2.get())
    shortNameList.append(shortName_3.get())
    shortNameList.append(shortName_4.get())
    shortNameList.append(shortName_5.get())
    shortNameList.append(shortName_6.get())
    shortNameList.append(shortName_7.get())
    shortNameList.append(shortName_8.get())
    shortNameList.append(shortName_9.get())
    shortNameList.append(shortName_10.get())
    
    enerList = []
    os_strList = []
    for fileNo in range(len(fileNameList)):
        enerEle = []
        osEle = []
        if fileNameList[fileNo] != '':
            enerEle, osEle = readout(fileNameList[fileNo])
            enerList.append(enerEle)
            os_strList.append(osEle)

    if len(enerList) == 1:
        en_min = enerList[0][0]
        en_max = enerList[0][0]
        for k in enerList[0]:
            if k < en_min:
                en_min = int(k)
            elif k > en_max:
                en_max = int(k)
        
        fig, ax1 = plt.subplots()
        ax2 = ax1.twinx()

        for plotNo in range(len(enerList)):
            if usr_lam_max.get() == 'def' or usr_lam_min.get() == 'def':
                x_range, ref_val = calcUV(usrSD, enerList[plotNo], os_strList[plotNo], en_min-100, en_max+100, dataPoint)
            else:
                x_range, ref_val = calcUV(usrSD, enerList[plotNo], os_strList[plotNo], int(usr_lam_min.get()), int(usr_lam_max.get()), dataPoint)
            ax1.plot(x_range, ref_val, c = sin_cur_color.get())
            for no2 in range(len(os_strList[plotNo])):
                ax2.plot((enerList[plotNo][no2], enerList[plotNo][no2]), (0, os_strList[plotNo][no2]), c = sin_sti_color.get())
        if usr_lam_max.get() == 'def' or usr_lam_min.get() == 'def':
            plt.xlim([en_min-100, en_max+100])
        else:
            plt.xlim([int(usr_lam_min.get()), int(usr_lam_max.get())])
        ax1.set_xlabel('Wavelength (nm)', fontsize = titleSize)
        ax1.tick_params(labelsize = tickSize)
        ax2.tick_params(labelsize = tickSize)
        ax1.set_ylabel(r'$\epsilon\;\;\rm{(L\;mol^{-1}\;cm^{-1})}$', fontsize = titleSize)
        ax2.set_ylabel('Oscillator Strength', fontsize = titleSize)
        plt.show()
        fig.savefig(f'{fileNameList[0][:-4]}.png', format = "png", dpi = 300)

    elif len(enerList) == 0:
        messagebox.showwarning("WARNING", "No output file was found!")

    else:
        en_min = enerList[0][0]
        en_max = enerList[0][0]
        for i in enerList:
            for j in i:
                if j < en_min:
                    en_min = int(j)
                elif j > en_max:
                    en_max = int(j)

        fig, ax1 = plt.subplots()
        ax2 = ax1.twinx()

        for plotNo in range(len(enerList)):
            if usr_lam_max.get() == 'def' or usr_lam_min.get() == 'def':
                x_range, ref_val = calcUV(usrSD, enerList[plotNo], os_strList[plotNo], en_min-100, en_max+100, dataPoint)
            else:
                x_range, ref_val = calcUV(usrSD, enerList[plotNo], os_strList[plotNo], int(usr_lam_min.get()), int(usr_lam_max.get()), dataPoint)
            ax1.plot(x_range, ref_val, c = colorCurList[plotNo], label = shortNameList[plotNo])
            for no2 in range(len(os_strList[plotNo])):
                ax2.plot((enerList[plotNo][no2], enerList[plotNo][no2]), (0, os_strList[plotNo][no2]), c = colorStiList[plotNo], alpha = float(usr_alpha.get()))
        if usr_lam_max.get() == 'def' or usr_lam_min.get() == 'def':
            plt.xlim([en_min-100, en_max+100])
        else:
            plt.xlim([int(usr_lam_min.get()), int(usr_lam_max.get())])
        ax1.set_xlabel('Wavelength (nm)', fontsize = titleSize)
        ax1.tick_params(labelsize = tickSize)
        ax2.tick_params(labelsize = tickSize)
        ax1.set_ylabel(r'$\epsilon\;\;\rm{(L\;mol^{-1}\;cm^{-1})}$', fontsize = titleSize)
        ax2.set_ylabel('Oscillator Strength', fontsize = titleSize)
        ax1.legend(fontsize = leSize)
        plt.show()

        slashNo2 = 0
        for charNo2 in range(len(fileName_1.get())):
            if fileName_1.get()[charNo2] == '/':
                slashNo2 = charNo2
        
        fig.savefig(f'{fileName_1.get()[:slashNo2+1]}abs_spec_{int(time.time())}.png', format = "png", dpi = 300)

def setting_panel():
    global sin_cur_color, sin_sti_color
    setWin = Toplevel()
    setWin.title('Setting')
    
    singColor = LabelFrame(setWin, text = ' Single Spectrum Color ', font = ('Helvetica', 16, 'bold'), padx = 10, pady = 10)
    singColor.pack(padx = 10, pady = 10, fill = 'x')
    Label(singColor, text = '  Absorption: ').grid(row = 0, column = 0)
    Entry(singColor, textvariable = sin_cur_color, width = 8).grid(row = 0, column = 1)
    Label(singColor, text = '  Oscillator Strength: ').grid(row = 0, column = 2)
    Entry(singColor, textvariable = sin_sti_color, width = 8).grid(row = 0, column = 3)

    multiColor = LabelFrame(setWin, text = ' Multi-Spectrum Color ', font = ('Helvetica', 16, 'bold'), padx = 10, pady = 10)
    multiColor.pack(padx = 10, pady = 10, fill = 'x')
    Label(multiColor, text = '  Absorption ').grid(row = 0, column = 1)
    Label(multiColor, text = '  Oscillator Strength ').grid(row = 0, column = 2)
    Label(multiColor, text = '  (1)  ').grid(row = 1, column = 0)
    Label(multiColor, text = '  (2)  ').grid(row = 2, column = 0)
    Label(multiColor, text = '  (3)  ').grid(row = 3, column = 0)
    Label(multiColor, text = '  (4)  ').grid(row = 4, column = 0)
    Label(multiColor, text = '  (5)  ').grid(row = 5, column = 0)
    Label(multiColor, text = '  (6)  ').grid(row = 6, column = 0)
    Label(multiColor, text = '  (7)  ').grid(row = 7, column = 0)
    Label(multiColor, text = '  (8)  ').grid(row = 8, column = 0)
    Label(multiColor, text = '  (9)  ').grid(row = 9, column = 0)
    Label(multiColor, text = '  (10)  ').grid(row = 10, column = 0)
    global mc_c_1, mc_c_2, mc_c_3, mc_c_4, mc_c_5, mc_c_6, mc_c_7, mc_c_8, mc_c_9, mc_c_10
    global mc_s_1, mc_s_2, mc_s_3, mc_s_4, mc_s_5, mc_s_6, mc_s_7, mc_s_8, mc_s_9, mc_s_10
    Entry(multiColor, textvariable = mc_c_1, width = 8).grid(row = 1, column = 1)
    Entry(multiColor, textvariable = mc_s_1, width = 8).grid(row = 1, column = 2)
    Entry(multiColor, textvariable = mc_c_2, width = 8).grid(row = 2, column = 1)
    Entry(multiColor, textvariable = mc_s_2, width = 8).grid(row = 2, column = 2)
    Entry(multiColor, textvariable = mc_c_3, width = 8).grid(row = 3, column = 1)
    Entry(multiColor, textvariable = mc_s_3, width = 8).grid(row = 3, column = 2)
    Entry(multiColor, textvariable = mc_c_4, width = 8).grid(row = 4, column = 1)
    Entry(multiColor, textvariable = mc_s_4, width = 8).grid(row = 4, column = 2)
    Entry(multiColor, textvariable = mc_c_5, width = 8).grid(row = 5, column = 1)
    Entry(multiColor, textvariable = mc_s_5, width = 8).grid(row = 5, column = 2)
    Entry(multiColor, textvariable = mc_c_6, width = 8).grid(row = 6, column = 1)
    Entry(multiColor, textvariable = mc_s_6, width = 8).grid(row = 6, column = 2)
    Entry(multiColor, textvariable = mc_c_7, width = 8).grid(row = 7, column = 1)
    Entry(multiColor, textvariable = mc_s_7, width = 8).grid(row = 7, column = 2)
    Entry(multiColor, textvariable = mc_c_8, width = 8).grid(row = 8, column = 1)
    Entry(multiColor, textvariable = mc_s_8, width = 8).grid(row = 8, column = 2)
    Entry(multiColor, textvariable = mc_c_9, width = 8).grid(row = 9, column = 1)
    Entry(multiColor, textvariable = mc_s_9, width = 8).grid(row = 9, column = 2)
    Entry(multiColor, textvariable = mc_c_10, width = 8).grid(row = 10, column = 1)
    Entry(multiColor, textvariable = mc_s_10, width = 8).grid(row = 10, column = 2)

    global usr_def_sd, usr_data_point, title_size, tick_size, legend_size, usr_lam_min, usr_lam_max
    plotPara = LabelFrame(setWin, text = ' Plotting Parameters ', font = ('Helvetica', 16, 'bold'), padx = 10, pady = 10)
    plotPara.pack(padx = 10, pady = 10, fill = 'x')
    Label(plotPara, text = ' Standard deviation (eV): ').grid(row = 0, column = 0, columnspan = 2)
    Spinbox(plotPara, textvariable = usr_def_sd, from_ = 0.1, to = 1, increment = 0.1, format = '%1.1f', width = 5).grid(row = 0, column = 2)
    Label(plotPara, text = ' Number of data point: ').grid(row = 1, column = 0, columnspan = 2)
    Entry(plotPara, textvariable = usr_data_point, width = 5).grid(row = 1, column = 2)
    Label(plotPara, text = 'Title').grid(row = 2, column = 1)
    Label(plotPara, text = 'Tick').grid(row = 2, column = 2)
    Label(plotPara, text = 'Legend').grid(row = 2, column = 3)
    Label(plotPara, text = 'Font size').grid(row = 3, column = 0)
    Spinbox(plotPara, textvariable = title_size, from_ = 1, to = 100, increment = 1, width = 5).grid(row = 3, column = 1)
    Spinbox(plotPara, textvariable = tick_size, from_ = 1, to = 100, increment = 1, width = 5).grid(row = 3, column = 2)
    Spinbox(plotPara, textvariable = legend_size, from_ = 1, to = 100, increment = 1, width = 5).grid(row = 3, column = 3)
    Label(plotPara, text = 'Transparency of oscillator stick (0.0 ~ 1.0):').grid(row = 4, column = 0, columnspan = 3)
    Spinbox(plotPara, textvariable = usr_alpha, from_ = 0.0, to = 1.0, increment = 0.1, format = '%1.1f', width = 5).grid(row = 4, column = 3)
    Label(plotPara, text = 'Wavelength range:').grid(row = 5, column = 0)
    Entry(plotPara, textvariable = usr_lam_min, width = 5).grid(row = 5, column = 1)
    Label(plotPara, text = ' nm      ~ ').grid(row = 5, column = 2)
    Entry(plotPara, textvariable = usr_lam_max, width = 5).grid(row = 5, column = 3)
    Label(plotPara, text = ' nm').grid(row = 5, column = 4)

    def set_ok():
        global colorCurList, colorStiList, usrSD, dataPoint, titleSize, tickSize, leSize
        colorCurList = [mc_c_1.get(), mc_c_2.get(), mc_c_3.get(), mc_c_4.get(), mc_c_5.get(), \
            mc_c_6.get(), mc_c_7.get(), mc_c_8.get(), mc_c_9.get(), mc_c_10.get()]
        colorStiList = [mc_s_1.get(), mc_s_2.get(), mc_s_3.get(), mc_s_4.get(), mc_s_5.get(), \
            mc_s_6.get(), mc_s_7.get(), mc_s_8.get(), mc_s_9.get(), mc_s_10.get()]
        usrSD = float(usr_def_sd.get())
        dataPoint = int(usr_data_point.get())
        titleSize = int(title_size.get())
        tickSize = int(tick_size.get())
        leSize = int(legend_size.get())

        setWin.destroy()

    ok_btn = Button(setWin, text = 'Ok', command = set_ok).pack()

def save_txt():
    global fileNameList, enerList, os_strList, usrSD, usr_lam_max, usr_lam_min
    for txtNo in range(len(enerList)):
        txtFile = open(f'{fileNameList[txtNo][:-3]}txt', 'w')
        txtFile.write('-------- Molar Extinction Coefficient --------\n')
        if usr_lam_max.get() == 'def' or usr_lam_min.get() == 'def':
            x_range_txt, ref_val_txt = calcUV(usrSD, enerList[txtNo], os_strList[txtNo], 1, 1000, 2000)
        else:
            x_range_txt, ref_val_txt = calcUV(usrSD, enerList[txtNo], os_strList[txtNo], int(usr_lam_min.get()), int(usr_lam_max.get()), 2000)
        for no_i in range(len(x_range_txt)):
            txtFile.write(f'{x_range_txt[no_i]}      {ref_val_txt[no_i]}\n')
        txtFile.write('\n\n-------- Oscillator Strength --------\n')
        for no_j in range(len(os_strList[txtNo])):
            txtFile.write(f'{enerList[txtNo][no_j]}      {os_strList[txtNo][no_j]}\n')
    messagebox.showwarning("Finished", ".txt files have been saved.")

def save_xlsx():
    global usr_lam_max, usr_lam_min
    for fiNa in fileNameList:
        if fiNa != '':
            if usr_lam_max.get() == 'def' or usr_lam_min.get() == 'def':
                save_excel(fiNa, usrSD, 1, 1000)
            else:
                save_excel(fiNa, usrSD, int(usr_lam_min.get()), int(usr_lam_max.get()))
    messagebox.showwarning("Finished", ".xlsx files have been saved.")

proVer = '1.0.0'
rlsDate = '2022-08-02'

root = Tk()
root.title(f'uv.Plotter v{proVer}')

wideLogo = ImageTk.PhotoImage(Image.open('assets/uvPlotter_main.png'))
Label(image = wideLogo).grid(row = 0, column = 0, columnspan = 4)

proInfoFrame = LabelFrame(root, borderwidth = 0, highlightthickness = 0)
proInfoFrame.grid(row = 1, column = 0, columnspan = 4, sticky = W + E)
Label(proInfoFrame, text = f'\nVer. {proVer} ({rlsDate})', font = ('Helvetica', 16, 'bold')).pack()
Label(proInfoFrame, text = 'Plot UV-vis spectrum from Gaussian outputs, and save as .txt/.xlsx files.').pack()
Label(proInfoFrame, text = 'Copyright © 2022 Zhe Wang (https://wongzit.github.io)\n\n').pack()

Label(root, text = 'Output', font = ('Helvetica', 13, 'bold')).grid(row = 2, column = 0)
Label(root, text = ' (1) ').grid(row = 3, column = 0)
Label(root, text = ' (2) ').grid(row = 4, column = 0)
Label(root, text = ' (3) ').grid(row = 5, column = 0)
Label(root, text = ' (4) ').grid(row = 6, column = 0)
Label(root, text = ' (5) ').grid(row = 7, column = 0)
Label(root, text = ' (6) ').grid(row = 8, column = 0)
Label(root, text = ' (7) ').grid(row = 9, column = 0)
Label(root, text = ' (8) ').grid(row = 10, column = 0)
Label(root, text = ' (9) ').grid(row = 11, column = 0)
Label(root, text = ' (10)').grid(row = 12, column = 0)

Label(root, text = 'Label', font = ('Helvetica', 13, 'bold')).grid(row = 2, column = 1)
shortName_1 = StringVar()
shortName_1.set('')
Entry(root, textvariable = shortName_1, width = 12).grid(row = 3, column = 1)
shortName_2 = StringVar()
shortName_2.set('')
Entry(root, textvariable = shortName_2, width = 12).grid(row = 4, column = 1)
shortName_3 = StringVar()
shortName_3.set('')
Entry(root, textvariable = shortName_3, width = 12).grid(row = 5, column = 1)
shortName_4 = StringVar()
shortName_4.set('')
Entry(root, textvariable = shortName_4, width = 12).grid(row = 6, column = 1)
shortName_5 = StringVar()
shortName_5.set('')
Entry(root, textvariable = shortName_5, width = 12).grid(row = 7, column = 1)
shortName_6 = StringVar()
shortName_6.set('')
Entry(root, textvariable = shortName_6, width = 12).grid(row = 8, column = 1)
shortName_7 = StringVar()
shortName_7.set('')
Entry(root, textvariable = shortName_7, width = 12).grid(row = 9, column = 1)
shortName_8 = StringVar()
shortName_8.set('')
Entry(root, textvariable = shortName_8, width = 12).grid(row = 10, column = 1)
shortName_9 = StringVar()
shortName_9.set('')
Entry(root, textvariable = shortName_9, width = 12).grid(row = 11, column = 1)
shortName_10 = StringVar()
shortName_10.set('')
Entry(root, textvariable = shortName_10, width = 12).grid(row = 12, column = 1)

Label(root, text = 'File Path', font = ('Helvetica', 13, 'bold')).grid(row = 2, column = 2)
fileName_1 = StringVar()
fileName_1.set('')
Entry(root, textvariable = fileName_1, width = 30).grid(row = 3, column = 2)
fileName_2 = StringVar()
fileName_2.set('')
Entry(root, textvariable = fileName_2, width = 30).grid(row = 4, column = 2)
fileName_3 = StringVar()
fileName_3.set('')
Entry(root, textvariable = fileName_3, width = 30).grid(row = 5, column = 2)
fileName_4 = StringVar()
fileName_4.set('')
Entry(root, textvariable = fileName_4, width = 30).grid(row = 6, column = 2)
fileName_5 = StringVar()
fileName_5.set('')
Entry(root, textvariable = fileName_5, width = 30).grid(row = 7, column = 2)
fileName_6 = StringVar()
fileName_6.set('')
Entry(root, textvariable = fileName_6, width = 30).grid(row = 8, column = 2)
fileName_7 = StringVar()
fileName_7.set('')
Entry(root, textvariable = fileName_7, width = 30).grid(row = 9, column = 2)
fileName_8 = StringVar()
fileName_8.set('')
Entry(root, textvariable = fileName_8, width = 30).grid(row = 10, column = 2)
fileName_9 = StringVar()
fileName_9.set('')
Entry(root, textvariable = fileName_9, width = 30).grid(row = 11, column = 2)
fileName_10 = StringVar()
fileName_10.set('')
Entry(root, textvariable = fileName_10, width = 30).grid(row = 12, column = 2)

Button(root, text = 'Open', command = open_file_1).grid(row = 3, column = 3)
Button(root, text = 'Open', command = open_file_2).grid(row = 4, column = 3)
Button(root, text = 'Open', command = open_file_3).grid(row = 5, column = 3)
Button(root, text = 'Open', command = open_file_4).grid(row = 6, column = 3)
Button(root, text = 'Open', command = open_file_5).grid(row = 7, column = 3)
Button(root, text = 'Open', command = open_file_6).grid(row = 8, column = 3)
Button(root, text = 'Open', command = open_file_7).grid(row = 9, column = 3)
Button(root, text = 'Open', command = open_file_8).grid(row = 10, column = 3)
Button(root, text = 'Open', command = open_file_9).grid(row = 11, column = 3)
Button(root, text = 'Open', command = open_file_10).grid(row = 12, column = 3)

readIcon = PhotoImage(file = r'assets/read_icon.png')
setIcon = PhotoImage(file = r'assets/set_icon.png')
webIcon = PhotoImage(file = r'assets/web_icon.png')
quitIcon = PhotoImage(file = r'assets/quit_icon.png')
xlsxIcon = PhotoImage(file = r'assets/excel_icon.png')
txtIcon = PhotoImage(file = r'assets/txt_icon.png')

btnFrame = LabelFrame(root, borderwidth = 0, highlightthickness = 0)
btnFrame.grid(row = 13, column = 0, columnspan = 4, sticky = W + E)

fileNameList = []
shortNameList = []
enerList = []
os_strList = []

sin_cur_color = StringVar()
sin_cur_color.set('darkblue')
sin_sti_color = StringVar()
sin_sti_color.set('darkred')

colorCurList = ['navy', 'darkred', 'darkgreen', 'steelblue', 'grey', 'teal', 'salmon', 'peru', 'olivedrab', 'sandybrown']
colorStiList = ['navy', 'darkred', 'darkgreen', 'steelblue', 'grey', 'teal', 'salmon', 'peru', 'olivedrab', 'sandybrown']

mc_c_1 = StringVar()
mc_c_1.set(colorCurList[0])
mc_s_1 = StringVar()
mc_s_1.set(colorStiList[0])
mc_c_2 = StringVar()
mc_c_2.set(colorCurList[1])
mc_s_2 = StringVar()
mc_s_2.set(colorStiList[1])
mc_c_3 = StringVar()
mc_c_3.set(colorCurList[2])
mc_s_3 = StringVar()
mc_s_3.set(colorStiList[2])
mc_c_4 = StringVar()
mc_c_4.set(colorCurList[3])
mc_s_4 = StringVar()
mc_s_4.set(colorStiList[3])
mc_c_5 = StringVar()
mc_c_5.set(colorCurList[4])
mc_s_5 = StringVar()
mc_s_5.set(colorStiList[4])
mc_c_6 = StringVar()
mc_c_6.set(colorCurList[5])
mc_s_6 = StringVar()
mc_s_6.set(colorStiList[5])
mc_c_7 = StringVar()
mc_c_7.set(colorCurList[6])
mc_s_7 = StringVar()
mc_s_7.set(colorStiList[6])
mc_c_8 = StringVar()
mc_c_8.set(colorCurList[7])
mc_s_8 = StringVar()
mc_s_8.set(colorStiList[7])
mc_c_9 = StringVar()
mc_c_9.set(colorCurList[8])
mc_s_9 = StringVar()
mc_s_9.set(colorStiList[8])
mc_c_10 = StringVar()
mc_c_10.set(colorCurList[9])
mc_s_10 = StringVar()
mc_s_10.set(colorStiList[9])

usr_def_sd = StringVar()
usr_def_sd.set('0.4')
usr_data_point = StringVar()
usr_data_point.set('20000')

usrSD = 0.4
dataPoint = 20000

# Label Size
title_size = StringVar()
tick_size = StringVar()
legend_size = StringVar()
title_size.set('15')
tick_size.set('13')
legend_size.set('13')

titleSize = 15
tickSize = 13
leSize = 13

usr_alpha = StringVar()
usr_alpha.set('0.6')

usr_lam_min = StringVar()
usr_lam_min.set('def')
usr_lam_max = StringVar()
usr_lam_max.set('def')

Label(btnFrame, text = '\n').grid(row = 0, column = 0)
Button(btnFrame, image = readIcon, padx = 20, pady = 6, command = read_file).grid(row = 1, column = 0)
Button(btnFrame, image = txtIcon, padx = 20, pady = 6, command = save_txt).grid(row = 1, column = 1)
Button(btnFrame, image = xlsxIcon, padx = 20, pady = 6, command = save_xlsx).grid(row = 1, column = 2)
Button(btnFrame, image = setIcon, padx = 20, pady = 6, command = setting_panel).grid(row = 1, column = 3)
Button(btnFrame, image = webIcon, padx = 20, pady = 6, command = open_web).grid(row = 1, column = 4)
Button(btnFrame, image = quitIcon, padx = 20, pady = 8, command = sys.exit).grid(row = 1, column = 5)

Label(btnFrame, text = '       Plot       ', font = ('Helvetica', 13, 'bold')).grid(row = 2, column = 0)
Label(btnFrame, text = '    Save .txt     ', font = ('Helvetica', 13, 'bold')).grid(row = 2, column = 1)
Label(btnFrame, text = '    Save .xlsx    ', font = ('Helvetica', 13, 'bold')).grid(row = 2, column = 2)
Label(btnFrame, text = '     Setting      ', font = ('Helvetica', 13, 'bold')).grid(row = 2, column = 3)
Label(btnFrame, text = '       Help       ', font = ('Helvetica', 13, 'bold')).grid(row = 2, column = 4)
Label(btnFrame, text = '       Quit       ', font = ('Helvetica', 13, 'bold')).grid(row = 2, column = 5)

root.mainloop()
