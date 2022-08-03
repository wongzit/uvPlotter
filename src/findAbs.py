import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.drawing.colors import ColorChoice

def readout(fName):
	with open(fName, 'r') as logFile:
		logLines = logFile.readlines()
	ener = []
	osStrength = []
	for logLine in logLines:
		if 'Excited State' in logLine:
			if float(logLine.split()[8][2:]) != 0.0:
				ener.append(float(logLine.split()[6]))
				osStrength.append(float(logLine.split()[8][2:]))
	return ener, osStrength

def calcUV(usr_sd, energies, os_st, osMin, osMax, usr_point):
	plt_range = np.linspace(osMin, osMax, usr_point)
	sum_ref = []
	for i in plt_range:
		sum_cur = 0.0
		for no in range(len(energies)):
			sum_cur += 1.3062974e8*os_st[no]/(1e7/(1239.84/usr_sd))*np.exp(-(((1/i-1/energies[no])/(1/(1239.84/usr_sd)))**2))
		sum_ref.append(sum_cur)
	return plt_range, sum_ref

def save_excel(fName2, usr_sd2, plt_min, plt_max):
	ener2, os2 = readout(fName2)
	x_range, abs_data = calcUV(usr_sd2, ener2, os2, plt_min, plt_max, 2000)

	wb = Workbook()
	ws = wb.active

	f_1st_row = ['Wavelength (nm)', 'Molar Extinction Coefficient']
	f_2nd_row = [x_range[0], abs_data[0]]
	f_3rd_row = [x_range[1], abs_data[1]]
	for no_ii in range(len(ener2)):
		f_1st_row.append('Wavelength')
		f_1st_row.append('Oscillator Strength')
		f_2nd_row.append(ener2[no_ii])
		f_2nd_row.append(0)
		f_3rd_row.append(ener2[no_ii])
		f_3rd_row.append(os2[no_ii])

	ws.append(f_1st_row)
	ws.append(f_2nd_row)
	ws.append(f_3rd_row)

	if len(x_range) > 2:
		for no in range(2, len(x_range)):
			ws.append([x_range[no], abs_data[no]])

	uv_chart = ScatterChart()
	x_value = Reference(ws, min_col = 1, min_row = 2, max_row = len(x_range) + 1)
	y_value = Reference(ws, min_col = 2, min_row = 2, max_row = len(x_range) + 1)
	series_abs = Series(y_value, x_value)
	uv_chart.series.append(series_abs)
	uv_chart.legend = None
	uv_chart.x_axis.title = 'Wavelength (nm)'
	uv_chart.y_axis.title = 'Molar Extinction Coefficient'

	os_chart = ScatterChart()
	for no_iii in range(len(ener2)):
		xf_value = Reference(ws, min_col = 3 + no_iii * 2, min_row = 2, max_row = 3)
		yf_value = Reference(ws, min_col = 4 + no_iii * 2, min_row = 2, max_row = 3)
		series_osf = Series(yf_value, xf_value)
		os_chart.series.append(series_osf)
		series_osf.spPr.ln.solidFill = ColorChoice(prstClr="dkRed")
	os_chart.legend = None
	os_chart.y_axis.axId = 200
	os_chart.y_axis.title = 'Oscillator Strength'

	os_chart.y_axis.crosses = "max"
	uv_chart += os_chart
	
	ws.add_chart(uv_chart, "C4")
	wb.save(f"{fName2[:-4]}.xlsx")

