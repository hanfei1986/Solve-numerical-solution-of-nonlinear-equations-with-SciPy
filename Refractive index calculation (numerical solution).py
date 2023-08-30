import glob
import csv
import math
import numpy as np
from matplotlib import pyplot as plt
from scipy.optimize import fsolve
import xlwt

wafer_thickness = 0.05
angle_of_incidence = 6
spectrum_start = 630
spectrum_end = 440

theta = angle_of_incidence*math.pi/180

def read_data(data_file_path,start_point,end_point,reflectance_column,transmittance_column):
    with open(data_file_path, 'r') as csvfile:
        reader = csv.reader(csvfile)
        reader_rows = [row for row in reader]
    i = 0
    while True:
        if reader_rows[i][0] == str(start_point):
            row_start = i
        elif reader_rows[i][0] == str(end_point):
            row_end = i
            break
        i +=1
    wavelength = [float(reader_rows[i][0]) for i in range(row_start,row_end+1)]
    reflectance = [float(reader_rows[i][reflectance_column]) for i in range(row_start,row_end+1)]
    transmittance = [float(reader_rows[i][transmittance_column]) for i in range(row_start,row_end+1)]
    return wavelength,reflectance,transmittance

def data_match(old_wavelength,old_data,new_wavelength):
    new_data = []
    for i in range(0, len(old_wavelength)):
        for j in range(0, len(new_wavelength)):
            if old_wavelength[i] == new_wavelength[j]:
                new_data.append(old_data[i])
    return new_data

refractive_index_path = 'C:/Users/fhan/Desktop/Absorption Coefficient Calculation/Standard sample\Refractive index.csv'
with open(refractive_index_path, 'r') as csvfile:
    reader = csv.reader(csvfile)
    reader_rows = [row for row in reader]
i = 0
while True:
    if reader_rows[i][0] == '700':
        row_700 = i
    elif reader_rows[i][0] == '350':
        row_350 = i
        break
    i +=1
wavelength_refractive_index = [float(reader_rows[i][0]) for i in range(row_700,row_350+1)]
original_refractive_index = [float(reader_rows[i][1]) for i in range(row_700,row_350+1)]

workbook = xlwt.Workbook(encoding='utf-8')

production_wafer_path = 'C:/Users/fhan/Desktop/Absorption Coefficient Calculation/Raw data/Production wafer measurement\*.csv'
for filename in glob.glob(production_wafer_path):
    wafer_ID = str.upper(filename.split('\\')[1][0:6])+'-'+filename.split('\\')[1][7:9]
    filename_for_saving = filename.split('\\')[1].split('.')[0]
    
    sheet1 = workbook.add_sheet(filename_for_saving)
    sheet1.write(0,0,'Wavelength')
    sheet1.write(0,1,'Refractive index for S-polarized light')
    sheet1.write(0,2,'Refractive index for P-polarized light')
    sheet1_row = 1
       
    wavelength,reflectance,transmittance = read_data(filename,spectrum_start,spectrum_end,5,7)
    refractive_index = data_match(wavelength_refractive_index,original_refractive_index,wavelength)
    calculated_refractive_index = []
    solution_true_or_false = []
    for n_literature,R,T in zip(refractive_index,reflectance,transmittance):
        def func(n):
            d_eff = wafer_thickness/math.cos(theta/n)
            alpha_abs = -math.log((R+T)/100)/d_eff
            r = ((math.cos(theta)-n*(1-(math.sin(theta)/n)**2)**0.5)/(math.cos(theta)+n*(1-(math.sin(theta)/n)**2)**0.5))**2
            t = T/100
            alpha_trans = -math.log(((1-r)**4/(4*t**2*r**4)+1/r**2)**0.5-(1-r)**2/(2*t*r**2))/d_eff
            return alpha_abs-alpha_trans
        calculated_refractive_index_element = fsolve(func,n_literature)
        calculated_refractive_index.append(calculated_refractive_index_element[0])
        solution_true_or_false.append(np.isclose(func(calculated_refractive_index_element[0]),0))
    print(solution_true_or_false)
    plt.plot(wavelength,refractive_index,color='black',linewidth=1.0,label='reported in literature')
    plt.plot(wavelength,calculated_refractive_index,color='grey',linewidth=1.0,label='calculated from absorbance and transmittance')
    plt.title(wafer_ID+', S-polarized')
    plt.xlim(400,700)
    plt.ylim(2.62,2.74)
    plt.xlabel('wavelength (nm)')
    plt.ylabel('refractive index')
    plt.legend()
    plt.grid()
    plt.savefig('C:/Users/fhan/Desktop/Absorption Coefficient Calculation/Spectrum/'+filename_for_saving+', refractive index for S-polarized light.jpg')
    plt.show()
    
    for i in range(0,len(wavelength)):
        sheet1.write(sheet1_row,0,wavelength[i])
        sheet1.write(sheet1_row,1,calculated_refractive_index[i])
        sheet1_row += 1
    sheet1_row = 1
        
    wavelength,reflectance,transmittance = read_data(filename,spectrum_start,spectrum_end,11,9)
    refractive_index = data_match(wavelength_refractive_index,original_refractive_index,wavelength)
    calculated_refractive_index = []
    solution_true_or_false = []
    for n_literature,R,T in zip(refractive_index,reflectance,transmittance):
        def func(n):
            d_eff = wafer_thickness/math.cos(theta/n)
            alpha_abs = -math.log((R+T)/100)/d_eff
            r = (((1-(math.sin(theta)/n)**2)**0.5-n*math.cos(theta))/((1-(math.sin(theta)/n)**2)**0.5+n*math.cos(theta)))**2
            t = T/100
            alpha_trans = -math.log(((1-r)**4/(4*t**2*r**4)+1/r**2)**0.5-(1-r)**2/(2*t*r**2))/d_eff
            return alpha_abs-alpha_trans
        calculated_refractive_index_element = fsolve(func,n_literature)
        calculated_refractive_index.append(calculated_refractive_index_element[0])
        solution_true_or_false.append(np.isclose(func(calculated_refractive_index_element[0]),0))
    print(solution_true_or_false)
    plt.plot(wavelength,refractive_index,color='black',linewidth=1.0,label='reported in literature')
    plt.plot(wavelength,calculated_refractive_index,color='grey',linewidth=1.0,label='calculated from absorbance and transmittance')
    plt.title(wafer_ID+', P-polarized')
    plt.xlim(400,700)
    plt.ylim(2.62,2.74)
    plt.xlabel('wavelength (nm)')
    plt.ylabel('refractive index')
    plt.legend()
    plt.grid()
    plt.savefig('C:/Users/fhan/Desktop/Absorption Coefficient Calculation/Spectrum/'+filename_for_saving+', refractive index for P-polarized light.jpg')
    plt.show()
    
    for i in range(0,len(wavelength)):
        sheet1.write(sheet1_row,2,calculated_refractive_index[i])
        sheet1_row += 1

workbook.save(r'C:/Users/fhan/Desktop/Absorption Coefficient Calculation/Summary\Calculated refractive index.xls')

    
    
        
