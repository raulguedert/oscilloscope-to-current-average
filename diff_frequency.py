from oscilloscope_mean import calculate_voltage_and_current_mean
from find_all_folders import find_folders
from win32com.client import Dispatch
import numpy as np

# Root path of all measurements
path = 'C:\\Users\\raulg\\OneDrive\\Doutorado\\Modelo Figado 1Hz 5kHz\\FIG\\'

# Define the percentage of pulse that will be considered to mean calculation
pulse_percentage = 50

ef = 100
voltage = 450

study_folders = find_folders(path)

study = 0

xlApp = Dispatch("Excel.Application")
xlApp.Visible = 1
xlApp.Workbooks.Add()

for study_folder in study_folders:
    study_index = study*6

    xlApp.ActiveSheet.Cells(
        (study_index + 1), 1).Value = 'Study ' + str(study + 1)
    xlApp.ActiveWorkbook.ActiveSheet.Cells(
        (study_index + 2), 1).Value = 'Frequency'
    xlApp.ActiveWorkbook.ActiveSheet.Cells(
        (study_index + 3), 1).Value = 'Mtr Voltage'
    xlApp.ActiveWorkbook.ActiveSheet.Cells(
        (study_index + 4), 1).Value = 'Osc Voltage'
    xlApp.ActiveWorkbook.ActiveSheet.Cells(
        (study_index + 5), 1).Value = 'Current'

    frequency_number = 1
    frequency_folders = find_folders(study_folder)

    for frequency_folder in frequency_folders:
        if (frequency_number == 1):
            voltage_and_current_mean = calculate_voltage_and_current_mean(
                pulse_percentage, frequency_folder, 1)

            xlApp.ActiveWorkbook.ActiveSheet.Cells(
                (study_index + 2), frequency_number + 1).Value = '1 Hz'
        else:
            voltage_and_current_mean = calculate_voltage_and_current_mean(
                pulse_percentage, frequency_folder, 5000)

            xlApp.ActiveWorkbook.ActiveSheet.Cells(
                (study_index + 2), frequency_number + 1).Value = '5 kHz'

        xlApp.ActiveSheet.Cells(
            (study_index + 1), frequency_number + 1).Value = str(ef) + ' kV/m'
        xlApp.ActiveWorkbook.ActiveSheet.Cells(
            (study_index + 3), frequency_number + 1).Value = str(voltage)
        xlApp.ActiveWorkbook.ActiveSheet.Cells(
            (study_index + 4), frequency_number + 1).Value = voltage_and_current_mean[0]
        xlApp.ActiveWorkbook.ActiveSheet.Cells(
            (study_index + 5), frequency_number + 1).Value = voltage_and_current_mean[1]

        frequency_number = frequency_number + 1

    study = study + 1

xlApp.ActiveWorkbook.SaveAs(path + 'Results')
del xlApp
