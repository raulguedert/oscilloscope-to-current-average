from oscilloscope_mean import calculate_voltage_and_current_mean
from find_all_folders import find_folders
from win32com.client import Dispatch
import numpy as np

# Root path of all measurements
path = 'C:\\Users\\raulg\\OneDrive\\Doutorado\\Modelo Cérebro\\'

# Define the percentage of pulse that will be considered to mean calculation
pulse_percentage = 50

study_folders = find_folders(path)

study = 0

xlApp = Dispatch("Excel.Application")
xlApp.Visible = 1
xlApp.Workbooks.Add()

for study_folder in study_folders:
    study_index = study*4

    xlApp.ActiveSheet.Cells(
        (study_index + 1), 1).Value = 'Study ' + str(study + 1)
    xlApp.ActiveWorkbook.ActiveSheet.Cells(
        (study_index + 2), 1).Value = 'Voltage'
    xlApp.ActiveWorkbook.ActiveSheet.Cells(
        (study_index + 3), 1).Value = 'Current'

    pulse = 1
    pulse_folders = find_folders(study_folder)

    for pulse_folder in pulse_folders:
        voltage_and_current_mean = calculate_voltage_and_current_mean(
            pulse_percentage, pulse_folder)

        xlApp.ActiveSheet.Cells(
            (study_index + 1), pulse + 1).Value = 'P' + str(pulse)
        xlApp.ActiveWorkbook.ActiveSheet.Cells(
            (study_index + 2), pulse + 1).Value = voltage_and_current_mean[0]
        xlApp.ActiveWorkbook.ActiveSheet.Cells(
            (study_index + 3), pulse + 1).Value = voltage_and_current_mean[1]

        pulse = pulse + 1

    study = study + 1

xlApp.ActiveWorkbook.SaveAs(path + 'Results')
del xlApp
