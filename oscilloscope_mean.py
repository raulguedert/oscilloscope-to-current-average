import pandas as pd
import numpy as np
from find_all_folders import find_csv_files_in_folder

voltage_channel = 1
current_channel = 2


def calculate_voltage_and_current_mean(pulse_percentage, path, frequency=1):

    files = find_csv_files_in_folder(path)

    if (frequency == 1):
        voltage_mean = calculate_mean(
            pulse_percentage, files[voltage_channel - 1])
        current_mean = calculate_mean(
            pulse_percentage, files[current_channel - 1])
    else:
        voltage_mean = calculate_mean_5k(
            pulse_percentage, files[voltage_channel - 1])
        current_mean = calculate_mean_5k(
            pulse_percentage, files[current_channel - 1])

    return np.array([voltage_mean, current_mean])


def calculate_mean(pulse_percentage, path):
    df = pd.read_csv(path)

    data_array = df.to_numpy()
    data = data_array[:, 4]

    sample_interval = data_array[0, 1]
    pulse_samples = int((100e-6)/float(sample_interval))
    considered_samples = int((pulse_percentage/100)*pulse_samples)

    # Found trigger sample
    trigger_sample = float(data_array[1, 1])

    # Isolate only considered data
    lower_limit = int((trigger_sample + pulse_samples/2) -
                      (considered_samples/2))
    upper_limit = int((trigger_sample + pulse_samples/2) +
                      (considered_samples/2))
    considered_data = data[lower_limit:upper_limit]

    data_mean = considered_data.mean()

    return data_mean


def calculate_mean_5k(pulse_percentage, path):
    df = pd.read_csv(path)

    data_array = df.to_numpy()
    data = data_array[:, 4]

    sample_interval = data_array[0, 1]
    pulse_samples = int((100e-6)/float(sample_interval))
    considered_samples = int((pulse_percentage/100)*pulse_samples)

    # Found trigger sample
    trigger_sample = float(data_array[1, 1])
    # Compensação de ajuste 10*7
    last_pulse_init = int(trigger_sample + pulse_samples*14 + 10*7)

    # Isolate only considered data
    lower_limit = int((last_pulse_init + pulse_samples/2) -
                      (considered_samples/2))
    upper_limit = int((last_pulse_init + pulse_samples/2) +
                      (considered_samples/2))
    considered_data = data[lower_limit:upper_limit]

    data_mean = considered_data.mean()

    return data_mean
