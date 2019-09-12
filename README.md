# Oscilloscope to Current Average

This script is designed to process and transform CSV data from an oscilloscope to an Excel spreadsheet.

Data processing consists of averaging voltage and current from the last pulse of a pulse train. It was used on an Electroporation Mathematical Modelling process.

Data must be organized into folders, following this pattern:
- Root
  - Study 01
    - Measurement 01
      - Channel01.csv
      - Channel02.csv
    - Measurement 02
     - ...
    - Measurement n
  - Study 02
    - ...
  - Study 03
    - ...
