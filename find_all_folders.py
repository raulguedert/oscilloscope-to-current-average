from glob import glob


def find_folders(path):
    try:
        folders = glob(path + '*\\')
        return folders
    except error as error:
        print(error)


def find_csv_files_in_folder(path):
    try:
        files = glob(path + '*.csv')
        return files
    except error as error:
        print(error)
