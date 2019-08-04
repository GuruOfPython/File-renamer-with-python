import pandas as pd
import os
import re
import shutil
import csv

def file_name_parser(file_name):
    name_body, name_extension = os.path.splitext(file_name)
    try:
        index = re.findall(r'\(([0-9]+)\)', name_body)[0]
    except:
        index = ''

    return [name_body, name_extension, index]

class Renamer():
    def __init__(self):
        df = pd.read_excel('ExampleRename.xlsx', sheet_name='Sheet1')

        org_name_list = list(df['Original Name'])
        rename_list = list(df['New Name'])
        self.name_list = {}
        for org_name, rename in zip(org_name_list, rename_list):
            self.name_list[org_name] = rename

        self.result_history = csv.writer(open('result_history.csv', 'w', encoding='utf-8', newline=''))
        self.result_history.writerow(['Original File', 'Renamed File', 'Result'])

    def start_renaming(self):
        self.input_directory = os.getcwd() + '/Input'
        self.save_directory = self.input_directory + '/Renamed'

        if not os.path.exists(self.save_directory):
            os.makedirs(self.save_directory)

        for i, file in enumerate(os.listdir(self.input_directory)):
            if os.path.isfile(os.path.join(self.input_directory, file)):
                rename_str = self.rename(file=file)
                if rename_str:
                    print("[{}] {} -->> {}".format(i, file, rename_str))
                    org_file_full_name = self.input_directory + '/' + file
                    rename_file_full_name = self.save_directory + '/' + rename_str
                    shutil.copyfile(org_file_full_name, rename_file_full_name)
                    self.result_history.writerow([file, rename_str, 'Renamed'])
                else:
                    print("[{}] {} -->> No match".format(i, file))
                    self.result_history.writerow([file, '', 'No match'])

    def rename(self, file):
        rename_str = ""
        for org_name in self.name_list:
            if org_name.strip().lower() in file.lower():
                name_body, name_extension, index = file_name_parser(file)
                if not index:
                    rename_str = "{}_P0".format(self.name_list[org_name]) + name_extension
                else:
                    rename_str = "{}_{}".format(self.name_list[org_name], int(index)-1) + name_extension
                break

        return rename_str


if __name__ == '__main__':
    app = Renamer()
    app.start_renaming()

