import pandas as pd
from config.config import columns


class ConvertVltToTXT():
    def __init__(self, file_path, file_name, columns):
        self.file_path = file_path
        self.file_name = file_name
        self.columns = columns

    def load_vlt(self):
        # достаем параметры записи
        head_title = []
        with open(f"{self.file_path}/{self.file_name}.vlt", encoding='cp1251') as file:
            lines = file.readlines()
        head_title.append(lines[0].split())
        head_title.append(lines[1].split())
        parameters = lines[1].split()

        # создаем dataframe чистых данных
        self.data = pd.read_table(f"{self.file_path}{self.file_name}.vlt", encoding='cp1251', sep='\t', header=None,
                                  skiprows=[0, 1],
                                  engine='python')
        self.data.columns = self.columns

        for col in columns:
            arr = []
            for item in self.data[f"{col}"]:
                arr.append(float(item.split()[-1].replace(",", ".")))
            self.data[f"{col}"] = arr

        self.set_parameters = parameters
        self.head_title = head_title

        # return data, parameters, head_title

    def conversion_vlt_to_txt(self):
        def voltage_conversion(column, k_voltage=1.2):
            return column * float(k_voltage)

        def current_conversion(column, k_current=550):
            return column * float(k_current)

        def temp_conversion(column, k_temp=2):
            if k_temp == 0:
                return column * 100

            if k_temp == 1:
                return column * 157 - 200

            if k_temp == 2:
                return column * 200

            if k_temp == 3:
                return column * 250

        def pressure_conversion(column):
            return column * 10.1

        def displacement_conversion(column, k_displ=1):
            if k_displ == 0:
                return (column - 5) * 1 / 5

            if k_displ == 1:
                return (column - 5) * 10 / 5

            if k_displ == 2:
                return (column - 5) * 100 / 5

        def displacement_rate_conversion(column, k_displ_rate=1):
            if k_displ_rate == 0:
                return (column - 5) * 2 / 5

            if k_displ_rate == 1:
                return (column - 5) * 1 / 5

            if k_displ_rate == 2:
                return (column - 5) * 0.5 / 5

            if k_displ_rate == 3:
                return (column - 5) * 0.2 / 5

            if k_displ_rate == 4:
                return (column - 5) * 0.1 / 5

            if k_displ_rate == 5:
                return (column - 5) * 0.05 / 5

            if k_displ_rate == 6:
                return (column - 5) * 0.02 / 5

        def pirani_conversion(column):
            return column * 3.6

        def conversion_vlt_to_txt_i():
            self.data["Voltage"] = voltage_conversion(self.data["Voltage"])
            self.data["Current"] = current_conversion(self.data["Current"])
            self.data["Temp"] = temp_conversion(self.data["Temp"], int(self.set_parameters[11]))
            self.data["Pressure"] = pressure_conversion(self.data["Pressure"])
            self.data["Displ"] = displacement_conversion(self.data["Displ"], int(self.set_parameters[12]))
            self.data["Displ_Rate"] = displacement_rate_conversion(self.data["Displ_Rate"],
                                                                   int(self.set_parameters[13]))
            self.data["PIRANI"] = pirani_conversion(self.data["PIRANI"])
            self.data = self.data.round(6)

            with open(f"{self.file_path}{self.file_name}_convert.txt", "a", encoding='utf8') as f:
                for line in self.head_title:
                    f.write(str(line) + '\n')
                dfAsString = self.data.to_string(header=False, index=False)
                f.write(dfAsString)

        conversion_vlt_to_txt_i()


# file_path = '/home/user/PycharmProjects/SPS_file_reader_ver_2.0/vlt_txt_excel/'
# file_name = "27"

# s = ConvertVltToTXT(file_path, file_name, columns)
# s.load_vlt()
# s.conversion_vlt_to_txt()
# /vlt_txt_excel/27.vlt

# def read_input_file():
#     # достаем параметры записи
#     head_title = []
#     with open("vlt_txt_excel/27.vlt", encoding='cp1251') as file:
#         lines = file.readlines()
#     head_title.append(lines[0].split())
#     head_title.append(lines[1].split())
#     parameters = lines[1].split()
#
#     # создаем dataframe чистых данных
#     data = pd.read_table("vlt_txt_excel/27.vlt", encoding='cp1251', sep='\t', header=None, skiprows=[0, 1],
#                          engine='python')
#     data.columns = columns
#
#     for col in columns:
#         arr = []
#         for item in data[f"{col}"]:
#             arr.append(float(item.split()[-1].replace(",", ".")))
#         data[f"{col}"] = arr
#
#     return data, parameters, head_title
#
#
# data, set_parameters, head_title = read_input_file()
# # print(head_title)


# def voltage_conversion(column, k_voltage=1.2):
#     return column * float(k_voltage)
#
#
# def current_conversion(column, k_current=550):
#     return column * float(k_current)
#
#
# def temp_conversion(column, k_temp=2):
#     if k_temp == 0:
#         return column * 100
#
#     if k_temp == 1:
#         return column * 157 - 200
#
#     if k_temp == 2:
#         return column * 200
#
#     if k_temp == 3:
#         return column * 250
#
#
# def pressure_conversion(column):
#     return column * 10.1
#
#
# def displacement_conversion(column, k_displ=1):
#     if k_displ == 0:
#         return (column - 5) * 1 / 5
#
#     if k_displ == 1:
#         return (column - 5) * 10 / 5
#
#     if k_displ == 2:
#         return (column - 5) * 100 / 5
#
#
# def displacement_rate_conversion(column, k_displ_rate=1):
#     if k_displ_rate == 0:
#         return (column - 5) * 2 / 5
#
#     if k_displ_rate == 1:
#         return (column - 5) * 1 / 5
#
#     if k_displ_rate == 2:
#         return (column - 5) * 0.5 / 5
#
#     if k_displ_rate == 3:
#         return (column - 5) * 0.2 / 5
#
#     if k_displ_rate == 4:
#         return (column - 5) * 0.1 / 5
#
#     if k_displ_rate == 5:
#         return (column - 5) * 0.05 / 5
#
#     if k_displ_rate == 6:
#         return (column - 5) * 0.02 / 5
#
#
# def pirani_conversion(column):
#     return column * 3.6
#
#
# def conversion_vlt_to_txt(data, head_title):
#     data["Voltage"] = voltage_conversion(data["Voltage"])
#     data["Current"] = current_conversion(data["Current"])
#     data["Temp"] = temp_conversion(data["Temp"], int(set_parameters[11]))
#     data["Pressure"] = pressure_conversion(data["Pressure"])
#     data["Displ"] = displacement_conversion(data["Displ"], int(set_parameters[12]))
#     data["Displ_Rate"] = displacement_rate_conversion(data["Displ_Rate"], int(set_parameters[13]))
#     data["PIRANI"] = pirani_conversion(data["PIRANI"])
#     data = data.round(6)
#
#     with open("output_data/ggf.txt", "a", encoding='utf8') as f:
#         for line in head_title:
#             f.write(str(line) + '\n')
#         dfAsString = data.to_string(header=False, index=False)
#         f.write(dfAsString)
