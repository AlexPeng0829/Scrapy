import os
import xlwt
import xlrd

class Excel_processer():
    
    def __init__(self, input_path_, start_row_, max_col_limit_ = 30, max_row_limit_ = 100):

        self.input_path = input_path_
        self.year = self.input_path.split('\\')[-1][0:4] + '-'
        self.header_exist_flag = False
        self.max_col_limit = max_col_limit_
        self.max_row_limit = max_row_limit_
        self.start_row = start_row_
        self.row_to_append = 0


    def read_excel(self, file_name, input_sheet, output_sheet):

        rows = min(input_sheet.nrows, self.max_row_limit)
        cols = min(input_sheet.ncols, self.max_col_limit)

        # if self.header_exist_flag == False:
        #     start_row = self.start_row - 1
        # else:
        #     start_row = self.start_row

        for each_row in range(self.start_row, rows): 
            for each_col in range(1, cols):
                # print(input_sheet.cell_value(each_row, each_col))
                each_cell_value = input_sheet.cell_value(each_row, each_col)
                if(each_row > self.start_row and each_col == 1):
                    next_col_value = input_sheet.cell_value(each_row, each_col + 1)
                    if each_cell_value == '' and next_col_value == '':
                        # self.header_exist_flag == True
                        return
                output_sheet.write(self.row_to_append, each_col, each_cell_value)
            date_prefix = self.year + file_name.split('.')[0]
            output_sheet.write(self.row_to_append, 0, date_prefix)
            self.row_to_append += 1
        # self.header_exist_flag == True
        return

    def rename_file_name(self, base_folder, old_file_name):

        name_patterns = old_file_name.split('.')
        month_patten = name_patterns[0][5:].rjust(2, '0')
        day_pattern = name_patterns[1].rjust(2, '0')
        new_file_name = month_patten + '-' + day_pattern + '.' + name_patterns[2]

        old_file = os.path.join(base_folder, old_file_name)
        new_file = os.path.join(base_folder, new_file_name)

        os.rename(old_file, new_file)
        print(new_file)


def main():
    
    input_path = "D:\\tmp\\静中心手术量\\2018手术安排"
    # input_path = "D:\\tmp\\静中心手术量\\2019"
    output_file = os.path.join(input_path, "overall.xls")
    start_row = 2

    excel_processer= Excel_processer(input_path, start_row)
    output_workbook = xlwt.Workbook()
    output_sheet = output_workbook.add_sheet("Sheet1")
    month_folder = os.listdir(input_path)
    print(month_folder)

    for each_month in month_folder:
        days_per_month = os.listdir(os.path.join(input_path, each_month))
        for each_day in days_per_month:

            each_day_base_path = os.path.join(input_path, each_month)
            each_day_file = os.path.join(each_day_base_path, each_day)
            print('Processing ', each_day_file)
            input_book = xlrd.open_workbook(each_day_file)
            input_sheet = input_book.sheet_by_name('Sheet1')
            excel_processer.read_excel(each_day, input_sheet, output_sheet)
            # rename_file_name(each_day_base_path, each_day)
             
    output_workbook.save(output_file)

if __name__ == "__main__":
    main()