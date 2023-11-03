import pandas as pd
from openpyxl.workbook import Workbook
from openpyxl.worksheet.table import TableStyleInfo, Table
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, DEFAULT_FONT
from startup_file_manage import FileManager



class ExcelManipulator:
    def __init__(self):
                
        self.file_manager = FileManager()
       
        self.chartsearch = self.file_manager.excel_manipulator_locations()[0]
        
        self.alphabet = {
            "1": "A",
            "2": "B",
            "3": "C",
            "4": "D",
            "5": "E",
            "6": "F",
            "7": "G",
            "8": "H",
            "9": "I",
            "10": "J",
            "11": "K",
            "12": "L",
            "13": "M",
            "14": "N",
            "15": "O",
            "16": "P",
            "17": "Q",
            "18": "R",
            "19": "S",
            "20": "T",
            "21": "U",
            "22": "V",
            "23": "W",
            "24": "X",
            "25": "Y",
            "26": "Z"
        }
        
        self.toggle_path = self.file_manager.excel_manipulator_locations()[1]
        
        self.column_lists = {
            
            'empty_inserted_columns': ['Current Week Client Comments',
                                       'Age',
                                       'Prior Week Client Comments',
                                       'RAI Reconciliation Comments'],
            
            'left_columns': ['DOS',
                             'Account #',
                             'MRN', 
                             'Patient Name',
                             'Carrier'],
            
            'hackensack_columns': ['DOS',
                                   'Account #',
                                   'MRN', 
                                   'Patient Name',
                                   'Carrier',
                                   'Department'],
            
            'singe_uac': ['UAC Reason - Provider(DOS)',
                          'UAC Reason'],
            
            'multiple_uac': ['UAC Reason 1 - Provider(DOS)',
                             'UAC Reason 1',
                             'UAC Reason 2 - Provider(DOS)',
                             'UAC Reason 2'],
            
            'right_columns': ['Current Week Client Comments',
                              'Age',
                              'Pro Date Sent To Client',
                              'Prior Week Client Comments',
                              'RAI Reconciliation Comments'],
            
            'lhi_search_list': ['Status',
                                'Comments']
            
        }

    def pandas_column_rearrange(self):
        
        self.toggle_dict = self.file_manager.json_dict(self.toggle_path)
        self.lhi_format = self.toggle_dict.get("Toggle LHI Search List")
        self.department_format = self.toggle_dict.get("Toggle Department")
        
        print(self.toggle_dict)
        print(self.lhi_format)
        print(self.department_format)
        
        excel_chartsearch = pd.read_excel(self.chartsearch)
       
        if self.lhi_format is False:
            for col in self.column_lists['empty_inserted_columns']:
                excel_chartsearch[col] = ""
        elif self.lhi_format is True:
            for col in self.column_lists['lhi_search_list']:
                excel_chartsearch[col] = ""
        else:
            print("Something went wrong...")

        name_col = excel_chartsearch.columns.tolist()
        
        if 'UAC Reason 1' in name_col and 'UAC Reason 2' in name_col:
            middle_columns = self.column_lists['multiple_uac']
        else:
            middle_columns = self.column_lists['singe_uac']
        
        
        if self.department_format is False:
            left_columns_chartsearch = excel_chartsearch[self.column_lists['left_columns']]  # noqa: E501
        elif self.department_format is True:
            left_columns_chartsearch = excel_chartsearch[self.column_lists['hackensack_columns']]  # noqa: E501
            print("hackensack used")
        else:
            quit("wrong input, exiting...")


        middle_columns_chartsearch = excel_chartsearch[middle_columns]


        if self.lhi_format is False:
            right_excel_chartsearch = excel_chartsearch[self.column_lists['right_columns']]  # noqa: E501
        elif self.lhi_format is True:
            right_excel_chartsearch = excel_chartsearch[self.column_lists['lhi_search_list']]  # noqa: E501
        else:
            quit("wrong input, exiting...")
        
        concatenated_excel_file = pd.concat([left_columns_chartsearch,
                                             middle_columns_chartsearch,
                                             right_excel_chartsearch],
                                             axis=1)
        
        return concatenated_excel_file

    def openpyxl_format_workbook(self, concatenated_excel_file): 
        alphabet = self.alphabet

        wb = Workbook()
        worksheet = wb.active
        worksheet.title = "RAI Report"

        for row in dataframe_to_rows(concatenated_excel_file, index=False, header=True):
            worksheet.append(row)

        col_names = []

        for col in concatenated_excel_file.columns:
            col_names.append(col)


        col_num = len(concatenated_excel_file.axes[1])
        col_num_str = str(col_num)
        row_num = len(concatenated_excel_file.axes[0])
        row_num_str = str(row_num + 1)

        col_to_letter = alphabet.get(col_num_str)

        table_dimension = "A1:" + col_to_letter + row_num_str

        excel_dimensions = (#col_names,
                            "Total columns:", col_num,
                            "Total rows:", row_num_str,
                            "Table Data:", table_dimension)
        
        print(excel_dimensions)
        
        if self.lhi_format is False:
        
            age_index = col_names.index('Age')
            date_index = col_names.index('Pro Date Sent To Client')
            
            age_index = str(age_index + 1)
            date_index = str(date_index + 1)
                
            
            col_to_letter_age = alphabet.get(age_index)
            
            #print("Age column", col_to_letter_age)

            col_to_letter_date = alphabet.get(date_index)
            
            #print("Date column", col_to_letter_date)
            
                    
            age_range = 2 
                    
            for row_num in range(age_range, int(row_num_str) + 1):
                worksheet[col_to_letter_age + '{}'.format(row_num)] = '=datedif(a{},today(),"D")'.format(str(age_range))  # noqa: E501
                age_range += 1
            
        col_to_letter = alphabet.get(str(col_num))
        
        
        table = Table(displayName = "table", ref = table_dimension)
        
            
        # Change table style to normal format
        style = TableStyleInfo(name = "TableStyleMedium2", showRowStripes = True)
            
        # Attatched the styles to table
        table.tableStyleInfo = style

        if self.lhi_format is False:

            for cell in worksheet[col_to_letter_date]:
                cell.alignment = Alignment(horizontal='center')  
                    
            for cell in worksheet[col_to_letter_age]:
                cell.alignment = Alignment(horizontal='center') 
                
            for cell in worksheet['A']:
                cell.alignment = Alignment(horizontal='center')  
            
            for cell in worksheet['B']:
                cell.alignment = Alignment(horizontal='center')  
            
            for cell in worksheet['C']:
                cell.alignment = Alignment(horizontal='left') 
                             
        # Attach table to worksheet
        worksheet.add_table(table)
            
        DEFAULT_FONT.size = 8
            
        return wb