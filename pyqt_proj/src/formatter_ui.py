from PyQt6.QtWidgets import QTabWidget, QMainWindow, QWidget, QVBoxLayout, QGridLayout
from PyQt6.QtWidgets import QApplication, QComboBox, QPushButton, QCheckBox, QTextEdit, QHBoxLayout  # noqa: E501
import sys
from datetime import datetime
from startup_file_manage import FileManager
from excel_macro import excel_macro
from excel_manipulation import ExcelManipulator


class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
                
        self.setWindowTitle("RAI Formatter")
        self.setGeometry(175, 25, 1000, 700)
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # Layouts
        main_layout = QVBoxLayout() # Changed to a vertical layout for the entire window
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget) # Add the tab widget to the main layout

        # Create the first tab
        first_tab = QWidget()
        tab_widget.addTab(first_tab, "Action Center") # Add the first tab to the tab widget  # noqa: E501

        button_layout = QGridLayout() # Grid layout for buttons
        
        # Create a QTextEdit widget for the first output
        self.output_text1 = QTextEdit()
        self.output_text1.setReadOnly(True)

        # Create a QTextEdit widget for the second output
        self.output_text2 = QTextEdit()
        self.output_text2.setReadOnly(True)

        # Add QTextEdit widgets directly to the layout
        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_text1, 2)  # 2/3 of horizontal space
        output_layout.addWidget(self.output_text2, 1)  # 1/3 of horizontal space

        # Add layouts and widgets to the first tab
        first_tab_layout = QVBoxLayout()
        first_tab_layout.addLayout(button_layout)
        first_tab_layout.addLayout(output_layout)  # Add the horizontal output layout
        first_tab.setLayout(first_tab_layout)  # Set the layout for the first tab

        # Create a container widget for the layouts
        container_widget = QWidget()
        container_widget.setLayout(main_layout)
        self.setCentralWidget(container_widget)

        # Apply dark mode stylesheet
        self.setStyleSheet("""
            QMainWindow {
                background-color: #333;
                color: #FFF;
            }
                  
            QPushButton, QRadioButton, QCheckBox {
                background-color: #444;
                color: #FFF;
                text-align: center;
                border: 1px solid #555;
                padding: 5px;
            }
            
            QComboBox, QComboBox QAbstractItemView::item {
                background-color: #444;
                text-align: center;
                color: #FFF;
            }
            
            QFrame {
                background-color: #444;
                border: 1px solid #555;
            }
            QLineEdit {
                background-color: #555;
                color: #FFF;
                border: 1px solid #666;
                padding: 5px;
            }
            QTextEdit {
                background-color: #555;
                color: #FFF;
                border: 1px solid #666;
                padding: 5px;
            }
        """)
   
        dropdown_position = {
            # First button, dropdown menu
            "Select Client": (0, 0)
        }
        
        button_positions = {
            
            # Top Button Row | Excluding first
            "Excel Manipulation":   (0, 1),
            "Excel Macro":          (0, 2),
            
            # Middle Button Row
            "Button 4":     (1, 0),
            "Button 5":     (1, 1),
            "Clear Output": (1, 2),
            
            # Bottom Toggle Row
            "Toggle LHI Search List": (2, 0),
            "Toggle Department":      (2, 1),
            "Toggle SSRS":            (2, 2)
        }

        for text, (row, col) in dropdown_position.items():
            self.drop_menu = QComboBox()
            button_layout.addWidget(self.drop_menu, row, col)
            if text == "Select Client":
                    self.drop_menu.activated.connect(self.on_drop_menu_execute)

        for text, (row, col) in button_positions.items():
            if text.startswith("Toggle"):
                button = QPushButton(text)
                toggle_button = QCheckBox(text)
                button_layout.addWidget(button, row, col)
                button_layout.addWidget(toggle_button, row, col)
                
                if text == "Toggle LHI Search List":
                    toggle_button.stateChanged.connect(self.is_lhi)
                elif text == "Toggle Department":
                    toggle_button.stateChanged.connect(self.has_department)
                elif text == "Toggle SSRS":
                    toggle_button.stateChanged.connect(self.is_ssrs)
            else:
                button = QPushButton(text)
                button_layout.addWidget(button, row, col)
                
                if text == "Excel Manipulation":
                    button.clicked.connect(self.excel_manip)
                elif text == "Excel Macro":
                    button.clicked.connect(self.on_excel_macro)
                elif text == "Button 4":
                    button.clicked.connect(self.on_button4_click)
                elif text == "Button 5":
                    button.clicked.connect(self.on_button5_click)
                elif text == "Clear Output":
                    button.clicked.connect(self.clear_output)

        
        self.file_manager = FileManager()
        self.excel_manipulator = ExcelManipulator()
        
        todays_date = datetime.now().strftime('%m.%d.%Y')

        self.search_list_format_info = {
                            "client_name" : "NO CLIENT SELECTED",
                            "format" : "Search List",
                            "date" : todays_date,
                            "file_type" : ".xlsx"
        }

        self.toggle_states = {
            "Toggle LHI Search List" : False,
            "Toggle Department" : False,
            "Toggle SSRS" : False
        }

        self.populate_combo_box()

    def populate_combo_box(self):
        self.file_manager.create_folders()
        self.file_manager.create_json_files()
        self.file_manager.save_toggle_states_to_json(self.toggle_states)
        self.output_text2.clear()
        self.output_text2.append(self.file_manager.print_clients_json()[0])
              
        self.client_list = self.file_manager.print_clients_json()[1]
        self.client_list = self.client_list.values() 
        self.drop_menu.addItems(self.client_list)
    
    def clear_repop(self):
        # Clear the existing items in the QComboBox
        self.output_text2.clear()
    
        value = self.drop_menu.currentText()
        
        if not value:
            self.output_text1.append("No client selected in the dropdown.")
            return
        
        self.output_text1.append(f"Selected Client: {value}")
        
        selected_client = self.file_manager.select_client_by_value(value)
        
        if selected_client is not None:
            print(f"\nSelected Client: {selected_client}\n") 
            self.search_list_format_info["client_name"] = selected_client
            
            for key, value in self.search_list_format_info.items():
                self.output_text2.append(f"{key}: {value}")  
            
            self.output_text2.append("")
            
            for key, value in self.toggle_states.items():
                self.output_text2.append(f"{key}: {value}")  
                
        else:
            print("\nInvalid option.\n")
    
    def on_drop_menu_execute(self):
        self.clear_repop()               
        
    def excel_manip(self):       
        concat_excel_file = self.excel_manipulator.pandas_column_rearrange()
        wb = self.excel_manipulator.openpyxl_format_workbook(concat_excel_file)
        self.file_manager.save_xlsx(wb, self.search_list_format_info)

    def on_excel_macro(self):
        self.output_text1.append("Excel Macro Selected")
        search_list_format_info = self.search_list_format_info
        
        if search_list_format_info.get("client_name") == "NO CLIENT SELECTED":
            self.output_text1.append("No Client Selected")
        
        else:
            excel_macro(search_list_format_info)

    def on_button4_click(self):
        self.output_text1.append("'Button 4' pressed")
        # Add your specific functionality for Button 4 here

    def on_button5_click(self):
        self.output_text1.append("'Button 5' pressed")
        # Add your specific functionality for Button 5 here

    # Define functions for toggle buttons
    def is_lhi(self):
        if self.sender().isChecked():
            self.output_text1.append("'Toggle LHI Search List' is ON")
            self.toggle_states["Toggle LHI Search List"] = True

        else:
            self.output_text1.append("'Toggle LHI Search List' is OFF")
            self.toggle_states["Toggle LHI Search List"] = False
        self.file_manager.save_toggle_states_to_json(self.toggle_states)
        self.clear_repop()

    def has_department(self):
        if self.sender().isChecked():
            self.output_text1.append("'Toggle Department' is ON")
            self.toggle_states["Toggle Department"] = True

        else:
            self.output_text1.append("'Toggle Department' is OFF")
            self.toggle_states["Toggle Department"] = False
        self.file_manager.save_toggle_states_to_json(self.toggle_states)
        self.clear_repop()

    def is_ssrs(self):
        if self.sender().isChecked():
            self.output_text1.append("'Toggle SSRS' is ON")
            self.toggle_states["Toggle SSRS"] = True

        else:
            self.output_text1.append("'Toggle SSRS' is OFF")
            self.toggle_states["Toggle SSRS"] = False
        self.file_manager.save_toggle_states_to_json(self.toggle_states)
        self.clear_repop()
        
    def clear_output(self):
        self.output_text1.clear()
        self.populate_combo_box()
            
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec())