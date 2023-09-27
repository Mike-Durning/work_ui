import os
import json
from tkinter.filedialog import askdirectory
import shutil

class FileManager:
    def __init__(self):
        # Initialize folder paths
        self.current_directory = os.path.dirname(os.path.abspath(__file__))
        self.parent_directory = os.path.dirname(self.current_directory)
        
        self.settings_folder_path = os.path.join(self.parent_directory,
                                                 'settings')
        self.clients_folder_path = os.path.join(self.parent_directory,
                                                'clients')
        
        # Initialize file paths
        self.chartsearch_xlsx_path = os.path.join(self.clients_folder_path,
                                             'ChartSearch.xlsx')
        self.alphabet_json_path = os.path.join(self.settings_folder_path,
                                          "alphabet.json")
        
        self.clients_json_path = os.path.join(self.settings_folder_path,
                                          "clients.json")
        
        self.locations_json_path = os.path.join(self.settings_folder_path,
                                          "locations.json")
        
        self.toggle_format = os.path.join(self.settings_folder_path,
                                          "toggles.json")

        # Initialize data dictionaries
        
        self.toggle_states = {
            "Toggle LHI Search List": False,
            "Toggle Department": False,
            "Toggle SSRS": False
        }
        
        self.files_and_locations = {
            "settings_folder": self.settings_folder_path + "\\",
            "excel_folder_path": self.clients_folder_path + "\\",
            "clients_path": "clients.json",
            "alphabet_json": "alphabet.json",
            "excel_file": "ChartSearch.xlsx",
            "toggle_format": "toggles.json"
        }

        self.clients_data = {}

        self.alphabet_data = {
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

    def excel_manipulator_locations(self):
        chartsearch = self.chartsearch_xlsx_path= os.path.join(self.clients_folder_path,
                                                                'ChartSearch.xlsx')
        
        alphabet = self.alphabet_json_path = os.path.join(self.settings_folder_path,
                                          "alphabet.json")
        
        toggle = self.toggle_format = os.path.join(self.settings_folder_path,
                                          "toggles.json")
        return chartsearch, alphabet, toggle

    def create_folders(self):
        os.makedirs(self.settings_folder_path, exist_ok=True)
        os.makedirs(self.clients_folder_path, exist_ok=True)

    def create_json_files(self):
        # Create and update clients JSON file
        self.update_json_file(self.clients_json_path, self.clients_data)

        # Create and update alphabet JSON file
        self.update_json_file(self.alphabet_json_path, self.alphabet_data)

        # Create and update locations JSON file
        self.update_json_file(self.locations_json_path, self.files_and_locations)
        
        self.update_json_file(self.toggle_format, self.toggle_states)

    def update_json_file(self, file_path, data):
        if not os.path.exists(file_path):
            with open(file_path, 'w') as json_file:
                json.dump(data, json_file, indent=4)

    def json_dict(self, file_path):
        if os.path.exists(file_path):
            with open(file_path, 'r') as json_file:
                data = json.load(json_file)
            return data

    def print_clients_json(self):
        if os.path.exists(self.clients_json_path):
            with open(self.clients_json_path, 'r') as json_file:
                clients_data = json.load(json_file)
                # Create a string to store the formatted JSON data
                formatted_data = ""
                for key, value in clients_data.items():
                    formatted_data += f"{key}: {value}\n"
                return formatted_data, clients_data
        else:
            print("Clients JSON file does not exist.")
    
    def get_settings_files_and_path(self):
        if os.path.exists(self.locations_json_path):
            with open(self.locations_json_path, 'r') as locations_path:
                locations_data = json.load(locations_path)
                return locations_data
        else:
            print("Locations file does not exist.")
            return None
               
    def add_client_entry(self, client_name):
        if os.path.exists(self.clients_json_path):
            with open(self.clients_json_path, 'r') as json_file:
                clients_data = json.load(json_file)

            # Find the next available key
            next_key = str(len(clients_data) + 1)

            # Add the new entry
            clients_data[next_key] = client_name

            # Update the JSON file
            with open(self.clients_json_path, 'w') as json_file:
                json.dump(clients_data, json_file, indent=4)
                
            client_folder_path = os.path.join(self.clients_folder_path, client_name)
            os.makedirs(client_folder_path, exist_ok=True)
        
        else:
            print("Clients JSON file does not exist.")

    def save_toggle_states_to_json(self, toggle_states):
        try:
            with open(self.toggle_format, "w") as json_file:
                json.dump(toggle_states, json_file, indent=4)
        except Exception as e:
            print(f"An error occurred while saving toggle states: {str(e)}")

       
    def delete_client_entry(self, key_to_delete):
        if os.path.exists(self.clients_json_path):
            with open(self.clients_json_path, 'r') as json_file:
                clients_data = json.load(json_file)

            # Check if the key exists before deleting
            if key_to_delete in clients_data:
                # Delete the entry
                del clients_data[key_to_delete]

                # Update the keys to be consecutive
                updated_clients_data = {}
                new_key = 1
                for old_key in sorted(map(int, clients_data.keys())):
                    updated_clients_data[str(new_key)] = clients_data[str(old_key)]
                    new_key += 1

                # Update the JSON file with the updated data
                with open(self.clients_json_path, 'w') as json_file:
                    json.dump(updated_clients_data, json_file, indent=4)
            else:
                print(f"Key {key_to_delete} does not exist in the clients JSON.")
        else:
            print("Clients JSON file does not exist.")

    def select_client_by_key(self, key):
        if os.path.exists(self.clients_json_path):
            with open(self.clients_json_path, 'r') as json_file:
                clients_data = json.load(json_file)

            if key in clients_data:
                return clients_data[key]
            else:
                print(f"Key {key} does not exist in the clients JSON.")
                return None
        else:
            print("Clients JSON file does not exist.")
            return None
          
    def select_client_by_value(self, selected_value):
        if os.path.exists(self.clients_json_path):
            with open(self.clients_json_path, 'r') as json_file:
                clients_data = json.load(json_file)

            for key, client_value  in clients_data.items():
                if client_value == selected_value:
                    return client_value

        else:
            print("Clients JSON file does not exist.")
            return None

    def client_folder_export(self, search_list_format_info):
        search_list = search_list_format_info["format"]
        todays_date = search_list_format_info["date"]
              
        # Check if the clients folder exists
        if not os.path.exists(self.clients_folder_path):
            print("Clients folder does not exist.")
            return

        # Ask the user for the export location
        download_directory = askdirectory(title="Select Location",
                                          initialdir=self.clients_folder_path)

        if not download_directory:
            print("Export canceled.")
            return

        try:
            # Create a directory to store the exported client folders
            export_directory = os.path.join(download_directory,
                                            f'{todays_date} {search_list}s')
            os.makedirs(export_directory, exist_ok=True)

            # Iterate through the contents of clients_folder_path
            for item in os.listdir(self.clients_folder_path):
                item_path = os.path.join(self.clients_folder_path, item)
                
                # Check if the item is a directory
                if os.path.isdir(item_path):
                    # Copy the contents of the directory to export_directory
                    shutil.copytree(item_path, os.path.join(export_directory, item))
                    
            # Delete the "ChartSearch.xlsx" file if it exists in export_directory
            chart_search_file = os.path.join(export_directory, "ChartSearch.xlsx")
            if os.path.exists(chart_search_file):
                os.remove(chart_search_file)

        except Exception as e:
            print(f"An error occurred during export: {str(e)}")
            
    def save_xlsx(self, wb, search_list_format_info):
        if wb is not None:
            todays_date = search_list_format_info["date"]
            file_type = search_list_format_info["file_type"]
            search_list = search_list_format_info["format"]
            client_name = search_list_format_info["client_name"]

            filename_save = client_name + " " + search_list + " " + todays_date + file_type  # noqa: E501

            download_Directory = askdirectory(title="Select Client Folder",
                                            initialdir=self.clients_folder_path)

            folder_path = download_Directory
            wb.save(os.path.join(folder_path, filename_save))
            wb.close()

            file_name_path = os.path.join(folder_path, filename_save)

            os.startfile(file_name_path)
            os.startfile(folder_path)
        else:
            print("Try restarting app, you may have tried to made duplicates.")