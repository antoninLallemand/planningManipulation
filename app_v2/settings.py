import json
import os

class settingsHandler():

    def __init__(self):
        # For development, use a local directory
        if (False):
            self.settings_dir = os.path.join(os.getcwd(), "datas")
        else:  # Production uses %APPDATA%
            self.settings_dir = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "Plannings pharmacie")
        
        self.settings_file = os.path.join(self.settings_dir, "settings.json")
        os.makedirs(self.settings_dir, exist_ok=True)
        
        if not os.path.exists(self.settings_file):
            self.save(
                name="A RENSEIGNER",
                role=2,
                colors={
                    "off": "#92d050",
                    "sick": "#ff99cc",
                    "undefined": "#55ffff",
                    "vacation": "#ffff00",
                    "work": "#ffaa00"
                }
            )

    def save (self, name, role, colors):
        settings =  {
            "name": name,
            "role": role,
            "colors": colors
        }
        try:
            with open(self.settings_file, 'w') as file:
                json.dump(settings, file, indent=4, sort_keys=True)
            print("settings saved successfully.")
            return True
        except Exception as e:
            print(f"An error occurred saving settings: {e}")
            return False

    def retrieve (self):
        try:
            with open(self.settings_file, 'r') as file:
                settings = json.load(file)
                print(settings)
        except FileNotFoundError:
            print("The file was not found")
        except json.JSONDecodeError:
            print("Error decoding JSON")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
        return settings