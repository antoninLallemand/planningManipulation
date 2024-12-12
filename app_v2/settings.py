import json

class settingsHandler():

    def save (self, name, role, colors):
        settings =  {
            "name": name,
            "role": role,
            "colors": colors
        }
        try:
            with open('./settings.json', 'w') as file:
                json.dump(settings, file, indent=4, sort_keys=True)
            print("settings saved successfully.")
        except Exception as e:
            print(f"An error occurred saving settings: {e}")

    def retrieve (self):
        try:
            with open('./settings.json', 'r') as file:
                settings = json.load(file)
                print(settings)
        except FileNotFoundError:
            print("The file was not found")
        except json.JSONDecodeError:
            print("Error decoding JSON")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
        return settings