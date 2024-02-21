import time


def welcome_initial_input():
    """Welcomes the user and displays the terminal menu options"""
    print(
        """\
 _____  _____  ______ 
|  __ \|  __ \|  ____|
| |__) | |__) | |__   
|  ___/|  ___/|  __|  
| |    | |    | |____ 
|_|    |_|    |______|
 __  __                                              _   
|  \/  |                                            | |  
| \  / | __ _ _ __   __ _  __ _ _ __ ___   ___ _ __ | |_ 
| |\/| |/ _` | '_ \ / _` |/ _` | '_ ` _ \ / _ \ '_ \| __|
| |  | | (_| | | | | (_| | (_| | | | | | |  __/ | | | |_ 
|_|  |_|\__,_|_| |_|\__,_|\__, |_| |_| |_|\___|_| |_|\__|
                           __/ |                         
  _____           _       |___/      
 / ____|         | |                
| (___  _   _ ___| |_ ___ _ __ ___  
 \___ \| | | / __| __/ _ \ '_ ` _ \ 
 ____) | |_| \__ \ ||  __/ | | | | |
|_____/ \__, |___/\__\___|_| |_| |_|
         __/ |                      
        |___/                                            
        """
    )


time.sleep(3)

options = [
    "1: New Equipment Input",
    "2: Quarantine An Item",
    "3: Repair Equipment Log",
    "4: Retire Equipment",
    "5: Update Equipment",
    "6: View Sheet",
]
