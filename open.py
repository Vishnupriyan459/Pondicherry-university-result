import ctypes
import time
# MessageBox types
MB_OK = 0x0  # OK button
MB_OKCANCEL = 0x1  # OK and Cancel buttons
MB_YESNO = 0x4  # Yes and No buttons
MB_ICONINFORMATION = 0x40  # Information icon
MB_ICONWARNING = 0x30  # Warning icon
MB_ICONERROR = 0x10  # Error icon

# Function to display message box
def show_loading_alert(message):
    ctypes.windll.user32.MessageBoxW(0, message, "Loading", MB_OK )

# Function to display message box
def Finish_alert(message):
    ctypes.windll.user32.MessageBoxW(0, f"Successfully data stored in {message}", "Finished", MB_OK )








# Function to display message box
def show_message_box(title, message, style):
    return ctypes.windll.user32.MessageBoxW(0, message, title, style)

def show():
    result = show_message_box("Alert", "would you want to execute this file?", MB_OK | MB_ICONINFORMATION |MB_YESNO)
    print(result)
    if result == 7:  # 7 corresponds to No button
        show_loading_alert('Do you want to close?')
        exit()
        
def load():
    # Simulate some loading process
    show_loading_alert("Loading, please wait...")
    time.sleep(1)  # Simulating loading process
        




time.sleep(1) 



