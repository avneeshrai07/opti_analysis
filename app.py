import subprocess
import time
import comtypes.client
import tkinter as tk
from tkinter import filedialog
from pythoncom import CoInitialize, CoUninitialize
import comtypes.automation
import os

def ensure_long_path(filepath):
    """ If path length exceeds the limit, use UNC path format """
    if len(filepath) > 260:
        filepath = r"\\?\\" + os.path.abspath(filepath)
    return filepath

def ask_user_for_file():
    """ Opens a file dialog and asks the user to select a .std file """
    print("entered into ask_user_for_file")
    root = tk.Tk()
    print("Tk() complete")
    
    root.withdraw()  # Hide the root window
    print("root withdraw complete")
    
    # Ensure root window is at least initialized
    root.update()  # Ensure the Tkinter window is updated before opening the dialog
    root.deiconify()  # Make sure root window is fully initialized (but still hidden)

    # Ask the user to select a .std file
    file_path = filedialog.askopenfilename(
        title="Select a STAAD file", 
        filetypes=[("STAAD Files", "*.std"), ("All Files", "*.*")]
    )
    
    root.destroy()  # Destroy the Tk root window after selection
    return file_path

def main():
    try:
        print("Initializing COM...")
        CoInitialize()

        # STAAD executable path
        staad_path = r"C:\Program Files\Bentley\Engineering\STAAD.Pro CONNECT Edition\STAAD\Bentley.Staad.exe"
        
        # Launch STAAD.Pro
        print("Launching STAAD.Pro...")
        subprocess.Popen([staad_path])
        time.sleep(14)  # Wait for STAAD to load

        # Connect to OpenSTAAD
        print("Connecting to OpenSTAAD...")
        openstaad = comtypes.client.GetActiveObject("StaadPro.OpenSTAAD")
        print("Connected to OpenSTAAD")
        command = openstaad.Command
        view = openstaad.View
        geometry = openstaad.Geometry

        # Flag necessary methods
        openstaad._FlagAsMethod("OpenSTAADFile")
        openstaad._FlagAsMethod("SetSilentMode")
        openstaad._FlagAsMethod("Analyze")
        openstaad._FlagAsMethod("isAnalyzing")
        openstaad._FlagAsMethod("SaveModel")
        openstaad._FlagAsMethod("GetErrorMessage")

        geometry._FlagAsMethod("SetFlagForHiddenEntities")
        geometry._FlagAsMethod("GetNodeList")
        geometry._FlagAsMethod("GetNodeCount")

        view._FlagAsMethod("ShowAllMembers")
        view._FlagAsMethod("ShowFront")
        command._FlagAsMethod("PerformAnalysis")

        print("Ask user for file path")
        # Ask the user to select a file
        filepath = ask_user_for_file()
        print("Got file path")
        # Handle case if the user cancels the file dialog
        if not filepath:
            print("No file selected. Exiting...")
            return

        # Ensure long path is handled properly
        filepath = ensure_long_path(filepath)

        # Open the STAAD file
        print(f"Opening STAAD file: {filepath}")
        openstaad.OpenSTAADFile(str(filepath))  # Pass the file path as a string
        time.sleep(3)

        openstaad.SetSilentMode(1)
        
        # View all the members (show hidden entities)
        print("Flags of Hidden")
        geometry.SetFlagForHiddenEntities(0)

       
        print("node_count",geometry.GetNodeCount())	


        
        # Perform initial analysis (with print option 6)
        print("Performing static analysis (PrintOption: 6)...")
        # command.PerformAnalysis(6)

       # Create a VARIANT object to hold the node list
        Nodelist = comtypes.automation.VARIANT()

        # Call GetNodeList to populate the Nodelist VARIANT
        print("Node",geometry.GetNodeList(Nodelist))

        # Convert the VARIANT to a Python list
        if Nodelist.value:  # Ensure there is data in the VARIANT
            NodeList = list(Nodelist.value)
            print("Node list:", NodeList)
        else:
            print("No nodes found.")


        # Save the model
        openstaad.SaveModel(1)
        time.sleep(6)

        # Enable silent mode
        openstaad.SetSilentMode(1)

        # Run the full analysis
        # print("Running full analysis in silent mode...")
        # openstaad.Analyze()

        # # Wait until the analysis completes
        # while openstaad.isAnalyzing():
        #     print("Analysis in progress...")
        #     time.sleep(5)

        # print("Analysis completed successfully.")


        errors = openstaad.GetErrorMessage()
        print("errors", errors)

    except Exception as e:
        print(f"Error occurred: {e}")

    finally:
        CoUninitialize()
        print("COM uninitialized.")

if __name__ == "__main__":
    main()
