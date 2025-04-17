import subprocess
import time
import comtypes.client
import tkinter as tk
from tkinter import filedialog
from pythoncom import CoInitialize, CoUninitialize
from comtypes import automation
import os
import Helpers

def ensure_long_path(filepath):
    """If path length exceeds the limit, use UNC path format."""
    if len(filepath) > 260:
        filepath = r"\\?\\" + os.path.abspath(filepath)
    return filepath

def ask_user_for_file():
    """Opens a file dialog and asks the user to select a .std file."""
    print("Entered into ask_user_for_file")
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
    except Exception as e:
        print(f"STAAD.Pro can't be launched: {e}")
    
    try:
        # Connect to OpenSTAAD
        print("Connecting to OpenSTAAD...")
        openstaad = comtypes.client.GetActiveObject("StaadPro.OpenSTAAD")
        print("Connected to OpenSTAAD")
    except Exception as e:
        print(f"Can't connect to OpenSTAAD: {e}")
    
    try:
        print("Connecting to OpenSTAAD modules")
        command = openstaad.Command
        view = openstaad.View
        geometry = openstaad.Geometry
        output = openstaad.Output
        support = openstaad.Support
    except Exception as e:
        print(f"Can't connect to OpenSTAAD Modules: {e}")

    try:
        print("Flagging the OpenSTAAD module functions")
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

        output._FlagAsMethod("AreResultsAvailable")

        command._FlagAsMethod("PerformAnalysis")


        support._FlagAsMethod("GetSupportCount")
        support._FlagAsMethod("GetSupportNodes")
    except Exception as e:
        print(f"Can't Flag the OpenSTAAD Modules Functions: {e}")

    try:
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
    except Exception as e:
        print(f"Can't get .std file: {e}")

    try:
        # Open the STAAD file
        print(f"Opening STAAD file: {filepath}")
        openstaad.OpenSTAADFile(str(filepath))  # Pass the file path as a string
        time.sleep(3)
    except Exception as e:
        print(f"Can't OPEN .std file in OpenSTAAD: {e}")

    try:
        print("Using OpenSTAAD functions")
        openstaad.SetSilentMode(1)  # Enable silent mode to suppress UI popups
        
        # View all the members (show hidden entities)
        print("Flags of Hidden")
        geometry.SetFlagForHiddenEntities(0)  # Set flag to include all nodes, even hidden ones
       
        print("Node count:", geometry.GetNodeCount())  # Print the number of nodes

        def get_node_list_from_com():
            try:
                # """Retrieve node list from OpenSTAAD"""
                # Create an empty VARIANT to store the node list
                # nNodeList = Helpers.make_variant_vt_ref()
                n_nodes =  geometry.GetNodeCount()
                # node_ids = list(range(1, n_nodes + 1))
                
                # safe_list = Helpers.make_safe_array_long(node_ids)
                safe_list  = Helpers.make_safe_array_long_size(n_nodes)
                nNodeList = Helpers.make_variant_vt_ref(safe_list,  automation.VT_ARRAY | automation.VT_I4)

                # self._geometry.GetNodeList(lista)

                # return (lista[0])
                
                # Call GetNodeList method to retrieve the node list
                geometry.GetNodeList(nNodeList)
                # print("final node list: ",final_node_list)
                
                # Extract the node list from the VARIANT and return it
                node_list = list(nNodeList.value)
                # print(f"Node list = { node_list}")
                return node_list
            except Exception as e:
                print(f"Error occurred: {e}")
                return []

        def get_support_nodes_list_from_com():
            try:
                print("inside_get_support_nodes_list_from_com ")
                # Retrieve the total number of nodes from OpenSTAAD
                n_supports = support.GetSupportCount()  # Get the total number of nodes
                print(f"Total_nodes: {n_supports}")
                
                # Create a SAFEARRAY with the correct size for node IDs
                safe_list = Helpers.make_safe_array_long_size(n_supports)  # Create SAFEARRAY with size equal to the number of nodes
                nNodeList = Helpers.make_variant_vt_ref(safe_list, automation.VT_ARRAY | automation.VT_I4)

                # Call GetSupportNodes method to populate the SAFEARRAY with supported node data
                support.GetSupportNodes(nNodeList)
                
                # Extract the supported node list from the VARIANT and return it
                support_node_list = list(nNodeList.value)  # Convert the VARIANT to a list
                print(f"Support node list: {support_node_list}")
                
                return support_node_list
            except Exception as e:
                print(f"Error occurred: {e}")
                return []

        # Example usage
        node_list = get_node_list_from_com()
        print("Node list = ", node_list)



        support_node_list = get_support_nodes_list_from_com()
        print("support_node_list = ", support_node_list)
            
        # Perform initial analysis (with print option 6)
        print("Performing static analysis (PrintOption: 6)...")
        command.PerformAnalysis(6)

       

        

        # Save the model
        openstaad.SaveModel(1)
        time.sleep(6)

        # Enable silent mode
        openstaad.SetSilentMode(1)

        # Run the full analysis
        print("Running full analysis in silent mode...")
        openstaad.Analyze()

        # Wait until the analysis completes
        while openstaad.isAnalyzing():
            print("Analysis in progress...")
            time.sleep(5)

        print("Analysis completed successfully.")

        Is_Results = output.AreResultsAvailable()
        print("Results available", Is_Results)

        errors = openstaad.GetErrorMessage()
        print("Errors", errors)

    except Exception as e:
        print(f"Error occurred: {e}")

    finally:
        CoUninitialize()  # Uninitialize COM
        print("COM uninitialized.")

if __name__ == "__main__":
    main()
