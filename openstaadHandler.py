import comtypes.client
from comtypes import automation
import Helpers

def openstaadModules(openstaad):
    try:
        command = openstaad.Command
        view = openstaad.View
        geometry = openstaad.Geometry
        output = openstaad.Output
        support = openstaad.Support
        return command, view, geometry, output, support
    except Exception as e:
        print(f"Error in openstaadModules: {e}")
        return None, None, None, None, None
# Initialize OpenSTAAD COM object and modules
def initialize_openstaad():
    try:
        # Try to connect to the existing OpenSTAAD object
        openstaad = comtypes.client.GetActiveObject("StaadPro.OpenSTAAD")
        print("Successfully connected to OpenSTAAD.")
        
        return openstaad  # Return the openstaad object to be used by other functions
    except Exception as e:
        print(f"Error connecting to OpenSTAAD: {e}")
        return None  # Return None if the connection fails

# Flag necessary OpenSTAAD methods
def flag_openstaad_methods(openstaad):
    command = openstaad.Command
    view = openstaad.View
    geometry = openstaad.Geometry
    output = openstaad.Output
    support = openstaad.Support

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

# Retrieve node list from OpenSTAAD

