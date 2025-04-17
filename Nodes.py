import openstaadHandler
import Helpers
from comtypes import automation

def get_node_list(openstaad):
    command, view, geometry, output, support = openstaadHandler.openstaadModules(openstaad)

    try:
        n_nodes = geometry.GetNodeCount()
        print("nodeCount ",n_nodes)
        safe_list = Helpers.make_safe_array_long_size(n_nodes)  # Create SAFEARRAY with correct size
        nNodeList = Helpers.make_variant_vt_ref(safe_list, automation.VT_ARRAY | automation.VT_I4)

        geometry.GetNodeList(nNodeList)
        node_list = list(nNodeList.value)  # Ensure the node list is populated correctly
        return node_list
    except Exception as e:
        print(f"Error occurred while getting node list: {e}")
        return []
    

def get_support_nodes(openstaad):
    # openstaad, command, view, geometry, output, support = openstaadHandler.initialize_openstaad()
    command, view, geometry, output, support = openstaadHandler.openstaadModules(openstaad)
    try:
        n_nodes = geometry.GetNodeCount()
        safe_list = Helpers.make_safe_array_long_size(n_nodes)  # Create SAFEARRAY with correct size
        nNodeList = Helpers.make_variant_vt_ref(safe_list, automation.VT_ARRAY | automation.VT_I4)

        geometry.GetSupportNodes(nNodeList)

        support_node_list = list(nNodeList.value)  # Ensure the support nodes list is populated correctly
        return support_node_list
    except Exception as e:
        print(f"Error occurred while getting support nodes: {e}")
        return []

