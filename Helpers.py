import ctypes
# from openstaad.tools import *
from comtypes import automation
from comtypes import client
from comtypes import CoInitialize
import os

def ensure_long_path(filepath):
    """ If path length exceeds the limit, use UNC path format """
    if len(filepath) > 260:
        filepath = r"\\?\\" + os.path.abspath(filepath)
    return filepath

def make_safe_array_double(size): 
    return automation._midlSAFEARRAY(ctypes.c_double).create([0]*size)

def make_safe_array_int(size): 
    return automation._midlSAFEARRAY(ctypes.c_int).create([0]*size)

def make_safe_array_long_size(size):
    """Create an empty SAFEARRAY of long integers."""
    return automation._midlSAFEARRAY(ctypes.c_long).create([0] * size)  # Create a SAFEARRAY with space for node IDs

def make_variant_vt_ref(obj, var_type):
    var = automation.VARIANT()
    var._.c_void_p = ctypes.addressof(obj)
    var.vt = var_type | automation.VT_BYREF
    return var

def make_safe_str(): 
    return automation.c_char_p()

def make_safe_array_string(size):
    return automation._midlSAFEARRAY(automation.BSTR).create([""]*size)

def make_safe_array_long(data):
        if not isinstance(data, list):
            raise TypeError("Data must be a list of integers")
        
        size = len(data)
        safe_array = automation._midlSAFEARRAY(ctypes.c_long).create([0] * size)
        
        for i in range(size):
            safe_array[i] = data[i]
        
        return safe_array

