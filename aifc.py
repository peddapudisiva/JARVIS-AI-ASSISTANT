"""
Compatibility shim for Python 3.13 where the standard library 'aifc' module was removed.
SpeechRecognition imports 'aifc' for AIFF file support. We don't use AIFF files in Jarvis,
so this stub prevents import errors. If called, it raises NotImplementedError.
"""

class Error(Exception):
    pass

def open(file, mode='r'):
    raise NotImplementedError("AIFF reading not supported in this environment.")
