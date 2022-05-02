# coding: utf-8
import scriptforge as SF
# other reading
# https://wiki.documentfoundation.org/Macros/Python_Design_Guide#Output_to_Consoles

def console(*args, **kwargs) -> None:
    serv: SF.SFScriptForge.SF_Exception = SF.CreateScriptService('ScriptForge.Exception')
    serv.PythonShell({**globals(), **locals()})
    print("Hello World")

g_exportedScripts = (console,)