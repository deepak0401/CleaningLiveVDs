REM Fetching physical "Path" from Virtual directories

Set objWMIService = GetObject("winmgmts:root/microsoftiisv2") 
dim param1: param1 = WScript.Arguments.UnNamed(0)

Set colVdirs = objWMIService.ExecQuery("Select Path from IIsWebVirtualDirSetting where Name='" & param1&"' ")
For Each objVdir in colVdirs
	Wscript.Echo objVdir.Path 
	Wscript.Echo "------------------------------" 	
Next

