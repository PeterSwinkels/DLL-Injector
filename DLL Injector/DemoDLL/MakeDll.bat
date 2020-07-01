@ECHO OFF
IF EXIST DemoDLL.dll DEL DemoDLL.dll
IF EXIST DemoDLL.obj DEL DemoDLL.obj

GoASM.exe DemoDLL.asm
IF NOT %errorlevel% == 0 GOTO Quit
GoLink.exe DemoDLL.obj User32.dll /DLL

:Quit
IF EXIST DemoDLL.obj DEL DemoDLL.obj
PAUSE
