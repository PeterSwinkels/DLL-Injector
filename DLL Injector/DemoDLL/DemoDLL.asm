; Demo DLL v1.01 - by: Peter Swinkels, ***2014***

#INCLUDE "windef.inc"
#INCLUDE "winuser.inc"

DATA SECTION
hInst DD NULL
QuitMessage DB "Bye!", 0x0
StartUpMessage DB "Hello!", 0x0
Title DB "DEMO DLL", 0x0

CODE SECTION
START:

FRAME hInstance, dwReason, lpvReserved

CMP D[dwReason], DLL_PROCESS_ATTACH
JNE >
 MOV EAX, [hInstance]
 MOV [hInst], EAX
 PUSH MB_ICONEXCLAMATION, ADDR Title, ADDR StartUpMessage, 0x0
 CALL MessageBoxA
:

CMP D[dwReason], DLL_PROCESS_DETACH
JNE >
 PUSH MB_ICONEXCLAMATION, ADDR Title, ADDR QuitMessage, 0x0
 CALL MessageBoxA
:

MOV EAX, TRUE
RET
ENDF
