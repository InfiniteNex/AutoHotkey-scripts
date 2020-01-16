; AutoRPGMover 2.0 (old name: PopulateRPGnew)
; 13.09.2019
; by Simeon P. Todorov

; Instructions
; Step 1 - Open RPGAssignments excel file and select the next empty cell in column A. Make sure there is enough space below for the new codes.
; Step 2 - Select and copy all unread emails containing new TWIN RPGs
; Step 3 - Activate this script


#SingleInstance, force

SetTitleMatchMode, 2																; ...
WinActivate, TWIN														    		; Activate outlook.
WinWait, TWIN															        	; Wait for the activation.
Send, ^c															            	; Copy the selected mails in case the user forgot to copy them.
ClipWait, 1															            	; Wait for the clipboard to fill or continue after 1 second.

FileAppend, %Clipboard%, %A_WorkingDir%\InputFile.txt								; Create a temporary file and insert the copied emails inside

Loop, Read, %A_WorkingDir%\InputFile.txt											; Read the file and determine the number of lines inside it.
{
   total_lines = %A_Index%															; ...
}

Loop, %total_lines%																	; Loop the trimmer by the amount of lines in the file.
{
FileReadLine, TrimThisString, %A_WorkingDir%\InputFile.txt, %A_Index%				; Read each line of the file.
Array := StrSplit(TrimThisString, ["(", ")"])										; Trim the line to the left and right of the parantheses.
TrimmedString := Array[2]															; Take the leftover code inbetween the parantheses and assign it to a variable.
FileAppend, %TrimmedString%`n, %A_WorkingDir%\OutputFile.txt						; Create a new file if it doesnt exist and insert the variable on a new line.
}

FileRead, OutputVar, %A_WorkingDir%\OutputFile.txt                                  ; Read output file
Sort, OutputVar, u                                                                  ; Sort the data inside, remove any duplicates (u) and save it into a variable
FileDelete, %A_WorkingDir%\OutputFile.txt                                           ; Delete the file
sleep, 300                                                                          ; ...
FileAppend, %OutputVar%,%A_WorkingDir%\OutputFile.txt                               ; Recreate the file and insert the data without any duplicates
sleep, 300                                                                          ; ...

Run, %A_WorkingDir%\OutputFile.txt													; Open the output file.
SetTitleMatchMode, 2																; ...
WinActivate, OutputFile																; Activate the output file.
WinWait, OutputFile																	; Wait for it to activate.
Send, {Delete}                                                                      ; Delete first empty row.
Send, ^a																			; Select all text.
Send, ^c																			; Copy it.
Send, ^s																			; Save the output file.
WinClose, OutputFile																; Close the output file.
WinActivate, RPG Assignment															; Activate the RPG Assignment... excel table.
WinWait, RPG Assignment																; Wait for it to activate.
Send, ^v																			; Paste the clipboard.
sleep, 500                                                                          ; Wait for contents to paste.
FileDelete, %A_WorkingDir%\InputFile.txt											; Delete the input temporary file.
FileDelete, %A_WorkingDir%\OutputFile.txt											; Delete the output temporary file.
WinActivate, TWIN                                                                   ; Activate Outlook.
WinWait, TWIN                                                                       ; Wait for it to activate.
Send, ^+1                                                                           ; Send hotkey to categorise all selected mails.
Send, ^q                                                                            ; Mark all selected mails as read.
ExitApp																				; Exit the script.