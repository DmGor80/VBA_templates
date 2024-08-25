Attribute VB_Name = "START"
Sub Satrt_All()


Load frmYesNO
frmYesNO.Show
Unload frmYesNO
sms = "proccesing"
'Start the time stamp
start_macros = Timer

Call Codes_first_last
 
'End macros Time stamp
end_macros = Timer
duration_macros = end_macros - start_macros

' Convert duration to minutes and seconds
    minutes = Int(duration_macros / 60)
    seconds = Round(duration_macros Mod 60, 0)
    
    ' Display the duration in MsgBox
MsgBox "Duration: " & minutes & " minutes " & seconds & " seconds"
'MsgBox ("Macros is finished. Check the results")

End Sub

