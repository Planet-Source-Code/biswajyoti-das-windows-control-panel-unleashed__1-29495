Attribute VB_Name = "Module1"
'This is the module which contains all the information about the param needed
'to call the Control Panel Functions ...
'Use it where you want it to ..
'Thanks ....
'Biswajyoti Das (askbiswa@hotmail.com)

'Accessability Options    ( ACCESS.CPL )
'---------------------------------------
'Accessability Properties
'Keyboard - 1
'Sound -    2
'Display -  3
'Mouse -    4
'General -  5

Function mdAcess(mdcode As Integer)
Shell "rundll32.exe shell32.dll,Control_RunDLL access.cpl,," & mdcode
End Function

'Add/Remove Programs    ( APPWIZ.CPL )
'-------------------------------------
'Add/Remove Programs Properties
'Install/Uninstall) - 1
'Windows Setup - 2
'Startup Disk - 3

Function mdAddRem(mdcode As Integer)
Shell "rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,," & mdcode
End Function

'Display Options    ( DESK.CPL )
'-------------------------------
'Display Properties
'Background - 0
'Screen Saver - 1
'Appearence - 2
'Settings - 3

Function mdDisp(mdcode As Integer)
Shell "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,," & mdcode
End Function

'Regional Settings    ( INTL.CPL )
'---------------------------------
'Regional Settings Properties
'Regional Settings - 0
'Number - 1
'Currency - 2
'Time - 3
'Date - 4
Function mdIntl(mdcode As Integer)
Shell "rundll32.exe shell32.dll,Control_RunDLL intl.cpl,," & mdcode
End Function

'Joystick Options    ( JOY.CPL )
'-------------------------------
'Joystick Properties

Function mdJoy()
Shell "rundll32.exe shell32.dll,Control_RunDLL joy.cpl"
End Function

'Mouse/Keyboard/Printers/Fonts Options    ( MAIN.CPL )
'-----------------------------------------------------
'Mouse Properties - 0
'Keyboard Properties - 1
'Printers - 2
'Fonts - 3

Function mdmkpf(mdcode As Integer)
Shell "rundll32.exe shell32.dll,Control_RunDLL main.cpl @" & mdcode
End Function
'Mail and Fax Options    ( MLCFG32.CPL )
'---------------------------------------
'Microsoft Exchange Profiles (General):
Function mdMaFax()
Shell "rundll32.exe shell32.dll,Control_RunDLL mlcfg32.cpl"
End Function

'Multimedia/Sounds Options    ( MMSYS.CPL )
'------------------------------------------
'Multimedia Properties
'Audio - 0
'Video - 1
'MIDI- 2
'CD Music - 3
'Advanced - 4
Function mdMul(mdcode As Integer)
Shell "rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl,," & mdcode
End Function

'Sounds Properties:
Function mdSound()
Shell "rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1"
End Function

'Modem Options    ( MODEM.CPL )
'------------------------------
'Modem Properties (General):
Function mdModem()
Shell "rundll32.exe shell32.dll,Control_RunDLL modem.cpl"
End Function
'
'Network Options    ( NETCPL.CPL )
'---------------------------------
'Network (Configuration):
Function mdNetwork()
Shell "rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl"
End Function
'Password Options    ( PASSWORD.CPL )
'------------------------------------
'Password Properties (Change Passwords):
Function mdPass()
Shell "rundll32.exe shell32.dll,Control_RunDLL password.cpl"
End Function
'System/Add New Hardware Options    ( SYSDM.CPL )
'------------------------------------------------
'System Properties
'General -0
'Device Manager - 1
'Hardware Profiles - 2
'Performance -3
Function mdSystem(mdcode As Integer)
Shell "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,," & mdcode
End Function

'Add New Hardware Wizard:
Function mdAddNH()
Shell "rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1"
End Function

'Date and Time Options    ( TIMEDATE.CPL )
'-----------------------------------------
'Date/Time Properties:
Function mdDT()
Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl"
End Function
