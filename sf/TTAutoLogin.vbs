Option Explicit	

'On Error Resume Next

	'Global Vars
		const HKLM = &H80000002
		const COMPUTER = "."

	'Vars
		dim objReg, objShell
		dim strPath, strCmd, strDefaultUserName, strDefaultPassword, strAutoAdminLogon, _
		    strDefaultDomainName, strAutoLogonCount, strForceAutoLogon, strFind, strLOCATION, _
		    strLastUsedUsername
		dim intGrade

	'Assign Vars
		strPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
		strDefaultUserName = "DefaultUserName"
		strDefaultPassword = "DefaultPassword"
		strAutoAdminLogon = "AutoAdminLogon"
		strDefaultDomainName = "DefaultDomainName"
		strAutoLogonCount = "AutoLogonCount"
		strLastUsedUsername = "LastUsedUsername"
		strCmd = "regedit /e c:\autologonbackup.reg " & chr(34) & "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Winlogon" & chr(34)
		strFind = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
		strLOCATION = InputBox("Please enter the three letter location acronym of the training tablet", "Location Acronym:")

		set objReg = GetObject("winmgmts:\\" & COMPUTER & "\root\default:StdRegProv")
		set objShell = WScript.CreateObject("WScript.Shell")

		if RegEntryExists(strFind) = True then
			intGrade = objShell.Run(strCmd, 0, TRUE)
			if intGrade = 0 then
				objReg.SetStringValue HKLM, strPath, strDefaultUserName, "MW" & strLOCATION & "Training"
				objReg.SetStringValue HKLM, strPath, strDefaultPassword, "training1"
				objReg.SetStringValue HKLM, strPath, strAutoAdminLogon, "1"
				objReg.SetStringValue HKLM, strPath, strDefaultDomainName, "wfm.pvt"
				objReg.SetStringValue HKLM, strPath, strLastUsedUsername, "MW" & strLOCATION & "Training"
'				objReg.SetStringValue HKLM, strPath, strForceAutoLogon, "1"
'				objReg.DeleteValue HKLM, strPath, strAutoLogonCount
				msgbox("AutoLogon Configured")
				msgbox("For changes to take please shutdown and restart your tablet.")
			else
				msgbox("Export Failed")				
			end if
		else
			msgbox("Unable to Find Reg") 
		end if


Function RegEntryExists(strFind)

on error resume next

	dim strEntry, ErrDescription, oShell
	set oShell = CreateObject("WScript.Shell")

	if (Right(strFind, 1)) = "\" then

		oShell.RegRead strFind

		select Case Err
			case 0
				RegEntryExists = TRUE
	
			case &H80070002
				ErrDescription = Replace(Err.description, strFind, "")
				Err.clear
				oShell.RegRead "HKEY_ERROR\"

				if (ErrDescription <> Replace(Err.description, "HKEY_ERROR\", "")) then
					RegEntryExists = TRUE
				else
					RegEntryExists = FALSE
				end if

			case else
				RegEntryExists = FALSE
		end select
	else
		oShell.RegRead strFind
		select case Err
			case 0
			RegEntryExists = TRUE
		case else
			RegEntryExists = FALSE
		end select
	end if

	on error goto 0

	set oShell = nothing

end function