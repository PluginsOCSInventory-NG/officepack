'----------------------------------------------------------
' Plugin for OCS Inventory NG 2.x
' Script :		Retrieve Microsoft Office informations
' Version :		2.25
' Date :		09/03/2020
' Author :		Creative Commons BY-NC-SA 3.0
' Author :		Nicolas DEROUET (nicolas.derouet[gmail]com)
' Contributor :		Stéphane PAUTREL (acb78.com), Eduardo Mozart de Oliveira
'----------------------------------------------------------
' OS checked [X] on	32b	64b	(Professionnal edition)
'	Windows XP	[X]
'	Windows Vista	[X]	[X]
'	Windows 7	[X]	[X]
'	Windows 8.1	[X]	[X]	
'	Windows 10	[X]	[X]
'	Windows 2k8R2		[X]
'	Windows 2k12R2		[X]
'	Windows 2k16		[X]
' ---------------------------------------------------------
' Note :	No checked on Windows 8
' Included :	Office 2016 and 365 versions
' ---------------------------------------------------------
On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002

Const OFFICE_ALL = "78E1-11D2-B60F-006097C998E7}.0001-11D2-92F2-00104BC947F0}.6000-11D3-8CFE-0050048383C9}.6000-11D3-8CFE-0150048383C9}.7000-11D3-8CFE-0150048383C9}.BE5F-4ED1-A0F7-759D40C7622E}.BDCA-11D1-B7AE-00C04FB92F3D}.6D54-11D4-BEE3-00C04F990354}.CFDA-404E-8992-6AF153ED1719}.{9AC08E99-230B-47e8-9721-4577B7F124EA}"
' Supported Office Families:
' Office 2000 -> KB230848 - https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/230848
' Office XP -> KB302663 - https://www.betaarchive.com/wiki/index.php?title=Microsoft_KB_Archive/302663
' Office 2003 -> KB832672 - https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/826217
' Office 2007 -> KB928516 - https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/928516
' Office 2010 -> KB2186281 - https://support.microsoft.com/en-us/topic/description-of-the-numbering-scheme-for-product-code-guids-in-office-2010-cceaef56-3d3f-1cae-8577-b4de3beaacfa
' Office 2013 -> KB2786054 - https://support.microsoft.com/help/2786054 
' Office 2016, O365 -> https://docs.microsoft.com/en-us/office/troubleshoot/office-suite-issues/numbering-scheme-for-product-guid
Const OFFICEID = "000-0000000FF1CE}"

Dim aOffID(5,1)
aOffID(0,0) = "XP"
aOffID(0,1) = "10.0"
aOffID(1,0) = "2003"
aOffID(1,1) = "11.0"
aOffID(2,0) = "2007"
aOffID(2,1) = "12.0"
aOffID(3,0) = "2010"
aOffID(3,1) = "14.0"
aOffID(4,0) = "2013"
aOffID(4,1) = "15.0"
aOffID(5,0) = "2016"
aOffID(5,1) = "16.0"

Dim aOSPPVersions

Set oCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
oCtx.Add "__ProviderArchitecture", 64

Set oLocator = CreateObject("Wbemscripting.SWbemLocator")
Set oReg = oLocator.ConnectServer("", "root\default", "", "", , , , oCtx).Get("StdRegProv")

osType = 32
oReg.GetStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "PROCESSOR_ARCHITECTURE", osProc
If osProc = "AMD64" Then osType = 64

wow = ""
If osType = "64" Then wow = "WOW6432Node\"
schKey97 "SOFTWARE\" & wow & "Microsoft\"
schKey2K "Office", "SOFTWARE\" & wow & "Microsoft\Office\9.0\", Array("0000","0001","0002","0003","0004","0010","0011","0012","0013","0014","0016","0017","0018","001A","004F"), "78E1-11D2-B60F-006097C998E7"
schKey2K "Visio", "SOFTWARE\" & wow & "Microsoft\Visio\6.0\", Array("B66F45DC"), "853B-11D3-83DE-00C04F3223C8"

For a = LBound(aOffID, 1) To UBound(aOffID, 1)
	schKey "SOFTWARE\Wow6432Node\Microsoft\Office\" & aOffID(a,1) & "\Registration", False
	schKey "SOFTWARE\Microsoft\Office\" & aOffID(a,1) & "\Registration", True
	schKey "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\" & aOffID(a,1) & "\Registration", False
	schKey "SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\" & aOffID(a,1) & "\Registration", True
Next

Sub schKey97(regKey)
	oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Office\8.0", "BinDirPath", oDir97
	If IsNull(oDir97) Then Exit Sub
	oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Microsoft Reference\BookshelfF\96L", "PID", oProdID
	oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Windows\CurrentVersion\Uninstall\Office8.0", "DisplayName", oProd
	oInstall = "1"
	If IsNullOrEmpty(oProd) Then
	oInstall = "0"
	oProd = "Microsoft Office 97"
	End If
	writeXML "97",oProd,oProdID,32,"",oInstall,"",""
End Sub

Sub schKey2K(name, regKey, guid1, guid2)
	oProd = Null
	oInstall = "0"
	oReg.GetBinaryValue HKEY_LOCAL_MACHINE, regKey & "Registration\DigitalProductID", "", aDPIDBytes
	oKey = ""
	If Not IsNull(aDPIDBytes) Then oKey = decodeKey(aDPIDBytes)

	oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Registration\ProductID", "", oProdID
	If IsNull(oProdID) Then Exit Sub

	oReg.EnumKey HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Windows\CurrentVersion\Uninstall\", aKeys
	If Not IsNull(aKeys) Then
	For Each guid In aKeys
		If UCase(Right(guid,Len(guid)-InStr(guid,"-"))) = guid2 & "}" Then
			For i = LBound(guid1) To UBound(guid1)
				If UCase(Left(guid,Len(guid1(i)) + 1)) = "{" & guid1(i) Then
					oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Windows\CurrentVersion\Uninstall\" & guid, "DisplayName", oProd
					oGUID = guid
					oInstall = "1"
				End If
			Next
		End If
	Next
	End If

	If IsNullOrEmpty(oProd) Then oProd = "Microsoft " & name & " 2000"
	writeXML "2000",oProd,oProdID,32,oGUID,oInstall,oKey,""
End Sub

Sub schKey(regKey, likeOS) 
	' WScript.Echo "HKEY_LOCAL_MACHINE\" & regKey
	
	If oReg.EnumKey(HKEY_LOCAL_MACHINE, regKey, aGUIDKeys) = 0 Then
		If Not IsNull(aGUIDKeys) Then
			For Each GUIDKey In aGUIDKeys
				' WScript.Echo "InStr(OFFICE_ALL, " & UCase(Right(GUIDKey, 28)) & ") > 0: " & CBool(InStr(OFFICE_ALL, UCase(Right(soGUID, 28))) > 0)
				' WScript.Echo "InStr(" & OFFICEID & ", " & UCase(Right(GUIDKey, 17)) & ") > 0: " & CBool(InStr(OFFICEID, UCase(Right(GUIDKey, 17))) > 0)
				If InStr(OFFICE_ALL, UCase(Right(GUIDKey, 28))) > 0 OR _
				InStr(OFFICEID, UCase(Right(GUIDKey, 17))) > 0 Then
					schKey regKey & "\" & GUIDKey, likeOS
					Exit Sub
				End If
			Next
		End If
	Else
		Exit Sub
	End If
	
	' WScript.Echo "HKEY_LOCAL_MACHINE\" & regKey
	
	oGUID = ""
	oSKUID = ""
	oReg.GetBinaryValue HKEY_LOCAL_MACHINE, regKey, "DigitalProductID", aDPIDBytes
	' WScript.Echo "IsNull(aDPIDBytes): " & IsNull(aDPIDBytes)
	If IsNull(aDPIDBytes) Then
	' aOffID(a,1) = 16.0
	' aOSPPVersions(0) = 16
	' aOSPPVersions(1) = 0
		aOSPPVersions = Split(aOffID(a,1), ".")
		Set oOfficeOSPPInfos = getOfficeOSPPInfos(aOSPPVersions(0))
		If Not oOfficeOSPPInfos.Count = 0 Then
			oProdID = oOfficeOSPPInfos.Item("ProductID")
		oSKUID = oOfficeOSPPInfos.Item("SKUID")
		' We'll use the oOfficeOSPPInfos.Item("LicenseName") only if we cannot found the Office Uninstall key.
		' oProd = oOfficeOSPPInfos.Item("LicenseName")
		' oVer is set from aOffID variable (see below).
		' oVer = oOfficeOSPPInfos.Item("LicenseDescription")
		oNote = oOfficeOSPPInfos.Item("ErrorDescription")
		oKey = oOfficeOSPPInfos.Item("PartialProductKey")
		' oInstall = "1"
		End If
	Else
		oKey = decodeKey(aDPIDBytes)
		oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey, "ProductID", oProdID
	End If
	If Mid (oProdID,7,3) = "OEM" Then oOEM = " OEM"
	oVer = aOffID(a,0) ' aOffID(a,0) = XP/20XX
	On Error Resume Next
	oEdit = ""
	If (oVer = "2010") Then:uBoundDPIDEdit = 312:Else:uBoundDPIDEdit = 320:End If
	For i = 280 to uBoundDPIDEdit Step 2
		If aDPIDBytes(i) <> 0 Then
			oEdit = oEdit & Chr(aDPIDBytes(i))
		End If
	Next
	oNote = oEdit
	On Error Goto 0
	If (oVer = "2016") Then:If IsNull(aDPIDBytes) Then:oVer = "2019":Else:oVer = "2016":End If:End If
	
	oGUID = Right(regKey,InStr(StrReverse(regKey),"\")-1)
	If Not oSKUID = "" Then oGUID = oSKUID
	oBit = osType
	If Not likeOS Then oBit = 32 
	If CInt(Left(aOffID(a,1),2)) > 11 Then:If Mid(oGUID, 21, 1) = "0" Then:oBit = 32:Else If Mid(oGUID, 20, 1) = "1" Then:oBit = 64:End If:End If
	oInstall = "0"
	wow = ""
	If Not likeOS Then wow = "WOW6432Node\"
	
	If (oVer = "XP") Then    
		hProd = Mid ((Right(oGUID,38)),4,2)
		Select Case hProd
			Case "11" oProd = "Microsoft Office XP Professional"
			Case "12" oProd = "Microsoft Office XP Standard"
			Case "13" oProd = "Microsoft Office XP Small Business"
			Case "14" oProd = "Microsoft Office XP Web Server"
			Case "15" oProd = "Microsoft Access 2002"
			Case "16" oProd = "Microsoft Excel 2002"
			Case "17" oProd = "Microsoft FrontPage 2002"
			Case "18" oProd = "Microsoft PowerPoint 2002"
			Case "19" oProd = "Microsoft Publisher 2002"
			Case "1A" oProd = "Microsoft Outlook 2002"
			Case "1B" oProd = "Microsoft Word 2002"
			Case "1C" oProd = "Microsoft Access 2002 Runtime"
			Case "27" oProd = "Microsoft Project 2002"
			Case "28" oProd = "Microsoft Office XP Professional with FrontPage"
			Case "31" oProd = "Microsoft Project 2002 Web Client"
			Case "32" oProd = "Microsoft Project 2002 Web Server"
			Case "3A" oProd = "Project 2002 Standard"
			Case "3B" oProd = "Project 2002 Professional"
			Case "51" oProd = "Microsoft Office Visio Professional 2002"
			Case "54" oProd = "Microsoft Office Visio Standard 2002"
		End Select
	End If
	
	If (oVer = "2003") Then
		hProd = Mid ((Right(oGUID,38)),4,2)
		Select Case hProd
			Case "11" oProd = "Microsoft Office Professional Enterprise Edition 2003"
			Case "12" oProd = "Microsoft Office Standard Edition 2003"
			Case "13" oProd = "Microsoft Office Basic Edition 2003"
			Case "14" oProd = "Microsoft Windows SharePoint Services 2.0"
			Case "15" oProd = "Microsoft Office Access 2003"
			Case "16" oProd = "Microsoft Office Excel 2003"
			Case "17" oProd = "Microsoft Office FrontPage 2003"
			Case "18" oProd = "Microsoft Office PowerPoint 2003"
			Case "19" oProd = "Microsoft Office Publisher 2003"
			Case "1A" oProd = "Microsoft Office Outlook Professional 2003"
			Case "1B" oProd = "Microsoft Office Word 2003"
			Case "1C" oProd = "Microsoft Office Access 2003 Runtime"
			Case "1E" oProd = "Microsoft Office 2003 User Interface Pack"
			Case "1F" oProd = "Microsoft Office 2003 Proofing Tools"
			Case "23" oProd = "Microsoft Office 2003 Multilingual User Interface Pack"
			Case "24" oProd = "Microsoft Office 2003 Resource Kit"
			Case "26" oProd = "Microsoft Office XP Web Components"
			Case "2E" oProd = "Microsoft Office 2003 Research Service SDK"
			Case "44" oProd = "Microsoft Office InfoPath 2003"
			Case "83" oProd = "Microsoft Office 2003 HTML Viewer"
			Case "92" oProd = "Windows SharePoint Services 2.0 English Template Pack"
			Case "93" oProd = "Microsoft Office 2003 English Web Parts and Components"
			Case "A1" oProd = "Microsoft Office OneNote 2003"
			Case "A4" oProd = "Microsoft Office 2003 Web Components"
			Case "A5" oProd = "Microsoft SharePoint Migration Tool 2003"
			Case "AA" oProd = "Microsoft Office PowerPoint 2003 Presentation Broadcast"
			Case "AB" oProd = "Microsoft Office PowerPoint 2003 Template Pack 1"
			Case "AC" oProd = "Microsoft Office PowerPoint 2003 Template Pack 2"
			Case "AD" oProd = "Microsoft Office PowerPoint 2003 Template Pack 3"
			Case "AE" oProd = "Microsoft Organization Chart 2.0"
			Case "CA" oProd = "Microsoft Office Small Business Edition 2003"
			Case "D0" oProd = "Microsoft Office Access 2003 Developer Extensions"
			Case "DC" oProd = "Microsoft Office 2003 Smart Document SDK"
			Case "E0" oProd = "Microsoft Office Outlook Standard 2003"
			Case "E3" oProd = "Microsoft Office Professional Edition 2003 (with InfoPath 2003)"
			Case "FD" oProd = "Microsoft Office Outlook 2003 (distributed by MSN)"
			Case "FF" oProd = "Microsoft Office 2003 Edition Language Interface Pack"
			Case "F8" oProd = "Remove Hidden Data Tool"
			Case "3A" oProd = "Microsoft Office Project Standard 2003"
			Case "3B" oProd = "Microsoft Office Project Professional 2003"
			Case "32" oProd = "Microsoft Office Project Server 2003"
			Case "51" oProd = "Microsoft Office Visio Professional 2003"
			Case "52" oProd = "Microsoft Office Visio Viewer 2003"
			Case "53" oProd = "Microsoft Office Visio Standard 2003"
			Case "55" oProd = "Microsoft Office Visio for Enterprise Architects 2003"
			Case "5E" oProd = "Microsoft Office Visio 2003 Multilingual User Interface Pack"
		End Select
	End If
	
	If (oVer = "2007") Then
		hProd = Mid ((Right(oGUID,38)),11,4)
		Select Case hProd
			Case "0011" oProd = "Microsoft Office Professional Plus 2007"
			Case "0012" oProd = "Microsoft Office Standard 2007"
			Case "0013" oProd = "Microsoft Office Basic 2007"
			Case "0014" oProd = "Microsoft Office Professional 2007"
			Case "0015" oProd = "Microsoft Office Access 2007"
			Case "0016" oProd = "Microsoft Office Excel 2007"
			Case "0017" oProd = "Microsoft Office SharePoint Designer 2007"
			Case "0018" oProd = "Microsoft Office PowerPoint 2007"
			Case "0019" oProd = "Microsoft Office Publisher 2007"
			Case "001A" oProd = "Microsoft Office Outlook 2007"
			Case "001B" oProd = "Microsoft Office Word 2007"
			Case "001C" oProd = "Microsoft Office Access Runtime 2007"
			Case "0020" oProd = "Microsoft Office Compatibility Pack for Word, Excel, and PowerPoint 2007 File Formats"
			Case "0026" oProd = "Microsoft Expression Web"
			Case "002E" oProd = "Microsoft Office Ultimate 2007"
			Case "002F" oProd = "Microsoft Office Home and Student 2007"
			Case "0030" oProd = "Microsoft Office Enterprise 2007"
			Case "0031" oProd = "Microsoft Office Professional Hybrid 2007"
			Case "0033" oProd = "Microsoft Office Personal 2007"
			Case "0035" oProd = "Microsoft Office Professional Hybrid 2007"
			Case "003A" oProd = "Microsoft Office Project Standard 2007"
			Case "003B" oProd = "Microsoft Office Project Professional 2007"
			Case "0044" oProd = "Microsoft Office InfoPath 2007"
			Case "0051" oProd = "Microsoft Office Visio Professional 2007"
			Case "0052" oProd = "Microsoft Office Visio Viewer 2007"
			Case "0053" oProd = "Microsoft Office Visio Standard 2007"
			Case "00A1" oProd = "Microsoft Office OneNote 2007"
			Case "00A3" oProd = "Microsoft Office OneNote Home Student 2007"
			Case "00A7" oProd = "Calendar Printing Assistant for Microsoft Office Outlook 2007"
			Case "00A9" oProd = "Microsoft Office InterConnect 2007"
			Case "00AF" oProd = "Microsoft Office PowerPoint Viewer 2007 (English)"
			Case "00B0" oProd = "The Microsoft Save as PDF add-in"
			Case "00B1" oProd = "The Microsoft Save as XPS add-in"
			Case "00B2" oProd = "The Microsoft Save as PDF or XPS add-in"
			Case "00BA" oProd = "Microsoft Office Groove 2007"
			Case "00CA" oProd = "Microsoft Office Small Business 2007"
			Case "10D7" oProd = "Microsoft Office InfoPath Forms Services"
			Case "110D" oProd = "Microsoft Office SharePoint Server 2007"
			Case "1122" oProd = "Windows SharePoint Services Developer Resources 1.2"
			Case "0010" oProd = "SKU - Microsoft Software Update for Web Folders (English) 12"
		End Select
		oEdit = oProd
	End If

	' oEdit was extracted from DigitalProductID binary value offset.
	' We use this approach to detect the Office edition because otherwise "Office Home and Business 2010" edition will be detected as
	' "Office Single Image 2010" (Msi ProductName) by the next function.
	If (oVer = "2010" Or oVer = "2013") Then
		Select Case oEdit
			Case "ProjectStdVL"    oProd = "Microsoft Office Project Standard " & oVer & " (VL)"
			Case "ProjectProVL"    oProd = "Microsoft Office Project Professional " & oVer & " (VL)"
			Case "ProjectProMSDNR" oProd = "Microsoft Project Professional " & oVer & " (MSDN)"
			Case "HomeBusinessR"   oProd = "Microsoft Office Home and Business " & oVer
			Case "ProfessionalR"   oProd = "Microsoft Office Professional " & oVer
			Case "ProPlusR"        oProd = "Microsoft Office Professional Plus " & oVer
			Case "StandardR"       oProd = "Microsoft Office Standard " & oVer
			Case "StandardVL"      oProd = "Microsoft Office Standard " & oVer & " (VL)"
			Case "HomeStudentR"    oProd = "Microsoft Office Home and Student " & oVer
			Case "AccessRuntimeR"  oProd = "Microsoft Office Access Runtime " & oVer
			Case "VisioSIR"        oProd = "Microsoft Office Visio Professional " & oVer
			Case "SPDR"            oProd = "Microsoft SharePoint Designer " & oVer
			Case "ProjectProR"     oProd = "Microsoft Project Professional " & oVer
			Case "ProjectStdR"     oProd = "Microsoft Project Standard " & oVer
			Case "VisioSIVL"       oProd = "Microsoft Visio " & oVer & " Standard (VL)"
			Case "InfoPathR"       oProd = "Microsoft Office InfoPath " & oVer
			Case Else              oProd = "Microsoft Office Unknown Edition " & oVer & ": " & oEdit    
		End Select
			
		' WScript.Echo "InStr(" & oProd & ", Microsoft Office Unknown Edition 2010): " & InStr(oProd, "Microsoft Office Unknown Edition 2010")
		If InStr(oProd, "Microsoft Office Unknown Edition 2010") Then
			hProd = Mid ((Right(oGUID,38)),11,4)
			' WScript.Echo "hProd: " & hProd
			Select Case hProd
				Case "0011" oProd = "Microsoft Office Professional Plus 2010"
				Case "0012" oProd = "Microsoft Office Standard 2010"
				Case "0013" oProd = "Microsoft Office Home and Business 2010"
				Case "0014" oProd = "Microsoft Office Professional 2010"
				Case "0015" oProd = "Microsoft Access 2010"
				Case "0016" oProd = "Microsoft Excel 2010"
				Case "0017" oProd = "Microsoft SharePoint Designer 2010"
				Case "0018" oProd = "Microsoft PowerPoint 2010"
				Case "0019" oProd = "Microsoft Publisher 2010"
				Case "001A" oProd = "Microsoft Outlook 2010"
				Case "001B" oProd = "Microsoft Word 2010"
				Case "001C" oProd = "Microsoft Access Runtime 2010"
				Case "001F" oProd = "Microsoft Office Proofing Tools Kit Compilation 2010"
				Case "002F" oProd = "Microsoft Office Home and Student 2010"
				Case "003A" oProd = "Microsoft Project Standard 2010"
				Case "003D" oProd = "Microsoft Office Single Image 2010"
				Case "003B" oProd = "Microsoft Project Professional 2010"
				Case "0044" oProd = "Microsoft InfoPath 2010"
				Case "0052" oProd = "Microsoft Visio Viewer 2010"
				Case "0057" oProd = "Microsoft Visio 2010"
				Case "007A" oProd = "Microsoft Outlook Connector"
				Case "008B" oProd = "Microsoft Office Small Business Basics 2010"
				Case "00A1" oProd = "Microsoft OneNote 2010"
				Case "00AF" oProd = "Microsoft PowerPoint Viewer 2010"
				Case "00BA" oProd = "Microsoft Office SharePoint Workspace 2010"
				Case "110D" oProd = "Microsoft Office SharePoint Server 2010"
				Case "110F" oProd = "Microsoft Project Server 2010"
			End Select
		End If
		
		If InStr(oProd, "Microsoft Office Unknown Edition 2013") Then
			hProd = Mid ((Right(oGUID,38)),11,4)
			Select Case hProd
				Case "0011" oProd = "Microsoft Office Professional Plus 2013"
				Case "0012" oProd = "Microsoft Office Standard 2013"
				Case "0013" oProd = "Microsoft Office Home and Business 2013"
				Case "0014" oProd = "Microsoft Office Professional 2013"
				Case "0015" oProd = "Microsoft Access 2013"
				Case "0016" oProd = "Microsoft Excel 2013"
				Case "0017" oProd = "Microsoft SharePoint Designer 2013"
				Case "0018" oProd = "Microsoft PowerPoint 2013"
				Case "0019" oProd = "Microsoft Publisher 2013"
				Case "001A" oProd = "Microsoft Outlook 2013"
				Case "001B" oProd = "Microsoft Word 2013"
				Case "001C" oProd = "Microsoft Access Runtime 2013"
				Case "001F" oProd = "Microsoft Office Proofing Tools Kit Compilation 2013"
				Case "002F" oProd = "Microsoft Office Home and Student 2013"
				Case "003A" oProd = "Microsoft Project Standard 2013"
				Case "003B" oProd = "Microsoft Project Professional 2013"
				Case "0044" oProd = "Microsoft InfoPath 2013"
				Case "0051" oProd = "Microsoft Visio Professional 2013"
				Case "0053" oProd = "Microsoft Visio Standard 2013"
				Case "00A1" oProd = "Microsoft OneNote 2013"
				Case "00BA" oProd = "Microsoft Office SharePoint Workspace 2013"
				Case "110D" oProd = "Microsoft Office SharePoint Server 2013"
				Case "110F" oProd = "Microsoft Project Server 2013"
				Case "012B" oProd = "Microsoft Lync 2013"
			End Select
		End If
	End If ' If (oVer = "2010" Or oVer = "2013")
	
	' Office XP/2003/2007
	If oInstall = "0" Then
		oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Windows\CurrentVersion\Uninstall\" & oGUID, "DisplayName", oProdTemp
		If Not IsNullOrEmpty(oProdTemp) Then
			If IsNullOrEmpty(oProd) Or InStr(oProd, "Microsoft Office Unknown Edition") Then
				oProd = oProdTemp
			End If
			oInstall = "1"
		End If
	End If
	
	' Office 2010/2013/2016
	If oInstall = "0" Then
	kEdit = UCase(oEdit)
		
	If Mid(oGUID,11,4) = "003D" Then
		kEdit = "SingleImage"
	End If
		
	oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Windows\CurrentVersion\Uninstall\Office" & Left(aOffID(a,1),2) & "." & kEdit, "DisplayName", oProdTemp
		If Not IsNullOrEmpty(oProdTemp) Then
			If IsNullOrEmpty(oProd) Or InStr(oProd, "Microsoft Office Unknown Edition") Then
				oProd = oProdTemp
			End If
			oInstall = "1"
		End If
	End If
	
	' Office 2019
	If IsNullOrEmpty(oProd) Then
		oReg.EnumKey HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall", aUninstallKeys
		If Not IsNull(aUninstallKeys) Then
			For Each UninstallKey In aUninstallKeys		 
				oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & UninstallKey, "UninstallString", sValue, "REG_SZ"
				If InStr(LCase(sValue), "microsoft office " & Left(aOffID(a,1),2)) > 0 OR _
				  (InStr(LCase(sValue), "productstoremove=") > 0 AND _
				  InStr(sValue, "." & Left(aOffID(a,1),2) & "_") > 0) Then
					oNote = Mid(sValue, InStr(sValue, "productstoremove="))
					oNote = Replace(oNote, "productstoremove=","")
					oNote = Left(oNote, InStr(oNote, ".1") - 1) ' E.g: HomeBusiness2019Retail
						
					oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" & UninstallKey, "DisplayName", oProd, "REG_SZ"
					
					oInstall = "1"
				End If
			Next
		End If
	End If

	' Office isn't installed.
	' oProd will be Empty only if this script cannot determine the Office edition from the oGUID.
	' Querying the value from "ProductName", "ConvertToEdition", "ProductNameNonQualified" or "ProductNameVersion" isn't a reliable method to detect the installed Office edition because it doesn't match the real Office edition installed into the machine. E.g: "Office Home and Business 2010" are reported as "Microsoft Office Professional 2010". 
	If IsNullOrEmpty(oProd) Then
		oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey, "ProductName", oProd
		If IsNullOrEmpty(oProd) Then oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey, "ConvertToEdition", oProd
	End If
	
	' WScript.Echo "oProd: " & oProd
	' WScript.Echo "IsNullOrEmpty(oProd): " & IsNullOrEmpty(oProd)
	' WScript.Echo "IsEmpty(oOfficeOSPPInfos): " & IsEmpty(oOfficeOSPPInfos)
	If IsNullOrEmpty(oProd) And Not (IsEmpty(oOfficeOSPPInfos)) Then
		oProd = oOfficeOSPPInfos.Item("LicenseName")
	End If
	If IsNullOrEmpty(oProd) Then oProd = "Unidentifiable Office " & oVer:Else:oProd = oProd & OEM:End If
	
	' WScript.Echo "oNote (" & Len(oNote) & " chars): " & oNote
	' If Len(oNote) = 0 Then oNote = GetProductReleaseIdFromPrimaryProductId(CInt(Left(aOffID(a,1),2)), Mid(oGUID,11,4))
	writeXML oVer,oProd,oProdID,oBit,oGUID,oInstall,oKey,oNote
End Sub

Sub writeXML(oVer,oProd,oProdID,oBit,oGUID,oInstall,oKey,oNote)
	Wscript.Echo _
	"<OFFICEPACK>" & vbCrLf & _
	"<OFFICEVERSION>" & oVer & "</OFFICEVERSION>" & vbCrLf & _
	"<PRODUCT>" & oProd & "</PRODUCT>" & vbCrLf & _
	"<PRODUCTID>" & oProdID & "</PRODUCTID>" & vbCrLf & _
	"<TYPE>" & oBit & "</TYPE>" & vbCrLf & _
	"<OFFICEKEY>" & oKey & "</OFFICEKEY>" & vbCrLf & _
	"<GUID>" & oGUID & "</GUID>" & vbCrLf & _
	"<INSTALL>" & oInstall & "</INSTALL>" & vbCrLf & _
	"<NOTE>" & oNote & "</NOTE>" & vbCrLf & _
	"</OFFICEPACK>"
End Sub

Function getOfficeOSPPInfos(version)
	Dim WshShell, oExec
	Dim mTab
	Dim key, value
	Dim path
	Dim writeProduct
	Dim objOfficeDict
		
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Set WshShellObj = WScript.CreateObject("WScript.Shell") 
	Set WshProcessEnv = WshShellObj.Environment("Process") 
	Set objOfficeDict = CreateObject("Scripting.Dictionary") 

	result = WshShell.Run("cmd /c cscript ""C:\Program Files (x86)\Microsoft Office\Office" & version & "\OSPP.VBS"" /dstatus > %USERPROFILE%\output.txt", 0, true)
	' Debug : if 32 bits version available ?
	' wscript.echo result

	' If file not there command throw an error and return is 1 and abover
	If result > 0 then
		' Try with the 64 bits version if available
		result = WshShell.Run("cmd /c cscript ""C:\Program Files\Microsoft Office\Office" & version & "\OSPP.VBS"" /dstatus > %USERPROFILE%\output.txt", 0, true)
		' Debug : if 64 bits version available ?
		' WScript.Echo result
	End If

	' Result = 0 if successfully executed
	If result = 0 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		' The USERNAME env var doesn't take the domain part into account which leads to a wrong directory which cannot be read
		path = WshProcessEnv("USERPROFILE") & "\output.txt"

		Set file = fso.OpenTextFile(path, 1)
		'strData = file.ReadLine
		' writeProduct = 0
		Do Until file.AtEndOfStream
			' Debug : echo each line 
			' WScript.echo file.ReadLine
			
			str = file.ReadLine
			' Debug : Show string before split 
			' WScript.Echo str
			
			mTab = Split(str, ":")
			arrCount = uBound(mTab) + 1
			
			If arrCount > 1 then
				key = mTab(0)
				value = mTab(1)
				
				' PRODUCT ID: XXXXX-XXXXX-XXXXX-XXXXX
				' SKU ID: 7fe09eef-5eed-4733-9a60-d7019df11cac
				' LICENSE NAME: Office 19, Office19HomeBusiness2019R_Retail edition
				' LICENSE DESCRIPTION: Office 19, RETAIL channel
				' BETA EXPIRATION: 01/01/1601
				' LICENSE STATUS:  ---LICENSED---
				' Last 5 characters of installed product key: R62YT
				
				Select Case key
				Case "PRODUCT ID"
					writeProduct = 1
					oProdID = Trim(mTab(1))
					objOfficeDict.Add "ProductID", oProdID
					' Debug : echo office data
					' WScript.echo "oProdId = " & oProdID
				Case "SKU ID"
					writeProduct = 1
					oGUID = Trim(mTab(1))
					objOfficeDict.Add "SKUID", oGUID
					' Debug : echo office data
					' WScript.echo "oGUID = " & oGUID
				Case "LICENSE NAME"
					oProd = Trim(mTab(1))
					objOfficeDict.Add "LicenseName", oProd
					' Debug : echo office data
					' WScript.echo "oProd = " & oProd
				Case "LICENSE DESCRIPTION"
					oVer = Trim(mTab(1))
					objOfficeDict.Add "LicenseDescription", oProd
					' Debug : echo office data
					' WScript.echo "oVer = " & oVer
				Case "ERROR DESCRIPTION"
					oNote = Trim(mTab(1))
					objOfficeDict.Add "ErrorDescription", oNote
					' Debug : echo office data
					' WScript.echo "oNote = " & oNote
				Case "Last 5 characters of installed product key"
					oKey = "XXXXX-XXXXX-XXXXX-XXXXX-" & Trim(mTab(1))
					objOfficeDict.Add "PartialProductKey", oKey
					' Debug : echo office data
					' WScript.echo "oKey = " & oKey                                
				End Select    
			Else
				' If writeProduct = 1 Then
				'	 oInstall = 1                                                               
				'	 oBit = 1
					 ' Check if Office is 365                                                    
				'	 If InStr(oProd, "O365") > 0 Then                                           
				'		oVer = "365"                                                  
				'		oProd = Right(oProd, len(oProd)-11)
				'		If oProd = " Office16O365BusinessR_Subscription edition" Then oProd = "Microsoft Office Business Subscription Edition 365" : End If
				'		If oProd = " Office16O365BusinessR_Grace edition" Then oProd = "Microsoft Office Business Grace Edition 365" : End If
				'	 End if                                                                     
				'	 writeXML oVer,oProd,oProdID,oBit,oGUID,oInstall,oKey,oNote
				'	writeProduct=0   
				' End If
			End If                                                                 
		Loop
		file.Close                                                                         
	End If      
	Set getOfficeOSPPInfos = objOfficeDict
End Function                                                                            

Function decodeKey(iValues)                                                        
	Dim arrDPID, foundKeys                                                           
	arrDPID = Array()
	foundKeys = Array()                                                              

	Select Case (UBound(iValues))
		Case 255:  ' 2000
			range = Array(52,66)
		Case 163:  ' XP, 2003, 2007
			range = Array(52,66)
		Case 1271: ' 2010, 2013
			range = Array(808,822)
		Case Else
			Exit Function
	End Select

	charset = "BCDFGHJKMPQRTVWXY2346789"

	For i = range(0) to range(1)
		ReDim Preserve arrDPID( UBound(arrDPID) + 1 )
		arrDPID( UBound(arrDPID) ) = iValues(i)
	Next

	withN = (arrDPID(UBound(arrDPID)) \ 6) And 1
	arrDPID(UBound(arrDPID)) = (arrDPID(UBound(arrDPID)) And &HF7) Or ((withN And 2) * 4)

	For i = 24 To 0 Step -1
		k = 0
		For j = 14 To 0 Step -1
			k = k * 256 Xor arrDPID(j)
			arrDPID(j) = k \ 24
			k = k Mod 24
		Next
		strProductKey = Mid(charset, k+1, 1) & strProductKey
	Next

	If (withN = 1) Then
		keypart = Mid(strProductKey,2,k)
		strProductKey = Replace(strProductKey, keypart, keypart & "N", 2, 1, 0)
		If k = 0 Then strProductKey = "N" & strProductKey
	End If

	decodeKey = ""
	For i = 1 To 25
		decodeKey = decodeKey & Mid(strProductKey,i,1)
		If i Mod 5 = 0 And i <> 25 Then decodeKey = decodeKey & "-"
	Next
End Function

Function IsNullOrEmpty(strValue)
	If IsNull(strValue) Then
		IsNullOrEmpty = True
	ElseIf IsEmpty(strValue) Then
		IsNullOrEmpty = True
	ElseIf (strValue = "") Then
		IsNullOrEmpty = True
	ElseIf (Len(strValue) = 0) Then
		IsNullOrEmpty = True
	End If
End Function
