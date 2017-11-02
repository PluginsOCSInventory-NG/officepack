'----------------------------------------------------------
' Plugin for OCS Inventory NG 2.x
' Script :		Retrieve Microsoft Office informations
' Version :		2.24
' Date :		14/06/2017
' Author :		Creative Commons BY-NC-SA 3.0
' Author :		Nicolas DEROUET (nicolas.derouet[gmail]com)
' Contributor :	Stéphane PAUTREL (acb78.com)
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
' NOTE : No checked on Windows 8
' Included : Office 2016 and 365 versions
' ---------------------------------------------------------
On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002

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
  schKey "SOFTWARE\Wow6432Node\Microsoft\Office\" & aOffID(a,1) & "\Registration", false
  schKey "SOFTWARE\Microsoft\Office\" & aOffID(a,1) & "\Registration", true
Next

getOffice16Infos

Sub schKey97(regKey)
  oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Office\8.0", "BinDirPath", oDir97
  If IsNull(oDir97) Then Exit Sub
  oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Microsoft Reference\BookshelfF\96L", "PID", oProdID
  oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey & "Windows\CurrentVersion\Uninstall\Office8.0", "DisplayName", oProd
  oInstall = "1"
  If IsNull(oProd) Then
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

  If IsNull(oProd) Then oProd = "Microsoft " & name & " 2000"
  writeXML "2000",oProd,oProdID,32,oGUID,oInstall,oKey,""
End Sub

Sub schKey(regKey, likeOS)
  oReg.GetBinaryValue HKEY_LOCAL_MACHINE, regKey, "DigitalProductID", aDPIDBytes
  If IsNull(aDPIDBytes) Then
    oReg.EnumKey HKEY_LOCAL_MACHINE, regKey, aGUIDKeys
    If Not IsNull(aGUIDKeys) Then
      For Each GUIDKey In aGUIDKeys
        schKey regKey & "\" & GUIDKey, likeOS
      Next
    End If
  Else
    oVer = aOffID(a,0)
    oProd = Null
    oKey = decodeKey(aDPIDBytes)
    oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey, "ProductID", oProdID
    oBit = osType
    If Not likeOS Then oBit = 32
    oGUID = Right(regKey,InStr(StrReverse(regKey),"\")-1)
    oInstall = "1"
    wow = ""
    If Not likeOS Then wow = "WOW6432Node\"

    oEdit = ""
    If (oVer = "2010" Or oVer = "2013") Then
      For i = 280 to 320 Step 2
        If aDPIDBytes(i) <> 0 Then oEdit = oEdit & Chr(aDPIDBytes(i))
      Next
    End If
    oNote = oEdit

    If IsNull(oProd) And (oVer = "2010" Or oVer = "2013" Or oVer = "2016") Then
      kEdit = UCase(oEdit)
      If Mid(oGUID,11,4) = "003D" Then kEdit = "SingleImage"
      oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Windows\CurrentVersion\Uninstall\Office" & Left(aOffID(a,1),2) & "." & kEdit, "DisplayName", oProd
    End If

    If IsNull(oProd) Then _
      oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Windows\CurrentVersion\Uninstall\" & oGUID, "DisplayName", oProd

    If IsNull(oProd) Then
      oInstall = "0"
      oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey, "ProductName", oProd
      If IsNull(oProd) Then oReg.GetStringValue HKEY_LOCAL_MACHINE, regKey, "ConvertToEdition", oProd

      ' Office Visio XP
      If IsNull(oProd) And (oVer = "XP") Then
        oReg.GetStringValue HKEY_LOCAL_MACHINE, "Software\" & wow & "Microsoft\Office\XP\Common\ProductVersion", "LastProduct", pVer
        ' Original / SP1 / SP2
        If ((pVer = "10.0.525") Or (pVer = "10.1.2514") Or (pVer = "10.2.5110")) Then
          oProd = "Microsoft Office Visio XP"
        End If
      End If

      ' Office Visio Viewer 2003
      If IsNull(oProd) And (oVer = "2003") And (oKey = "MF4QD-3T4PM-26X66-4KH7R-QGTYT") Then
        oProd = "Microsoft Office Visio Viewer 2003"
      End If

      If IsNull(oProd) Then oProd = "Unidentifiable Office " & oVer
    End If
    writeXML oVer,oProd,oProdID,oBit,oGUID,oInstall,oKey,oNote
  End If
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

Sub getOffice16Infos
	Dim WshShell, oExec
	Dim mTab
	Dim key, value
	Set WshShell = WScript.CreateObject("WScript.Shell")
    Set WshShellObj = WScript.CreateObject("WScript.Shell") 
    Set WshProcessEnv = WshShellObj.Environment("Process") 

	result = WshShell.Run("cmd /c cscript ""C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS"" /dstatus > %USERPROFILE%\output.txt", 0, true)
	' Debug : if 32 bits version available ?
	' wscript.echo result

	' If file not there command throw an error and return is 1 and abover
	if result > 0 then
		' Try with the 64 bits version if available
		result = WshShell.Run("cmd /c cscript ""C:\Program Files\Microsoft Office\Office16\OSPP.VBS"" /dstatus > %USERPROFILE%\output.txt", 0, true)
		' Debug : if 64 bits version available ?
		' WScript.Echo result
	end if

	' Result = 0 if successfully executed
	if result = 0 then
		Set fso  = CreateObject("Scripting.FileSystemObject")
		Set file = fso.OpenTextFile("C:\Users\" & WshProcessEnv("USERNAME") & "\output.txt", 1)
		'strData = file.ReadLine
		
		Do Until file.AtEndOfStream
			' Debug : echo each line 
			' WScript.echo file.ReadLine
			
			str = file.ReadLine
			' Debug : Show string before split 
			' WScript.Echo str
			
			mTab = Split(str, ":")
			arrCount = uBound(mTab) + 1
			
			if arrCount > 1 then
				key = mTab(0)
				value = mTab(1)
				
				Select Case key
				Case "PRODUCT ID"
					oProdID = mTab(1)
					' Debug : echo office data
					' WScript.echo "oProdId = " & oProdID
				Case "SKU ID"
					oGUID = mTab(1)
					' Debug : echo office data
				    ' WScript.echo "oGUID = " & oGUID
				Case "LICENSE NAME"
					oProd = mTab(1)
					' Debug : echo office data
					' WScript.echo "oProd = " & oProd
				Case "LICENSE DESCRIPTION"
					oVer = mTab(1)
					' Debug : echo office data
					' WScript.echo "oVer = " & oVer
				Case "ERROR DESCRIPTION"
					oNote = mTab(1)
					' Debug : echo office data
					' WScript.echo "oNote = " & oNote
				Case "Last 5 characters of installed product key"
					oKey = "XXXXX-XXXXX-XXXXX-XXXXX-" & mTab(1)
					' Debug : echo office data
					' WScript.echo "oKey = " & oKey                                
				End Select                                                         
			End If                                                                 

		Loop

		oInstall = 1                                                               
		oBit = 1

		file.Close                                                                 

        'Check if Office is 365                                                    
        If InStr(oProd, "O365") > 0 Then                                           
            oVer = "365"                                                  
            oProd = Right(oProd, len(oProd)-11)
			If oProd = " Office16O365BusinessR_Subscription edition" Then oProd = "Microsoft Office Business Subscription Edition 365" : End If
			If oProd = " Office16O365BusinessR_Grace edition" Then oProd = "Microsoft Office Business Grace Edition 365" : End If
        End if                                                                     
		writeXML oVer,oProd,oProdID,oBit,oGUID,oInstall,oKey,oNote                 
	End If                                                                         

End Sub                                                                            

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
