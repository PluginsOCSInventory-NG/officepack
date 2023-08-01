<#
.Synopsis
----------------------------------------------------------
 Plugin for OCS Inventory NG 2.x
 Script :		Retrieve Microsoft Office informations
 Version :		2.25
 Date :		19/07/2023
 Author :		Creative Commons BY-NC-SA 3.0
 Author :		Nicolas DEROUET (nicolas.derouet[gmail]com)
 Contributor :		Stéphane PAUTREL (acb78.com), Eduardo Mozart de Oliveira, Antoine ROBIN
----------------------------------------------------------
 OS checked [X] on	32b	64b	(Professionnal edition)
	Windows XP	    [X]
	Windows Vista	[X]	[X]
	Windows 7	    [X]	[X]
	Windows 8.1	    [X]	[X]	
	Windows 10	    [X]	[X]
	Windows 2k8R2		[X]
	Windows 2k12R2		[X]
	Windows 2k16		[X]
 ---------------------------------------------------------
 Note :	No checked on Windows 8
 Included :	Office 2016 and 365 versions
 ---------------------------------------------------------

.Description
Retrieve Microsoft Office informations

#>

###
# Functions
###
Function EnumKey
{
    param
    (
      [UInt32]
      $hDefKey,

      [String]
      $sSubKeyName
    )

    $keys = (Invoke-CimMethod -ClassName StdRegProv -MethodName EnumKey -Arguments $PSBoundParameters).sNames
    Return $keys
}

Function GetBinaryValue
{
    param
    (
        [UInt32]
        $hDefKey,

        [String]
        $sSubKeyName,

        [String]
        $sValueName
    )

    $value = (Invoke-CimMethod -ClassName StdRegProv -MethodName GetBinaryValue -Arguments $PSBoundParameters).sValue

    if({GetBinaryValue -hDefKey $hDefKey -sSubKeyName $sSubKeyName -sValueName $sValueName} -ine $null)
    {
        Return $value
    }
    else
    {
        Return ''
    }
}

Function GetStringValue
{
    param
    (
        [UInt32]
        $hDefKey,

        [String]
        $sSubKeyName,

        [String]
        $sValueName
    )

    $value = (Invoke-CimMethod -ClassName StdRegProv -MethodName GetStringValue -Arguments $PSBoundParameters).sValue

    if({GetStringValue -hDefKey $hDefKey -sSubKeyName $sSubKeyName -sValueName $sValueName} -ine $null)
    {
        Return $value
    }
    else
    {
        Return ''
    }
}

Function decodeKey
{
    Param(
        [string]
        $iValues
    )

    $arrDPID = ''
    $foundKeys = ''
    switch ($iValues.Length-1)
    {
        255 # 2000
        {
            $range = 52,66
        }
		163 # XP, 2003, 2007
        {
			$range = 52,66
        }
		1271 # 2010, 2013
        {
			$range = 808,822
        }
		default
        {
            Return
        }
    }
    $charset = "BCDFGHJKMPQRTVWXY2346789"
    For ($i = $range[0]; $i -lt $range[1]; $i++)
    {
        $arrDPID = $arrDPID + $iValues[$i]
    }
    $withN = $arrDPID[$arrDPID.Length-1] / 6 -and 1
    $arrDPID[$arrDPID.Length-1] = $arrDPID[$arrDPID.Length-1] -and '&HF7' -or ($withN -and 2) * 4
    For ($i = 24; $i -gt 0; $i--)
    {
        $k = 0
        For ($j = 14; $j -gt 0; $j--)
        {
            $k = $k * 256 -xor $arrDPID[$j]
            $arrDPID[$j] = $k / 24
            $k = $k % 24
        }
        $strPorductKey = $charset.Substring(2,$k) + $strPorductKey
    }
    If ($withN -eq 1)
    {
        $keypart = $strPorductKey.Substring(2,$k)
        $strPorductKey = $strPorductKey -replace($keypart, $keypart + 'N', 2, 1, 0)
        If ($k = 0)
        {
            $strPorductKey = 'N' + $strPorductKey
        }
    }
    For ($i = 1; $i -lt 25; $i++)
    {
        $decodeKey = $decodeKey + $strPorductKey.Substring($i, 1)
        If ($i % 5 -eq 0 -and $i -ine 25)
        {
            $decodeKey = $decodeKey + '-'
        }
    }
}

Function getOfficeOSPPInfos
{
    Param (
        [string]
        $version
    )

    $objOfficeDict = @()
    $objOffice = @{}
    $index = 0
    $path = 'C:\Program Files (x86)\Microsoft Office\Office' + $version + '\OSPP.VBS'  
    If ($osType -eq '64')
    {
        $path = 'C:\Program Files\Microsoft Office\Office' + $version + '\OSPP.VBS'
    }
    cscript.exe  $path /dstatus > $env:USERPROFILE\output.txt
    If (Test-Path $env:USERPROFILE\output.txt -PathType Leaf)
    {
        $data = Get-Content $env:USERPROFILE\output.txt
        Foreach ($str in $data)
        {
            $key = $str.Split(':')[0]
            $value = $str.Split(':')[1]
            switch ($key)
            {
                'PRODUCT ID'
                {
                    $productID = 1
                    $writeProduct = 1
                    $oProdID = $value.Trim()
                    $objOffice.Add('ProductID', $oProdID)
                }
                'SKU ID'
                {
                    $SKUID = 1
                    $writeProduct = 1
                    $oGUID = $value.Trim()
                    $objOffice.Add('SKUID', $oGUID)
                }
                'LICENSE NAME'
                {
                    $LicenseName = 1
                    $writeProduct = 1
                    $oProd = $value.Trim()
                    $objOffice.Add('LicenseName', $oProd)
                }
                'LICENSE DESCRIPTION'
                {
                    $LicenseDescription = 1
                    $writeProduct = 1
                    $oVer = $value.Trim()
                    $objOffice.Add('LicenseDescription', $oVer)
                }
                'ERROR DESCRIPTION'
                {
                    $writeProduct = 1
                    $oNote = $value.Trim()
                    $objOffice.Add('ErrorDescription', $oNote)
                }
                'Last 5 characters of installed product key'
                {
                    $ProductKey = 1
                    $writeProduct = 1
                    $oKey = 'XXXXX-XXXXX-XXXXX-XXXXX-' + $value.Trim()
                    $objOffice.Add('PartialProductKey', $oKey)
                }
            }
            If (
                ($productID -eq 1) -and
                ($SKUID -eq 1) -and
                ($LicenseName -eq 1) -and
                ($LicenseDescription -eq 1) -and
                ($ProductKey -eq 1)
            )
            {
                $productID = 0
                $SKUID = 0
                $LicenseName = 0
                $LicenseDescription = 0
                $ProductKey = 0
                $objOfficeDict += $objOffice
                $objOffice = @{}
            }
        }
    }
    Return $objOfficeDict
}

Function schKey97
{
    Param(
        [string]$regKey
    )

    $oDir97 = GetStringValue -hDefKey '0x80000002' -sSubKeyName {$regKey + 'Office\8.0'} -sValueName 'BinDirPath'
    If ($oDir97.Length -eq 0)
    {
        Return
    }
    $oProdID = GetStringValue -hDefKey '0x80000002' -sSubKeyName {$regKey + 'Microsoft Reference\BookshelfF\96L'} -sValueName 'PID'
    $oProd = GetStringValue -hDefKey '0x80000002' -sSubKeyName {$regKey + 'Windows\CurrentVersion\Uninstall\Office8.0'} -sValueName 'DisplayName'
	$oInstall = '1'
	If ($oProd.Length -eq 0)
    {
        $oInstall = '0'
	    $oProd = 'Microsoft Office 97'
    }
	writeXML -oVer '97' -oProd $oProd -oProdID $oProdID -oBit 32 -oGUID '' -oInstall $oInstall -oKey '' -oNote ''
}

Function schKey2K
{
    Param(
        [string]$name,
        
        [string]$regKey,

        [array]$guid1,

        [string]$guid2
    )

	$oInstall = '0'
    $aDPIDBytes = GetStringValue -hDefKey '0x80000002' -sSubKeyName {$regKey + 'Registration\DigitalProductID'}
	$oKey = ''
	If ($aDPIDBytes -ine $null)
    {
        $oKey = decodeKey($aDPIDBytes)
    }
    $oProdID = GetStringValue -hDefKey '0x80000002' -sSubKeyName {$regKey + 'Registration\ProductID'}
	If ($oProdID.Length -eq 0)
    {
        Return
    }
    $aKeys = EnumKey -hDefKey '0x80000002' -sSubKeyName {$regKey + 'Microsoft\Windows\CurrentVersion\Uninstall\'}
    If ($aKeys.Length -gt 0)
    {
        ForEach ($guid in $aKeys)
        {
            Write-Output $guid
            If ($guid.Substring($guid, $guid.Length - $guid.indexof('-')).ToUpper() -eq $guid2 + '}')
            {
                For ($i = 0; $i -lt $guid1.Length; $i++)
                {
                   If ($guid.Substring($guid, $guid1[$i].Length+1) -eq '{' + $guid1[$i])
                   {
                        $oProd = GetStringValue -hDefKey '0x80000002' -sSubKeyName {'Software\' + $wow + 'Microsoft\Windows\CurrentVersion\Uninstall\' + $guid} -sValueName 'DisplayName'
                        $oGUID = $guid
                        $oInstall = '1'
                   }
                }
            }
        }
    }
    If ($oProd.Length -eq 0)
    {
        $oProd = 'Microsoft' + $name + ' 2000'

    }
	writeXML -oVer '2000' -oProd $oProd -oProdID $oProdID -oBit 32 -oGUID $oGUID -oInstall $oInstall -oKey $oKey -oNote ''
}

Function schKey
{
    param (
        [string]$regKey,

        [string]$likeOS,

        [int]$version
    )

    $aGUIDKeys = EnumKey -hDefKey '0x80000002' -sSubKeyName $regKey
    If ($aGUIDKeys -ine $null)
    {
        Foreach ($GUIDKey in $aGUIDKeys)
        {
            If (
                $OFFICE_ALL.Contains($GUIDKey.Substring($GUIDKey.Length-28, 28).ToUpper() -gt 0 -or
                $OFFICEID.Contains($GUIDKey.Substring($GUIDKey.Length-17, 17).ToUpper() -gt 0))
            )
            {
                schKey -regKey {$regKey + '\' + $GUIDKey} -likeOS $likeOS
                Return
            }
        }
    }
    Else
    {
        Return
    }
    $aDPIDBytes = GetBinaryValue -hDefKey '0x80000002' -sSubKeyName $regKey -sValueName 'DigitalProductID'
    If ($aDPIDBytes.Length -eq 0)
    {
        $aOSPPVersions = $aOffID[$version][1].Split('.')
        $oOfficeOSPPInfos = getOfficeOSPPInfos -version $aOSPPVersions[0]
        If ($oOfficeOSPPInfos.Length -ine 0)
        {
            $oProdID = $oOfficeOSPPInfos['ProductID']
            $oSKUID = $oOfficeOSPPInfos['SKUID']
            If ($oOfficeOSPPInfos['ErrorDescription'])
            {
                $oNote = $oOfficeOSPPInfos['ErrorDescription']
            }
            Else
            {
                $oNote = ''
            }
            $oKey = $oOfficeOSPPInfos['PartialProductKey']
        }
    }
    Else
    {
        $oKey = decodeKey -iValues $aDPIDBytes
        $oProdID = GetStringValue -hDefKey '0x80000002' -sSubKeyName $regKey -sValueName 'ProductID'
    }
    If($oProdID.Length -gt 0)
    {
        If ($oProdID.SubString(7, 3) -eq 'OEM')
        {
            $oOEM = ' OEM'
        }
        $oVer = $aOffID[$version][0]
        $oEdit = ''
        If ($oVer -eq '2010')
        {
            $uBoundDPIDEdit = 312
        }
        Else
        {
            $uBoundDPIDEdit = 320
        }
        For ($i = 280; $i -lt $uBoundDPIDEdit; $i++)
        {
            If ($aDPIDBytes -ine $null)
            {
                $oEdit = $oEdit + [char]$aDPIDBytes[$version]
            }
        }
        $oNote = $oEdit
        If ($oVer -eq '2016')
        {
            If ($aDPIDBytes.Length -eq 0)
            {
                $oVer = '2019'
            }
        }
        for ($i = $regKey.Length - 1; $i -gt 0; $i--)
        {
            $reverseRegKey = $reverseRegKey + ($regKey.Substring($i,1))
        }
        $oGUID = $regKey.Substring($regKey.Length-$reverseRegKey.IndexOf('\')-1, $reverseRegKey.IndexOf('\'))
        If($oSKUID.Length -gt 0)
        {
            $oGUID = $oSKUID
        }
        $oBit = $osType
        If (-not $likeOS)
        {
            If ([Int]$aOffID[$version][1].Substring($aOffID, 2) -gt 11)
            {
                If ($oGUID.Substring(21, 1) -eq '0')
                {
                    $oBit = 32
                }
                Else
                {
                    If ($oGUID.Substring(20,1) -eq '1')
                    {
                        $oBit = 64
                    }
                }
            }
        }
        $oInstall = '0'
        If (-not $likeOS)
        {
            $wow = 'WOW6432Node'
        }
        If ($oVer -eq 'XP')
        {
            $hProd = $oGUID.Substring($oGUID.Length-38,38).Substring(4,2)
            switch ($hProd)
            {
                '11'
                {
                    $oProd = 'Microsoft Office XP Professional'
                }
                '12'
                {
                    $oProd = "Microsoft Office XP Standard"
                }
                '13'
                {
                    $oProd = "Microsoft Office XP Small Business"
                }
                '14'
                {
                    $oProd = "Microsoft Office XP Web Server"
                }
                '15'
                {
                    $oProd = "Microsoft Access 2002"
                }
                '16'
                {
                    $oProd = "Microsoft Excel 2002"
                }
                '17'
                {
                    $oProd = "Microsoft FrontPage 2002"
                }
                '18'
                {
                    $oProd = "Microsoft PowerPoint 2002"
                }
                '19'
                {
                    $oProd = "Microsoft Publisher 2002"
                }
                '1A'
                {
                    $oProd = "Microsoft Outlook 2002"
                }
                '1B'
                {
                    $oProd = "Microsoft Word 2002"
                }
                '1C'
                {
                    $oProd = "Microsoft Access 2002 Runtime"
                }
                '27'
                {
                    $oProd = "Microsoft Project 2002"
                }
                '28'
                {
                    $oProd = "Microsoft Office XP Professional with FrontPage"
                }
                '31'
                {
                    $oProd = "Microsoft Project 2002 Web Client"
                }
                '32'
                {
                    $oProd = "Microsoft Project 2002 Web Server"
                }
                '3A'
                {
                    $oProd = "Project 2002 Standard"
                }
                '3B'
                {
                    $oProd = "Project 2002 Professional"
                }
                '51'
                {
                    $oProd = "Microsoft Office Visio Professional 2002"
                }
                '54'
                {
                    $oProd = "Microsoft Office Visio Standard 2002"
                }
            }
        }
        If ($oVer -eq '2003')
        {
            $hProd = $oGUID.Substring($oGUID.Length-38,38).Substring(4,2)
            Switch ($hProd)
            {
                '11'
                {
                    $oProd = "Microsoft Office Professional Enterprise Edition 2003"
                }
                '12'
                {
                    $oProd = "Microsoft Office Standard Edition 2003"
                }
                '13'
                {
                    $oProd = "Microsoft Office Basic Edition 2003"
                }
                '14'
                {
                    $oProd = "Microsoft Windows SharePoint Services 2.0"
                }
                '15'
                {
                    $oProd = "Microsoft Office Access 2003"
                }
                '16'
                {
                    $oProd = "Microsoft Office Excel 2003"
                }
                '17'
                {
                    $oProd = "Microsoft Office FrontPage 2003"
                }
                '18'
                {
                    $oProd = "Microsoft Office PowerPoint 2003"
                }
                '19'
                {
                    $oProd = "Microsoft Office Publisher 2003"
                }
                '1A'
                {
                    $oProd = "Microsoft Office Outlook Professional 2003"
                }
                '1B'
                {
                    $oProd = "Microsoft Office Word 2003"
                }
                '1C'
                {
                    $oProd = "Microsoft Office Access 2003 Runtime"
                }
                '1E'
                {
                    $oProd = "Microsoft Office 2003 User Interface Pack"
                }
                '1F'
                {
                    $oProd = "Microsoft Office 2003 Proofing Tools"
                }
                '23'
                {
                    $oProd = "Microsoft Office 2003 Multilingual User Interface Pack"
                }
                '24'
                {
                    $oProd = "Microsoft Office 2003 Resource Kit"
                }
                '26'
                {
                    $oProd = "Microsoft Office XP Web Components"
                }
                '2E'
                {
                    $oProd = "Microsoft Office 2003 Research Service SDK"
                }
                '44'
                {
                    $oProd = "Microsoft Office InfoPath 2003"
                }
                '83'
                {
                    $oProd = "Microsoft Office 2003 HTML Viewer"
                }
                '92'
                {
                    $oProd = "Windows SharePoint Services 2.0 English Template Pack"
                }
                '93'
                {
                    $oProd = "Microsoft Office 2003 English Web Parts and Components"
                }
                'A1'
                {
                    $oProd = "Microsoft Office OneNote 2003"
                }
                'A4'
                {
                    $oProd = "Microsoft Office 2003 Web Components"
                }
                'A5'
                {
                    $oProd = "Microsoft SharePoint Migration Tool 2003"
                }
                'AA'
                {
                    $oProd = "Microsoft Office PowerPoint 2003 Presentation Broadcast"
                }
                'AB'
                {
                    $oProd = "Microsoft Office PowerPoint 2003 Template Pack 1"
                }
                'AC'
                {
                    $oProd = "Microsoft Office PowerPoint 2003 Template Pack 2"
                }
                'AD'
                {
                    $oProd = "Microsoft Office PowerPoint 2003 Template Pack 3"
                }
                'AE'
                {
                    $oProd = "Microsoft Organization Chart 2.0"
                }
                'CA'
                {
                    $oProd = "Microsoft Office Small Business Edition 2003"
                }
                'D0'
                {
                    $oProd = "Microsoft Office Access 2003 Developer Extensions"
                }
                'DC'
                {
                    $oProd = "Microsoft Office 2003 Smart Document SDK"
                }
                'E0'
                {
                    $oProd = "Microsoft Office Outlook Standard 2003"
                }
                'E3'
                {
                    $oProd = "Microsoft Office Professional Edition 2003 (with InfoPath 2003)"
                }
                'FD'
                {
                    $oProd = "Microsoft Office Outlook 2003 (distributed by MSN)"
                }
                'FF'
                {
                    $oProd = "Microsoft Office 2003 Edition Language Interface Pack"
                }
                'F8'
                {
                    $oProd = "Remove Hidden Data Tool"
                }
                '3A'
                {
                    $oProd = "Microsoft Office Project Standard 2003"
                }
                '3B'
                {
                    $oProd = "Microsoft Office Project Professional 2003"
                }
                '32'
                {
                    $oProd = "Microsoft Office Project Server 2003"
                }
                '51'
                {
                    $oProd = "Microsoft Office Visio Professional 2003"
                }
                '52'
                {
                    $oProd = "Microsoft Office Visio Viewer 2003"
                }
                '53'
                {
                    $oProd = "Microsoft Office Visio Standard 2003"
                }
                '55'
                {
                    $oProd = "Microsoft Office Visio for Enterprise Architects 2003"
                }
                '5E'
                {
                    $oProd = "Microsoft Office Visio 2003 Multilingual User Interface Pack"
                }
            }
        }
        If ($oVer -eq '2007')
        {
            $hProd = $oGUID.Substring($oGUID.Length-38,38).Substring(4,2)
            Switch ($hProd)
            {
                '0011'
                {
                    $oProd = "Microsoft Office Professional Plus 2007"
                }
                '0012'
                {
                    $oProd = "Microsoft Office Standard 2007"
                }
                '0013'
                {
                    $oProd = "Microsoft Office Basic 2007"
                }
                '0014'
                {
                    $oProd = "Microsoft Office Professional 2007"
                }
                '0015'
                {
                    $oProd = "Microsoft Office Access 2007"
                }
                '0016'
                {
                    $oProd = "Microsoft Office Excel 2007"
                }
                '0017'
                {
                    $oProd = "Microsoft Office SharePoint Designer 2007"
                }
                '0018'
                {
                    $oProd = "Microsoft Office PowerPoint 2007"
                }
                '0019'
                {
                    $oProd = "Microsoft Office Publisher 2007"
                }
                '001A'
                {
                    $oProd = "Microsoft Office Outlook 2007"
                }
                '001B'
                {
                    $oProd = "Microsoft Office Word 2007"
                }
                '001C'
                {
                    $oProd = "Microsoft Office Access Runtime 2007"
                }
                '0020'
                {
                    $oProd = "Microsoft Office Compatibility Pack for Word, Excel, and PowerPoint 2007 File Formats"
                }
                '0026'
                {
                    $oProd = "Microsoft Expression Web"
                }
                '002E'
                {
                    $oProd = "Microsoft Office Ultimate 2007"
                }
                '002F'
                {
                    $oProd = "Microsoft Office Home and Student 2007"
                }
                '0030'
                {
                    $oProd = "Microsoft Office Enterprise 2007"
                }
                '0031'
                {
                    $oProd = "Microsoft Office Professional Hybrid 2007"
                }
                '0033'
                {
                    $oProd = "Microsoft Office Personal 2007"
                }
                '0035'
                {
                    $oProd = "Microsoft Office Professional Hybrid 2007"
                }
                '003A'
                {
                    $oProd = "Microsoft Office Project Standard 2007"
                }
                '003B'
                {
                    $oProd = "Microsoft Office Project Professional 2007"
                }
                '0044'
                {
                    $oProd = "Microsoft Office InfoPath 2007"
                }
                '0051'
                {
                    $oProd = "Microsoft Office Visio Professional 2007"
                }
                '0052'
                {
                    $oProd = "Microsoft Office Visio Viewer 2007"
                }
                '0053'
                {
                    $oProd = "Microsoft Office Visio Standard 2007"
                }
                '00A1'
                {
                    $oProd = "Microsoft Office OneNote 2007"
                }
                '00A3'
                {
                    $oProd = "Microsoft Office OneNote Home Student 2007"
                }
                '00A7'
                {
                    $oProd = "Calendar Printing Assistant for Microsoft Office Outlook 2007"
                }
                '00A9'
                {
                    $oProd = "Microsoft Office InterConnect 2007"
                }
                '00AF'
                {
                    $oProd = "Microsoft Office PowerPoint Viewer 2007 (English)"
                }
                '00B0'
                {
                    $oProd = "The Microsoft Save as PDF add-in"
                }
                '00B1'
                {
                    $oProd = "The Microsoft Save as XPS add-in"
                }
                '00B2'
                {
                    $oProd = "The Microsoft Save as PDF or XPS add-in"
                }
                '00BA'
                {
                    $oProd = "Microsoft Office Groove 2007"
                }
                '00CA'
                {
                    $oProd = "Microsoft Office Small Business 2007"
                }
                '10D7'
                {
                    $oProd = "Microsoft Office InfoPath Forms Services"
                }
                '110D'
                {
                    $oProd = "Microsoft Office SharePoint Server 2007"
                }
                '1122'
                {
                    $oProd = "Windows SharePoint Services Developer Resources 1.2"
                }
                '0010'
                {
                    $oProd = "SKU - Microsoft Software Update for Web Folders (English) 12"
                }
            }
            $oEdit = $oProd
        }
        If ($oVer -eq '2010' -or $oVer -eq '2013')
        {
            Switch ($oEdit)
            {
                'ProjectStdVL'
                {
                    $oProd = 'Microsoft Office Project Standard ' + $oVer + ' (VL)'
                }
                'ProjectProVL'
                {
                    $oProd = 'Microsoft Office Project Professional ' + $oVer + ' (VL)'
                }
                'ProjectProMSDNR'
                {
                    $oProd = 'Microsoft Project Professional ' + $oVer + ' (MSDN)'
                }
                'HomeBusinessR'
                {
                    $oProd = 'Microsoft Office Home and Business ' + $oVer
                }
                'ProfessionalR'
                {
                    $oProd = 'Microsoft Office Professional ' + $oVer
                }
                'ProPlusR'
                {
                    $oProd = 'Microsoft Office Professional Plus ' + $oVer
                }
                'StandardR'
                {
                    $oProd = 'Microsoft Office Standard ' + $oVer
                }
                'StandardVL'
                {
                    $oProd = 'Microsoft Office Standard ' + $oVer + ' (VL)'
                }
                'HomeStudentR'
                {
                    $oProd = 'Microsoft Office Home and Student ' + $oVer
                }
                'AccessRuntimeR'
                {
                    $oProd = 'Microsoft Office Access Runtime ' + $oVer
                }
                'VisioSIR'
                {
                    $oProd = 'Microsoft Office Visio Professional ' + $oVer
                }
                'SPDR'
                {
                    $oProd = 'Microsoft SharePoint Designer ' + $oVer
                }
                'ProjectProR'
                {
                    $oProd = 'Microsoft Project Professional ' + $oVer
                }
                'ProjectStdR'
                {
                    $oProd = 'Microsoft Project Standard ' + $oVer
                }
                'VisioSIVL'
                {
                    $oProd = 'Microsoft Visio ' + $oVer + ' Standard (VL)'
                }
                'InfoPathR'
                {
                    $oProd = 'Microsoft Office InfoPath ' + $oVer
                }
                Default
                {
                    $oProd= 'Microsoft Office Unknown Edition ' + $oVer + ': ' + $oEdit
                }
            }
            If ($oProd.Contains('Microsoft Office Unknown Edition 2010'))
            {
                $hProd = $oGUID.Substring($oGUID.Length-38,38).Substring(11,4)
                Switch ($hProd)
                {
                    '0011'
                    {
                        $oProd = 'Microsoft Office Professional Plus 2010'
                    }
                    '0012'
                    {
                        $oProd = 'Microsoft Office Standard 2010'
                    }
                    '0013'
                    {
                        $oProd = 'Microsoft Office Home and Business 2010'
                    }
                    '0014'
                    {
                        $oProd = 'Microsoft Office Professional 2010'
                    }
                    '0015'
                    {
                        $oProd = 'Microsoft Access 2010'
                    }
                    '0016'
                    {
                        $oProd = 'Microsoft Excel 2010'
                    }
                    '0017'
                    {
                        $oProd = 'Microsoft SharePoint Designer 2010'
                    }
                    '0018'
                    {
                        $oProd = 'Microsoft PowerPoint 2010'
                    }
                    '0019'
                    {
                        $oProd = 'Microsoft Publisher 2010'
                    }
                    '001A'
                    {
                        $oProd = 'Microsoft Outlook 2010'
                    }
                    '001B'
                    {
                        $oProd = 'Microsoft Word 2010'
                    }
                    '001C'
                    {
                        $oProd = 'Microsoft Access Runtime 2010'
                    }
                    '001F'
                    {
                        $oProd = 'Microsoft Office Proofing Tools Kit Compilation 2010'
                    }
                    '002F'
                    {
                        $oProd = 'Microsoft Office Home and Student 2010'
                    }
                    '003A'
                    {
                        $oProd = 'Microsoft Project Standard 2010'
                    }
                    '003D'
                    {
                        $oProd = 'Microsoft Office Single Image 2010'
                    }
                    '003B'
                    {
                        $oProd = 'Microsoft Project Professional 2010'
                    }
                    '0044'
                    {
                        $oProd = 'Microsoft InfoPath 2010'
                    }
                    '0052'
                    {
                        $oProd = 'Microsoft Visio Viewer 2010'
                    }
                    '0057'
                    {
                        $oProd = 'Microsoft Visio 2010'
                    }
                    '007A'
                    {
                        $oProd = 'Microsoft Outlook Connector'
                    }
                    '008B'
                    {
                        $oProd = 'Microsoft Office Small Business Basics 2010'
                    }
                    '00A1'
                    {
                        $oProd = 'Microsoft OneNote 2010'
                    }
                    '00AF'
                    {
                        $oProd = 'Microsoft PowerPoint Viewer 2010'
                    }
                    '00BA'
                    {
                        $oProd = 'Microsoft Office SharePoint Workspace 2010'
                    }
                    '110D'
                    {
                        $oProd = 'Microsoft Office SharePoint Server 2010'
                    }
                    '110F'
                    {
                        $oProd = 'Microsoft Project Server 2010'
                    }
                }
            }
            If ($oProd.Contains('"Microsoft Office Unknown Edition 2013'))
            {
                $hProd = $oGUID.Substring($oGUID.Length-38,38).Substring(11,4)
                Switch ($hProd)
                {
                    '0011'
                    {
                        $oProd = 'Microsoft Office Professional Plus 2013'   
                    }
                    '0012'
                    {
                        $oProd = 'Microsoft Office Standard 2013'   
                    }
                    '0013'
                    {
                        $oProd = 'Microsoft Office Home and Business 2013'   
                    }
                    '0014'
                    {
                        $oProd = 'Microsoft Office Professional 2013'   
                    }
                    '0015'
                    {
                        $oProd = 'Microsoft Access 2013'   
                    }
                    '0016'
                    {
                        $oProd = 'Microsoft Excel 2013'   
                    }
                    '0017'
                    {
                        $oProd = 'Microsoft SharePoint Designer 2013'   
                    }
                    '0018'
                    {
                        $oProd = 'Microsoft PowerPoint 2013'   
                    }
                    '0019'
                    {
                        $oProd = 'Microsoft Publisher 2013'   
                    }
                    '001A'
                    {
                        $oProd = 'Microsoft Outlook 2013'   
                    }
                    '001B'
                    {
                        $oProd = 'Microsoft Word 2013'   
                    }
                    '001C'
                    {
                        $oProd = 'Microsoft Access Runtime 2013'   
                    }
                    '001F'
                    {
                        $oProd = 'Microsoft Office Proofing Tools Kit Compilation 2013'   
                    }
                    '002F'
                    {
                        $oProd = 'Microsoft Office Home and Student 2013'   
                    }
                    '003A'
                    {
                        $oProd = 'Microsoft Project Standard 2013'   
                    }
                    '003B'
                    {
                        $oProd = 'Microsoft Project Professional 2013'   
                    }
                    '0044'
                    {
                        $oProd = 'Microsoft InfoPath 2013'   
                    }
                    '0051'
                    {
                        $oProd = 'Microsoft Visio Professional 2013'   
                    }
                    '0053'
                    {
                        $oProd = 'Microsoft Visio Standard 2013'   
                    }
                    '00A1'
                    {
                        $oProd = 'Microsoft OneNote 2013'   
                    }
                    '00BA'
                    {
                        $oProd = 'Microsoft Office SharePoint Workspace 2013'   
                    }
                    '110D'
                    {
                        $oProd = 'Microsoft Office SharePoint Server 2013'   
                    }
                    '110F'
                    {
                        $oProd = 'Microsoft Project Server 2013'   
                    }
                    '012B'
                    {
                        $oProd = 'Microsoft Lync 2013'   
                    }
                }
            }
        }
# Office XP/2003/2007
        If ($oInstall -eq 0)
        {
            $oProdTemp = GetStringValue -hDefKey '0x80000002' -sSubKeyName {'Software\' + $wow + 'Microsoft\Windows\CurrentVersion\Uninstall\' + $oGUID} -sValueName 'DisplayName'
            if ($oProdTemp.Length -gt 0)
            {
                If (
                    $oProd.Length -eq 0 -or
                    $oProd.Contains('Microsoft Office Unknown Edition')
                )
                {
                    $oProd = $oProdTemp
                }
                $oInstall = '1'
            }
        }
# Office 2010/2013/2016
        If ($oInstall -eq '0')
        {
            If ($oEdit.Length -gt 0)
            {
                $kEdit = $oEdit.ToUpper()            
            }
            If ($oGUID.Substring(11,4) -eq '003D')
            {
                $kEdit = 'SingleImage'
            }
            $oProdTemp = GetStringValue -hDefKey '0x80000002' -sSubKeyName {'Software\' + $wow + 'Microsoft\Windows\CurrentVersion\Uninstall\Office' + $aOffID[$version][1].Substring(0, 2) + '.' + $kEdit} -sValueName 'DisplayName'
            If ($oProdTemp.Length -gt 0)
            {
                If (
                    $oProd.Length -gt 0 -or
                    $oProd.Contains('Microsoft Office Unknown Edition')
                )
                {
                    $oProd = $oProdTemp
                }
                $oInstall = 1
            }
        }
# Office 2019
        If ($oProd.Length -eq 0)
        {
            $aUninstallKeys = EnumKey -hDefKey '0x80000002' -sSubKeyName 'Software\Microsoft\Windows\CurrentVersion\Uninstall'
            If ($aUninstallKeys.Length -gt 0)
            {
                Foreach ($UninstallKey in $aUninstallKeys)
                {
                    $path = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\' + $UninstallKey
                    $sValue = GetStringValue -hDefKey '0x80000002' -sSubKeyName $path -sValueName 'UninstallString'
                    If ($sValue.Length -gt 0)
                    {
                        If (
                            $sValue.ToLower().contains('microsoft office ' + $aOffID[$version][1].Substring(0,2)) -or
                            $sValue.ToLower().contains('productstoremove=') -gt 0 -and $sValue.contains('.' + $aOffID[$version][1].Substring(0,2) + '_') -gt 0
                        )
                        {
                            $oNote = $sValue.Substring($sValue.IndexOf('productstoremove='))
                            $oNote = $oNote -replace('productstoremove=','')
                            $oNote = $oNote.Substring(0, $oNote.IndexOf('.1'))
                            $path = "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + $UninstallKey
                            $oProd = GetStringValue -hDefKey '0x80000002' -sSubKeyName $path -sValueName 'DisplayName'
                            $oInstall = '1'
                        }
                    }
                }
            }
        }
        If ($oProd.Length -eq 0)
        {
            $Prod = GetStringValue -hDefKey '0x80000002' -sSubKeyName $regKey -sValueName 'ProductName'
            If ($oProd.Length -gt 0)
            {
                $oProd = GetStringValue -hDefKey '0x80000002' -sSubKeyName $regKey -sValueName 'ConvertToEdition'
            }
        }
        If (
            $oProd.Length -eq 0 -and
            $oOfficeOSPPInfos.Length -gt 0
        )
        {
            $oProd = $oOfficeOSPPInfos['LicenseName']
        }
        If ($oProd.Length -eq 0)
        {
            $oProd = 'Unidentifiable Office' + $oVer
        }
        Else
        {
            If ($oProdID.SubString(7, 3) -eq 'OEM')
            {
                $oProd = $oProd + $oOEM
            }
        }
        writeXML -oVer $oVer -oProd $oProd -oProdID $oProdID -oBit $oBit -oGUID $oGUID -oInstall $oInstall -oKey $oKey -oNote $oNote
    }
}

Function writeXML
{
    Param(
        [Parameter(Mandatory)]
        [string]$oVer,

        [string]$oProd,

        [string]$oProdID,

        [string]$oBit,

        [string]$oGUID,

        [string]$oInstall,

        [string]$oKey,

        [string]$oNote
    )
    $xml = ''
    $xml += "<OFFICEPACK>`n"
	$xml += "<OFFICEVERSION>" + $oVer + "</OFFICEVERSION>`n"
	$xml += "<PRODUCT>" + $oProd + "</PRODUCT>`n"
	$xml += "<PRODUCTID>" + $oProdID + "</PRODUCTID>`n"
	$xml += "<TYPE>" + $oBit + "</TYPE>`n"
	$xml += "<OFFICEKEY>" + $oKey + "</OFFICEKEY>`n"
	$xml += "<GUID>" + $oGUID + "</GUID>`n"
	$xml += "<INSTALL>" + $oInstall + "</INSTALL>`n"
	$xml += "<NOTE>" + $oNote + "</NOTE>`n"
	$xml += "</OFFICEPACK>`n"
    $PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
    [Console]::WriteLine($xml)
}

###
# core
###
<#
Supported Office Families:
    Office 2000 -> KB230848 - https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/230848
    Office XP -> KB302663 - https://www.betaarchive.com/wiki/index.php?title=Microsoft_KB_Archive/302663
    Office 2003 -> KB832672 - https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/826217
    Office 2007 -> KB928516 - https://www.betaarchive.com/wiki/index.php/Microsoft_KB_Archive/928516
    Office 2010 -> KB2186281 - https://support.microsoft.com/en-us/topic/description-of-the-numbering-scheme-for-product-code-guids-in-office-2010-cceaef56-3d3f-1cae-8577-b4de3beaacfa
    Office 2013 -> KB2786054 - https://support.microsoft.com/help/2786054 
    Office 2016, O365 -> https://docs.microsoft.com/en-us/office/troubleshoot/office-suite-issues/numbering-scheme-for-product-guid
#>
$OFFICE_ALL = '78E1-11D2-B60F-006097C998E7}.0001-11D2-92F2-00104BC947F0}.6000-11D3-8CFE-0050048383C9}.6000-11D3-8CFE-0150048383C9}.7000-11D3-8CFE-0150048383C9}.BE5F-4ED1-A0F7-759D40C7622E}.BDCA-11D1-B7AE-00C04FB92F3D}.6D54-11D4-BEE3-00C04F990354}.CFDA-404E-8992-6AF153ED1719}.{9AC08E99-230B-47e8-9721-4577B7F124EA}'
$OFFICEID = '000-0000000FF1CE}'
$aOSPPVersions = ''
$aOffID = @(
    @('XP', '10.0'),
    @('2003', '11.0'),
    @('2007', '12.0'),
    @('2010', '14.0'),
    @('2013', '15.0'),
    @('2016', '16.0')
)
$osType = 32
if ({GetStringValue -hDefKey '0x80000002' -sSubKeyName 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment' -sValueName 'PROCESSOR_ARCHITECTURE'} -eq 'AMD64')
{
    $osType = 64
}
$wow = ''
if ($osType = '64')
{
    $wow = 'WOW6432Node\'
}
schKey97 -regKey {'SOFTWARE\' + $wow + 'Microsoft\'}
schKey2K -name "Office" -regKey {"SOFTWARE\" + $wow + "Microsoft\Office\9.0\"} -guid1 "0000","0001","0002","0003","0004","0010","0011","0012","0013","0014","0016","0017","0018","001A","004F" -guid2 "78E1-11D2-B60F-006097C998E7"
schKey2K -name "Visio" -regKey {"SOFTWARE\" + $wow + "Microsoft\Visio\6.0\"} -guid1 "B66F45DC" -guid2 "853B-11D3-83DE-00C04F3223C8"
For ($i = 0; $i -lt $aOffID.Length; $i++)
{
    $path = 'SOFTWARE\Wow6432Node\Microsoft\Office\' + $aOffID[$i][1] + '\Registration'
    schKey -regKey $path -likeOS False -version $i
    $path = 'SOFTWARE\Microsoft\Office\' + $aOffID[$i][1] + '\Registration'
	schKey -regKey $path -likeOS True -version $i
    $path = 'SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\Microsoft\Office\' + $aOffID[$i][1]+ '\Registration'
	schKey -regKey $path -likeOS False -version $i
    $path = 'SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\' + $aOffID[$i][1] + '\Registration'
	schKey -regKey $path -likeOS True -version $i
}
