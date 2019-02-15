#include <Date.au3> ; needed for _UninstallList function



; Examples ##########################################################################################################
#Include <Array.au3> ; Just for _ArrayDisplay
Local $aList

; Lists all uninstall keys
$aList = _UninstallList()
_ArrayDisplay($aList, "All uninstall keys", Default, Default, Default, "RegistryPath|RegistrySubKey|DisplayName|Date")

; Lists all keys, where the publisher name (Publisher value) starts with "adobe"
$aList = _UninstallList("Publisher", "Adobe")
_ArrayDisplay($aList, "Adobe Publisher", Default, Default, Default, "RegistryPath|RegistrySubKey|DisplayName|Date")

; Lists all x86 keys only, where the name (DisplayName value) contains "flash"
$aList = _UninstallList("DisplayName", "Flash", "", 1, 1)
_ArrayDisplay($aList, "Flash (x86)", Default, Default, Default, "RegistryPath|RegistrySubKey|DisplayName|Date")

; Lists all keys matching a Java version (using a regular expression)
$aList = _UninstallList("DisplayName", "(?i)Java \d+ Update \d+", "", 3)
_ArrayDisplay($aList, "Java", Default, Default, Default, "RegistryPath|RegistrySubKey|DisplayName|Date")

; Lists all x64 keys only, where the quiet uninstall string (QuietUninstallString value) is set
$aList = _UninstallList("QuietUninstallString", ".+", "QuietUninstallString", 3, 2)
_ArrayDisplay($aList, "QuietUninstallString (x64)", Default, Default, Default, "RegistryPath|RegistrySubKey|DisplayName|Date|QuietUninstallString")

; List all x86 keys only, where the name (DisplayName value) start with "Autoit"  and retrieve the
; UninstallString, DisplayVersion values
$aList = _UninstallList("DisplayName", "Autoit", "UninstallString|DisplayVersion", 0, 1)
_ArrayDisplay($aList, "Autoit (x86)", Default, Default, Default, "RegistryPath|RegistrySubKey|DisplayName|Date|UninstallString|DisplayVersion")
; ###################################################################################################################




; #FUNCTION# ====================================================================================================================
; Name ..........: _UninstallList
; Description ...: Returns an array of matching uninstall keys from registry, with an optional filter
; Syntax ........: _UninstallList([$sValueName = ""[, $sFilter = ""[, $sCols = ""[, $iSearchMode = 0[,$ iArch = 3]]]]]])
; Parameters ....: $sValueName       - [optional] Registry value used for the filter.
;                                          Default is all keys ($sFilter do not operates).
;                  $sFilter          - [optional] String to search in $sValueName. Filter is not case sensitive.
;                  $sCols            - [optional] Additional values to retrieve. Use "|" to separate each value.
;                                          Each value adds a column in the returned array
;                  $iSearchMode      - [optional] Search mode. Default is 0.
;                                          0 : Match string from the start.
;                                          1 : Match any substring.
;                                          2 : Exact string match.
;                                          3 : $sFilter is a regular expression
;                  $iArch            - [optional] Registry keys to search in. Default is 3.
;                                          1 : x86 registry keys only
;                                          2 : x64 registry keys only
;                                          3 : both x86 and x64 registry keys
; Return values .: Returns a 2D array of registry keys and values :
;                      $array[0][0] : Number of keys
;                      $array[n][0] : Registry key path
;                      $array[n][1] : Registry subkey
;                      $array[n][2] : Display name
;                      $array[n][3] : Installation date (YYYYMMDD format)
;                      $array[n][4] : 1st additional value specified in $sCols (only if $sCols is set)
;                      $array[n][5] : 2nd additional value specified in $sCols (only if $sCols contains at least 2 entries)
;                      $array[n][x] : Nth additional value ...
; Author ........: jguinch
; ===============================================================================================================================
Func _UninstallList($sValueName = "", $sFilter = "", $sCols = "", $iSearchMode = 0, $iArch = 3)
    Local $sHKLMx86, $sHKLM64, $sHKCU = "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    Local $aKeys[1] = [ $sHKCU ]
    Local $sDisplayName, $sSubKey, $sKeyDate, $sDate, $sValue, $iFound, $n, $aResult[1][4], $iCol
    Local $aCols[1] = [0]

    If NOT IsInt($iArch) OR $iArch < 0 OR $iArch > 3 Then Return SetError(1, 0, 0)
    If NOT IsInt($iSearchMode) OR $iSearchMode < 0 OR $iSearchMode > 3 Then Return SetError(1, 0, 0)

    $sCols = StringRegExpReplace( StringRegExpReplace($sCols, "(?i)(DisplayName|InstallDate)\|?", ""), "\|$", "")
    If $sCols <> "" Then $aCols = StringSplit($sCols, "|")

    If @OSArch = "X86" Then
        $iArch = 1
        $sHKLMx86 = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    Else
        If @AutoitX64 Then
            $sHKLMx86 = "HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            $sHKLM64 = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        Else
            $sHKLMx86 = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $sHKLM64 = "HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        EndIf
    EndIf

    If BitAND($iArch, 1) Then
        Redim $aKeys[ UBound($aKeys) + 1]
        $aKeys [ UBound($aKeys) - 1] = $sHKLMx86
    EndIf

    If BitAND($iArch, 2) Then
        Redim $aKeys[ UBound($aKeys) + 1]
        $aKeys [ UBound($aKeys) - 1] = $sHKLM64
    EndIf


    For $i = 0 To UBound($aKeys) - 1
        $n = 1
        While 1
            $iFound = 1
            $aSubKey = _RegEnumKeyEx($aKeys[$i], $n)
            If @error Then ExitLoop

            $sSubKey = $aSubKey[0]
            $sKeyDate = StringRegExpReplace($aSubKey[1], "^(\d{4})/(\d{2})/(\d{2}).+", "$1$2$3")
            $sDisplayName = RegRead($aKeys[$i] & "\" & $sSubKey, "DisplayName")
            $sDate = RegRead($aKeys[$i] & "\" & $sSubKey, "InstallDate")
            If $sDate = "" Then $sDate = $sKeyDate

            If $sDisplayName <> "" Then
                 If $sValueName <> "" Then
                    $iFound = 0
                    $sValue = RegRead( $aKeys[$i] & "\" & $sSubKey, $sValueName)
                    If ( $iSearchMode = 0 AND StringInStr($sValue, $sFilter) = 1 ) OR _
                       ( $iSearchMode = 1 AND StringInStr($sValue, $sFilter) ) OR _
                       ( $iSearchMode = 2 AND $sValue = $sFilter ) OR _
                       ( $iSearchMode = 3 AND StringRegExp($sValue, $sFilter) ) Then
                            $iFound = 1
                    EndIf
                EndIf

                If $iFound Then
                    Redim $aResult[ UBound($aResult) + 1][ 4 + $aCols[0] ]
                    $aResult[ UBound($aResult) - 1][0] = $aKeys[$i]
                    $aResult[ UBound($aResult) - 1][1] = $sSubKey
                    $aResult[ UBound($aResult) - 1][2] = $sDisplayName
                    $aResult[ UBound($aResult) - 1][3] = $sDate

                    For $iCol = 1 To $aCols[0]
                        $aResult[ UBound($aResult) - 1][3 + $iCol] = RegRead( $aKeys[$i] & "\" & $sSubKey, $aCols[$iCol])
                    Next
                EndIf
            EndIf

            $n += 1
        WEnd
    Next

    $aResult[0][0] = UBound($aResult) - 1
    Return $aResult
EndFunc




; #FUNCTION# ====================================================================================================================
; Name ..........: _RegEnumKeyEx
; Description ...: Enumerates the subkeys of the specified open registry key. The function retrieves information about one subkey
;                  each time it is called.
; Syntax ........: _RegEnumKeyEx($sKey, $iInstance)
; Parameters ....: $sKey                - The registry key to read.
;                  $iInstance           - The 1-based key instance to retrieve.
; Return values .: Success              - A 1D array :
;                                          $aArray[0] = subkey name
;                                          $aArray[1] = time at which the enumerated subkey was last written
;                  Failure               - Returns 0 and set @eror to non-zero value
; Author ........: jguinch
; ===============================================================================================================================
Func _RegEnumKeyEx($sKey, $iInstance)
    If NOT IsDeclared("KEY_WOW64_32KEY") Then Local Const $KEY_WOW64_32KEY = 0x0200
    If NOT IsDeclared("KEY_WOW64_64KEY") Then Local Const $KEY_WOW64_64KEY = 0x0100
    If NOT IsDeclared("KEY_ENUMERATE_SUB_KEYS") Then Local Const $KEY_ENUMERATE_SUB_KEYS = 0x0008

    If NOT IsDeclared("tagFILETIME") Then Local Const $tagFILETIME = "struct;dword Lo;dword Hi;endstruct"

    Local $iSamDesired = $KEY_ENUMERATE_SUB_KEYS

    Local $iX64Key = 0, $sRootKey, $aResult[2]

    Local $sRoot = StringRegExpReplace($sKey, "\\.+", "")
    Local $sSubkey = StringRegExpReplace($sKey, "^[^\\]+\\", "")

    $sRoot = StringReplace($sRoot, "64", "")
    If @extended Then $iX64Key = 1

    If NOT IsInt($iInstance) OR $iInstance < 1 Then Return SetError(2, 0, 0)

    Switch $sRoot
        Case "HKCR", "HKEY_CLASSES_ROOT"
            $sRootKey = 0x80000000
        Case "HKLM", "HKEY_LOCAL_MACHINE"
            $sRootKey = 0x80000002
        Case "HKCU", "HKEY_CURRENT_USER"
            $sRootKey = 0x80000001
        Case "HKU", "HKEY_USERS"
            $sRootKey = 0x80000003
        Case  "HKCC", "HKEY_CURRENT_CONFIG"
            $sRootKey = 0x80000005
        Case Else
            Return SetError(1, 0, 0)
    EndSwitch

    If StringRegExp(@OSArch, "64$") Then
        If @AutoItX64 OR $iX64Key Then
            $iSamDesired = BitOR($iSamDesired, $KEY_WOW64_64KEY)
        Else
            $iSamDesired = BitOR($iSamDesired, $KEY_WOW64_32KEY)
        EndIf
    EndIf

    Local $aRetOPen = DllCall('advapi32.dll', 'long', 'RegOpenKeyExW', 'handle', $sRootKey, 'wstr', $sSubKey, 'dword', 0, 'dword', $iSamDesired, 'ulong_ptr*', 0)
    If @error Then Return SetError(@error, @extended, 0)
    If $aRetOPen[0] Then Return SetError(10, $aRetOPen[0], 0)

    Local $hKey = $aRetOPen[5]

    Local $tFILETIME = DllStructCreate($tagFILETIME)
    Local $lpftLastWriteTime = DllStructGetPtr($tFILETIME)

    Local $aRetEnum = DllCall('Advapi32.dll', 'long', 'RegEnumKeyExW', 'long', $hKey, 'dword', $iInstance - 1, 'wstr', "", 'dword*', 255, 'dword', "", 'ptr', "", 'dword', "", 'ptr', $lpftLastWriteTime)
    If Not IsArray($aRetEnum) OR $aRetEnum[0] <> 0 Then Return SetError( 3, 0, 1)

    Local $tFILETIME2 = _Date_Time_FileTimeToLocalFileTime($lpftLastWriteTime)
    Local $localtime = _Date_Time_FileTimeToStr($tFILETIME2, 1)

    $aResult[0] = $aRetEnum[3]
    $aResult[1] = $localtime

    Return $aResult
EndFunc