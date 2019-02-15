#include <Array.au3>
; https://www.autoitscript.com/forum/topic/184049-uninstall-with-autoit/?tab=comments#comment-1321816
Local $sSubkey
Dim $ax86[1][2]
Dim $ax64[1][2]
Local $sKey = "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
Local $sx64Key = "HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

$a = 1
While 1
	$sSubkey = RegEnumKey($sKey, $a)
	If @error Then
		ExitLoop
	Else
		$sKey1 = $sKey & $sSubkey
		$sDisplay = RegRead($sKey1, "DisplayName")
		$fill = $sSubkey & '|' & $sDisplay
		_ArrayAdd($ax86, $fill)
	EndIf
	$a += 1
WEnd

$b = 1
While 1
	$sSubkey = RegEnumKey($sx64Key, $b)
	If @error Then
		ExitLoop
	Else
		$sKey1 = $sx64Key & $sSubkey
		$sDisplay = RegRead($sKey1, "DisplayName")
		$fill = $sSubkey & '|' & $sDisplay
		_ArrayAdd($ax64, $fill)
	EndIf
	$b += 1
WEnd

$ax86[0][0] = "Registry Key"
$ax86[0][1] = "Display Name"
$ax64[0][0] = "Registry Key"
$ax64[0][1] = "Display Name"

_ArrayDisplay($ax86, "Keys Under " & $sKey)
_ArrayDisplay($ax64, "Keys Under " & $sx64Key)
