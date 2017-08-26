#include <Array.au3>
#include <Excel.au3>

Global $DM = ObjCreate("dm.dmsoft")
Global $FindPic
Global $PicPos
Global $aFormExcel
Global $form_bt

FileInstall("1htsjywxfd.bmp", @TempDir & "\1htsjywxfd.bmp")
FileInstall("1htsjywxfd1.bmp", @TempDir & "\1htsjywxfd1.bmp")
FileInstall("2qyhyy.bmp", @TempDir & "\2qyhyy.bmp")
FileInstall("3cjbd.bmp", @TempDir & "\3cjbd.bmp")
FileInstall("4txbd.bmp", @TempDir & "\4txbd.bmp")
FileInstall("5ywgd.bmp", @TempDir & "\5ywgd.bmp")
FileInstall("6qtxbt.bmp", @TempDir & "\6qtxbt.bmp")
FileInstall("7qqbm1.bmp", @TempDir & "\7qqbm1.bmp")
FileInstall("8qqbm2.bmp", @TempDir & "\8qqbm2.bmp")
FileInstall("9sjlx1.bmp", @TempDir & "\9sjlx1.bmp")
FileInstall("9sjlx2.bmp", @TempDir & "\9sjlx2.bmp")
FileInstall("10qqms.bmp", @TempDir & "\10qqms.bmp")
FileInstall("11jzsc.bmp", @TempDir & "\11jzsc.bmp")
FileInstall("12ljtj.bmp", @TempDir & "\12ljtj.bmp")
FileInstall("13js.bmp", @TempDir & "\13js.bmp")
FileInstall("14ywc.bmp", @TempDir & "\14ywc.bmp")
FileInstall("15qd.bmp", @TempDir & "\15qd.bmp")

$aFormExcel = ReadFormExcel()
_ArrayDisplay($aFormExcel)
;~ OpenForm()
$aFormExcel_row = UBound($aFormExcel)
ConsoleWrite($aFormExcel_row & @CRLF)
For $i = 1 To $aFormExcel_row - 1 Step 1
;~ For $i = 1 To 1 Step 1
	ConsoleWrite($aFormExcel[$i][0] & @CRLF)
	Txbd($aFormExcel[$i][0], $aFormExcel[$i][1], $aFormExcel[$i][2], $aFormExcel[$i][3], $aFormExcel[$i][4], $aFormExcel[$i][5], $aFormExcel[$i][6])
Next

Func ReadFormExcel()
	Local $oExcel = _Excel_Open(0)
	Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & "\FormExcel.xlsx", True)
	Local $aResult = _Excel_RangeRead($oWorkbook, Default)
	_Excel_Close($oExcel)
	Return ($aResult)
EndFunc   ;==>ReadFormExcel


Func OpenForm()
	WinActivate("[CLASS:WeChatMainWndForPC]")
	WinWaitActive("[CLASS:WeChatMainWndForPC]")
;~ 	WinSetOnTop("[CLASS:WeChatMainWndForPC]", "", 1)
	$a = WinGetPos("[CLASS:WeChatMainWndForPC]")
;~ 	ConsoleWrite(_ArrayToString($a) & @CRLF)
	$m = FindPos(@TempDir & "\1htsjywxfd.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		$mp = MouseGetPos()
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		MouseMove($mp[0], $mp[1], 0)
		Sleep(100)
		$m[1] = -1
		$m[2] = -1
	EndIf
	$a = WinGetPos("[CLASS:WeChatMainWndForPC]")
	$m = FindPos(@TempDir & "\1htsjywxfd1.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		$mp = MouseGetPos()
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		MouseMove($mp[0], $mp[1], 0)
		Sleep(100)
		$m[1] = -1
		$m[2] = -1
	EndIf
	$a = WinGetPos("[CLASS:WeChatMainWndForPC]")
	$m = FindPos(@TempDir & "\2qyhyy.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		$mp = MouseGetPos()
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		MouseMove($mp[0], $mp[1], 0)
		Sleep(100)
		$m[1] = -1
		$m[2] = -1
	EndIf
	$a = WinGetPos("[CLASS:WeChatMainWndForPC]")
	$m = FindPos(@TempDir & "\3cjbd.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		$mp = MouseGetPos()
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		MouseMove($mp[0], $mp[1], 0)
		Sleep(100)
		$m[1] = -1
		$m[2] = -1
	EndIf
;~ 	WinSetOnTop("[CLASS:WeChatMainWndForPC]", "", 0)
EndFunc   ;==>OpenForm

Func Txbd($form_bt, $form_qqbm1, $form_qqbm2, $form_sjlx1, $form_sjlx2, $form_qqms, $form_ywc)
	WinActivate("[CLASS:WeChatMainWndForPC]")
	WinWaitActive("[CLASS:WeChatMainWndForPC]")
	$a = WinGetPos("[CLASS:WeChatMainWndForPC]")
	$m = FindPos(@TempDir & "\4txbd.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		$mp = MouseGetPos()
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		MouseMove($mp[0], $mp[1], 0)
		Sleep(100)
		$m[1] = -1
		$m[2] = -1
	Else
		Exit
	EndIf
	WinWait("[CLASS:CefWebViewWnd]")
	WinActivate("[CLASS:CefWebViewWnd]")
	WinWaitActive("[CLASS:CefWebViewWnd]")
	Sleep(2000)
	$a = WinGetPos("[CLASS:CefWebViewWnd]")
	$m = FindPos(@TempDir & "\5ywgd.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		$mp = MouseGetPos()
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		MouseMove($mp[0], $mp[1], 0)
		Sleep(100)
		$m[1] = -1
		$m[2] = -1
	Else
		Exit
	EndIf
	Sleep(2000)
	$a = WinGetPos("[CLASS:CefWebViewWnd]")
	$mp = MouseGetPos()
	$m = FindPos(@TempDir & "\6qtxbt.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		Sleep(100)
		Send($form_bt)
		Switch $form_qqbm1
			Case "航天神洁"
				$m_y_delat = 115
			Case Else
				$m_y_delat = 100
		EndSwitch
	Else
		Exit
	EndIf
	$m = FindPos(@TempDir & "\7qqbm1.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		Sleep(100)
		MouseClick("", $m[1] + 10, $m[2] + $m_y_delat, 1, 0)
		Switch $form_qqbm2
			Case "IT服务中心"
				$m_y_delat = 180
			Case Else
				$m_y_delat = 100
		EndSwitch
	Else
		Exit
	EndIf
	$m = FindPos(@TempDir & "\8qqbm2.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		Sleep(100)
		MouseClick("", $m[1] + 10, $m[2] + $m_y_delat, 1, 0)
		Switch $form_sjlx1
			Case "其他"
				$m_y_delat = 130
			Case Else
				$m_y_delat = 100
		EndSwitch
	Else
		Exit
	EndIf
	$m = FindPos(@TempDir & "\9sjlx1.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		Sleep(100)
		MouseClick("", $m[1] + 10, $m[2] + $m_y_delat, 1, 0)
		Switch $form_sjlx1
			Case "其他"
				$m_y_delat = 155
			Case Else
				$m_y_delat = 100
		EndSwitch
	Else
		Exit
	EndIf
	$m = FindPos(@TempDir & "\9sjlx2.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		Sleep(100)
		MouseClick("", $m[1] + 10, $m[2] + $m_y_delat, 1, 0)
	Else
		Exit
	EndIf
	$m = FindPos(@TempDir & "\10qqms.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		MouseClick("", $m[1] + 10, $m[2] + 40, 1, 0)
		Sleep(100)
		Send($form_qqms)
	Else
		Exit
	EndIf
	$m = FindPos(@TempDir & "\11jzsc.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		Sleep(1500)
		Send("{TAB}")
		Send("{TAB}")
	Else
		Exit
	EndIf
	$m = FindPos(@TempDir & "\12ljtj.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		Sleep(100)
	Else
		Exit
	EndIf
	MouseMove($mp[0], $mp[1], 0)
	Sleep(2000)
	$a = WinGetPos("[CLASS:CefWebViewWnd]")
	$mp = MouseGetPos()
	Sleep(100)
	MouseClick("", $a[0] + 200, $a[1] + 170, 1, 0)
	MouseMove($mp[0], $mp[1], 0)
	Sleep(2000)
	$a = WinGetPos("[CLASS:CefWebViewWnd]")
	$mp = MouseGetPos()
	$m = FindPos(@TempDir & "\13js.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
	If $m[2] > 0 Then
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		Sleep(100)
		$m = FindPos(@TempDir & "\14ywc.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
		MouseClick("", $m[1] + 100, $m[2] + 30, 1, 0)
		Sleep(100)
		Send($form_ywc)
		$m = FindPos(@TempDir & "\15qd.bmp", $a[0], $a[1], $a[0] + $a[2], $a[1] + $a[3])
		MouseClick("", $m[1] + 10, $m[2] + 10, 1, 0)
		Sleep(100)
	Else
		Exit
	EndIf
	Sleep(1000)
;~ 	WinClose("[CLASS:CefWebViewWnd]")
EndFunc   ;==>Txbd

Func FindPos($PicToFind, $W_x, $W_y, $W_Width, $W_Height)
	$PicPos = $DM.FindPicE($W_x, $W_y, $W_Width, $W_Height, $PicToFind, "202020", 1, 0)
	ConsoleWrite($PicPos & @TAB & $PicToFind & @CRLF)
	$PicPos = StringSplit($PicPos, "|", 2)
	Return $PicPos
EndFunc   ;==>FindPos


