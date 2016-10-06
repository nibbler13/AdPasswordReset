#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <AD.au3>
#include <Array.au3>
#include <GuiListView.au3>
#include <ColorConstants.au3>
#include <GuiListView.au3>
#include <WinAPI.au3>


local $iWidthGui = 1144
Local $iHeightGui = 390

Local $iWidthLabel = 85
Local $iHeightLabel = 17

Local $iWidthInput = 270
Local $iHeightInput = 21

Local $iWidthButton = 100
Local $iHeightButton = 30

Local $iWidthTotal = $iWidthInput + $iWidthLabel
Local $iPositionLeft = 20
Local $iPositionTop = 20
Local $iGap = 8
Local $iIconSize = 32

Local $idInputNameToSearch = -666
Local $idButtonSearch = -667
Local $idButtonNext = -668
Local $idListViewResults = -669

Local $aAccountProperties = 0

If @LogonDomain <> "" Then
    ; Open Connection to the Active Directory
    _AD_Open()
    If @error Then Exit MsgBox(16, _
		"Active Directory Example Skript", _
		"Function _AD_Open encountered a problem. @error = " & _
		@error & ", @extended = " & @extended)
;~     $SUserID1 = @UserName
;~     $SUserID2 = @LogonDomain & "\" & @UserName
;~     $SUserId3 = @UserName & "@" & @LogonDNSDomain
;~     $SDNSDomain = $sAD_DNSDomain
;~     $SHostServer = $sAD_HostServer
;~     $SConfiguration = $sAD_Configuration
Else
	MsgBox(0, "", "error")
EndIf

$wProcOld = 0
Global Const $VK_RETURN = 0x0D ;Enter key
GUIRegisterMsg($WM_NOTIFY, "MY_WM_NOTIFY")
GUIRegisterMsg($WM_COMMAND, "My_WM_COMMAND")

$wProcHandle = DllCallbackRegister("_WindowProc", "int", "hwnd;uint;wparam;lparam")

FormMainGui()





Func FormMainGui()
	Local $idGuiMain = GUICreate('Сброс паролей пользователей клиники "Будь здоров"', $iWidthGui, $iHeightGui, -1, -1)

	Local $idLabelName = GUICtrlCreateLabel("Имя пользователя:", 0, 0)
	Local $aControlPosition = ControlGetPos($idGuiMain, "", $idLabelName)
	$iPositionLeft = ($iWidthGui - $aControlPosition[2] - $iWidthButton - $iWidthInput - $iGap * 3) / 2
	GUICtrlSetPos($idLabelName, $iPositionLeft, $iPositionTop + 3)
	$aControlPosition = ControlGetPos($idGuiMain, "", $idLabelName)

	$idInputNameToSearch = GUICtrlCreateInput("temp", $aControlPosition[0] + $aControlPosition[2] + $iGap, _
		$iPositionTop, $iWidthInput, $iHeightInput)
	$aControlPosition = ControlGetPos($idGuiMain, "", $idInputNameToSearch)

	$idButtonSearch = GUICtrlCreateButton("Поиск", $aControlPosition[0] + $aControlPosition[2] + $iGap * 2, _
		$iPositionTop - ($iHeightButton - $iHeightInput) / 2, $iWidthButton, $iHeightButton)
	$aControlPosition = ControlGetPos($idGuiMain, "", $idButtonSearch)
;~ 	GUICtrlSetState(-1, $GUI_DISABLE)

	$iPositionLeft = 20
	$idButtonNext = GUICtrlCreateButton("Далее", $iWidthGui - $iWidthButton - $iPositionLeft, _
		$aControlPosition[1], $iWidthButton, $iHeightButton)
	GUICtrlSetState(-1, $GUI_DISABLE)

	Local $labelResults = GUICtrlCreateLabel("Результаты поиска:", $iPositionLeft, _
		$aControlPosition[1] + $aControlPosition[3] + $iGap)
	$aControlPosition = ControlGetPos($idGuiMain, "", $labelResults)

	Local $widthList = $iWidthGui - $iPositionLeft * 2
	Local $heightList = $iHeightGui - $aControlPosition[1] - $aControlPosition[3] - $iPositionLeft
	$idListViewResults = GUICtrlCreateListView("ФИО|Филиал|Подразделение|Должность|Логин", $iPositionLeft, _
		$aControlPosition[1] + $aControlPosition[3], $widthList, $heightList, -1, _
		BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_INFOTIP))
	_GUICtrlListView_SetColumnWidth($idListViewResults, 0, $widthList * 0.20)
	_GUICtrlListView_SetColumnWidth($idListViewResults, 1, $widthList * 0.23)
	_GUICtrlListView_SetColumnWidth($idListViewResults, 2, $widthList * 0.29)
	_GUICtrlListView_SetColumnWidth($idListViewResults, 3, $widthList * 0.26)
	_GUICtrlListView_SetColumnWidth($idListViewResults, 4, 0)


	$wProcOld = _WinAPI_SetWindowLong($idListViewResults, $GWL_WNDPROC, DllCallbackGetPtr($wProcHandle))
	_WinAPI_SetWindowLong($idListViewResults, $GWL_WNDPROC, $wProcOld)

	GUISetState(@SW_SHOW)

	Local $boolNeedToSearch = False
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				_AD_Close()
				DllCallbackFree($wProcHandle)
				Exit

			Case $idInputNameToSearch
				$boolNeedToSearch = True

			Case $idButtonSearch
				$boolNeedToSearch = True

			Case $idButtonNext
				ButtonNextPressed()
		EndSwitch

		If $boolNeedToSearch Then
			$boolNeedToSearch = False

			Local $stringEnteredText = GUICtrlRead($idInputNameToSearch)
			If Not $stringEnteredText Then ContinueLoop
			Local $aObjects = _AD_GetObjectsInOU("", _
				"(&(objectCategory=person)(objectClass=user)(name=" & $stringEnteredText & "*))", 2, _
				"displayName,company,department,title,sAMAccountName")

			_GUICtrlListView_DeleteAllItems($idListViewResults)

			If Not IsArray($aObjects) Or UBound($aObjects, $UBOUND_ROWS) < 2 Then
				_GUICtrlListView_AddItem($idListViewResults, "Не найдено записей")
			Else
				_ArrayDelete($aObjects, 0)
				_GUICtrlListView_AddArray($idListViewResults, $aObjects)
			EndIf
		EndIf

		Sleep(20)
	WEnd
EndFunc


Func ButtonNextPressed()
	For $i = 0 To _GUICtrlListView_GetItemCount($idListViewResults) - 1
		If Not _GUICtrlListView_GetItemSelected($idListViewResults, $i) Then ContinueLoop

		Local $dataFromListView = _GUICtrlListView_GetItemTextArray($idListViewResults, $i)
		FormDetailedView($dataFromListView[5])
		Return
	Next
EndFunc


Func FormDetailedView($sAMAccountName)
	$aAccountProperties = _AD_GetObjectProperties($sAMAccountName)
	If Not IsArray($aAccountProperties) Then
		MsgBox(0, "error", "no properties")
		Return
	EndIf

	Local $detailedViewForm = GUICreate("Form1", $iWidthGui, $iHeightGui, -1, -1)
	GUISetBkColor(0xFFFFFF)

	$iWidthLabel = 85
	$iHeightLabel = 17
	$iWidthInput = 270
	$iHeightInput = 21

	$iPositionLeft = 10
	$iPositionTop = 12

	Local $tmp = 0
	Local $accountInfo[0][3]
	_ArrayAdd($accountInfo, CreateAttributeArray("Выводимое имя", "displayName"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Описание:", "description"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Должность:", "title"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Отдел:", "department"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Организация:", "company"))
	$tmp = _AD_GetObjectProperties(GetAttributeText("manager"), "displayName")
	$tmp = IsArray($tmp) ? $tmp[1][1] : ""
	_ArrayAdd($accountInfo, CreateAttributeArray("Руководитель:", $tmp, True))
	_ArrayAdd($accountInfo, CreateAttributeArray("Последний вход:", "lastLogon"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Последний выход:", "lastLogoff"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Последняя неудачная попытка входа:", "badPasswordTime"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Неудачных попыток входа:", "badPwdCount"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Время блокировки:", "lockoutTime"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Имя входа пользователя:", $sAMAccountName, True))
	_ArrayAdd($accountInfo, CreateAttributeArray("Домашний:", "homePhone"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Мобильный:", "mobile"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Рабочий:", "telephoneNumber"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Руководитель:", "", True))
	_ArrayAdd($accountInfo, CreateAttributeArray("Другой:", "", True))
	_ArrayAdd($accountInfo, CreateAttributeArray("Эл. почта:", "mail"))
	_ArrayAdd($accountInfo, CreateAttributeArray("Другой:", "", True))
	_ArrayAdd($accountInfo, CreateAttributeArray("Новый пароль:", "", True))

;~ 	_ArrayDisplay($accountInfo)


	;-------------------- ICON AND DISPLAY NAME ---------------------
	GUICtrlCreateIcon("shell32.dll", 269, $iPositionLeft, $iPositionTop, $iIconSize, $iIconSize)
	$tmp = $iWidthInput + $iWidthLabel - $iIconSize * 2
	GUICtrlCreateLabel($accountInfo[0][1], $iPositionLeft + $iIconSize * 2, $iPositionTop, $tmp, $iIconSize, $SS_CENTERIMAGE)
	$iPositionTop += $iHeightInput + 25
	CreateLine($iWidthInput + $iWidthLabel, 1)


	;-------------------- MAIN ---------------------
	For $i = 1 To 5
		Local $style = ($i < 5 ? "" : $ES_READONLY)
		$accountInfo[$i][2] = CreateLabelAndInput($accountInfo[$i][0], $accountInfo[$i][1], $style)
	Next

	$iPositionTop += $iGap * 2 - $iGap
	CreateLine($iWidthInput + $iWidthLabel, 1)


	;-------------------- EXPIRES ---------------------
	$ctrl = GUICtrlCreateGroup("Срок действия учетной записи", $iPositionLeft, $iPositionTop, _
		$iWidthInput + $iWidthLabel, 84)
	$iPositionTop += 22
	$iPositionLeft += $iGap * 2 - 5

	Local $accountExpiresDate = GetAttributeText("accountExpires")
	Local $accountExpires[2]

	$accountExpires[0] = GUICtrlCreateRadio("Никогда", $iPositionLeft, $iPositionTop, -1, $iHeightLabel)
	GUICtrlSetState(-1, $GUI_CHECKED)
	$iPositionTop += $iGap + $iHeightInput

	$accountExpires[1] = GUICtrlCreateRadio("Истекает:", $iPositionLeft, $iPositionTop, 73, $iHeightLabel)

	Local $dateX = $iPositionLeft + 80
	$iWidthInput -= 1
	Local $dateExpires = GUICtrlCreateDate("", $dateX, $iPositionTop - 4, $iWidthInput + $iWidthLabel - $dateX, 25)
	GUICtrlSetState(-1, $GUI_DISABLE)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	If $accountExpiresDate Then
		GUICtrlSetState($accountExpires[0], $GUI_UNCHECKED)
		GUICtrlSetState($accountExpires[1], $GUI_CHECKED)
		GUICtrlSetData($dateExpires, $accountExpiresDate)
		GUICtrlSetState($dateExpires, $GUI_ENABLE)
	EndIf

	$previosPos = ControlGetPos($detailedViewForm, "", $ctrl)
	$iPositionLeft = $previosPos[0]
	$iPositionTop = $previosPos[1] + $previosPos[3] + $iGap


;~ 	need to check this state and disable password reset if true
;~ 	Local $ctrl = GUICtrlCreateCheckbox("Срок действия учетной записи истек", $iPositionLeft, $iPositionTop)
;~ 	GUICtrlSetState(-1, $GUI_DISABLE)
;~ 	GUICtrlSetState(-1, _AD_IsAccountExpired($sAMAccountName) ? $GUI_CHECKED : $GUI_UNCHECKED)


	;-------------------- VERICAL LINE ---------------------
	Local $previosPos = ControlGetPos($detailedViewForm, "", $ctrl)
	$iPositionLeft = 10 + $iWidthLabel + $iWidthInput + $iGap * 2
	$iPositionTop = 12
	CreateLine(1, $previosPos[1] + $previosPos[3] - $iPositionTop, False)


	;-------------------- BAD PASSWORD ---------------------
	$iWidthInput = 155
	$iWidthLabel = 200

	For $i = 6 To 10
		CreateLabelAndInput($accountInfo[$i][0], $accountInfo[$i][1], $ES_READONLY)
	Next

	Local $checkboxes[0][3]
	_ArrayAdd($checkboxes, CreateAttributeArray("Требовать смены пароля при следующем входе в систему", _
		0, True));_AD_IsPasswordExpired($sAMAccountName), True))
	_ArrayAdd($checkboxes, CreateAttributeArray("Запретить смену пароля пользователем", _
		0, True))
	_ArrayAdd($checkboxes, CreateAttributeArray("Срок действия пароля не ограничен", _
		StringInStr(GetAttributeText("userAccountControl"), "DONT_EXPIRE_PASSWD"), True))
	_ArrayAdd($checkboxes, CreateAttributeArray("Отключить учетную запись", _
		_AD_IsObjectDisabled($sAMAccountName), True))
	_ArrayAdd($checkboxes, CreateAttributeArray("Разблокировать учтеную запись", _
		_AD_IsObjectLocked($sAMAccountName), True))

;~ 	_ArrayDisplay($checkboxes)

	$checkboxes[4][2] = GUICtrlCreateCheckbox($checkboxes[4][0], $iPositionLeft, $iPositionTop, $iWidthInput + $iWidthLabel, $iHeightLabel)
	GUICtrlSetState($checkboxes[4][2], $checkboxes[4][1] ? $GUI_ENABLE : $GUI_DISABLE)
	$iPositionTop += $iGap * 2 + $iHeightInput - 4

;~ 	CreateLine($iWidthInput + $iWidthLabel, 1)

	;-------------------- ACCOUN PARAMETERS ---------------------
	GUICtrlCreateLabel("Параметры учетной записи:", $iPositionLeft, $iPositionTop + 3, $iWidthLabel + $iWidthInput, $iHeightLabel)
	$iPositionTop += $iGap + $iHeightLabel

	$ctrl = GUICtrlCreateLabel("", $iPositionLeft, $iPositionTop, $iWidthInput + $iWidthLabel, 110, Default, $WS_EX_CLIENTEDGE)
	GUICtrlSetBkColor(-1, 0xf0f0f0)
	$iPositionTop += $iGap
	GUICtrlSetState(-1, $GUI_DISABLE)
	$iPositionLeft += $iGap * 1.5

	For $i = 0 To UBound($checkboxes, $UBOUND_ROWS) - 2
		$checkboxes[$i][2] = GUICtrlCreateCheckbox($checkboxes[$i][0], $iPositionLeft, $iPositionTop, -1, $iHeightLabel)
		If $i < 2 Then GUICtrlSetState(-1, $GUI_DISABLE)
		GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
		If $checkboxes[$i][1] Then GUICtrlSetState(-1, $GUI_CHECKED)
		$iPositionTop += $iGap + $iHeightLabel
	Next


	;-------------------- VERICAL LINE ---------------------
	$previosPos = ControlGetPos($detailedViewForm, "", $ctrl)
	$iPositionLeft = $previosPos[0] + $previosPos[2] + $iGap * 2
	$iPositionTop = 12
	CreateLine(1, $previosPos[1] + $previosPos[3] - $iPositionTop, False)


	;-------------------- LOGIN ---------------------
	$iWidthLabel = 140
	$iWidthInput = 215
	CreateLabelAndInput($accountInfo[11][0], $accountInfo[11][1], $ES_READONLY)
	$iPositionTop += $iGap * 2 - $iGap
	CreateLine($iWidthInput + $iWidthLabel, 1)


	;-------------------- PHONE NUMBERS ---------------------
	$iWidthLabel = 100
	$iWidthInput = 255
	$ctrl = GUICtrlCreateGroup("Номер телефона для отправки СМС уведомления:", $iPositionLeft, $iPositionTop, _
		$iWidthTotal, 6 * $iHeightInput + 5 * $iGap + 4)

	$iPositionTop += 20
	$iPositionLeft += $iGap * 2 - 5
	$iWidthInput -= $iGap * 3 - 2
	For $i = 12 To 16
		GUICtrlCreateRadio($accountInfo[$i][0], $iPositionLeft, $iPositionTop, $iWidthLabel)
;~ 		Local $style = $i < 16 ? $ES_READONLY : -1
		$accountInfo[$i][2] = GUICtrlCreateInput($accountInfo[$i][1], $iPositionLeft + $iWidthLabel, _
			$iPositionTop + 1, $iWidthInput, $iHeightInput)
		$iPositionTop += $iGap + $iHeightInput
	Next

	$previosPos = ControlGetPos($detailedViewForm, "", $ctrl)
	$iPositionTop = $previosPos[1] + $previosPos[3]


	$iPositionTop += $iGap
;~ 	CreateLine($iWidthInput + $iWidthLabel, 1)


	;-------------------- EMAIL ---------------------
	$ctrl = GuiCtrlCreateGroup("Адрес электронной почты для уведомления:", $iPositionLeft - $iGap * 2 + 5, $iPositionTop, _
		$iWidthTotal, 3 * $iHeightInput + 2 * $iGap + 4)
	$iPositionTop += 20

	For $i = 17 To 18
		GUICtrlCreateRadio($accountInfo[$i][0], $iPositionLeft, $iPositionTop, $iWidthLabel)
;~ 		Local $style = $i < 18 ? $ES_READONLY : -1
		$accountInfo[$i][2] = GUICtrlCreateInput($accountInfo[$i][1], $iPositionLeft + $iWidthLabel, $iPositionTop + 1, _
			$iWidthInput, $iHeightInput)
		$iPositionTop += $iGap + $iHeightInput
	Next


	;-------------------- HORIZONTAL LINE ---------------------

	$previosPos = ControlGetPos($detailedViewForm, "", $ctrl)
	$iPositionLeft = 10
	$iPositionTop = $previosPos[1] + $previosPos[3] + $iGap * 2
	CreateLine($iWidthGui - $iGap * 2, 1)

	;-------------------- PASSWORD ---------------------

	$iPositionLeft = 318

	$ctrl = GUICtrlCreateLabel("Контроллер домена:", $iPositionLeft, $iPositionTop + 3, -1, $iHeightLabel)

	$previosPos = ControlGetPos($detailedViewForm, "", $ctrl)
	Local $comboDC = GUICtrlCreateCombo("", $previosPos[0] + $previosPos[2] + $iGap, $previosPos[1] - 3)
	Local $arrayDCList = _AD_ListDomainControllers()
	If IsArray($arrayDCList) Then
		_ArrayDelete($arrayDCList, 0)
		_ArraySort($arrayDCList)
		GUICtrlSetData(-1, _ArrayToString($arrayDCList, "|", -1, -1, "|", 0, 0))
	EndIf

	$previosPos = ControlGetPos($detailedViewForm, "", $comboDC)
	$ctrl = GUICtrlCreateLabel($accountInfo[19][0], $previosPos[0] + $previosPos[2] + $iGap * 2, _
		$previosPos[1] + 3, -1, $iHeightLabel)

	$iWidthInput = 150
	$previosPos = ControlGetPos($detailedViewForm, "", $ctrl)
	$accountInfo[19][2] = GUICtrlCreateInput($accountInfo[19][1], $previosPos[0] + $previosPos[2] + $iGap, _
		$previosPos[1] - 3, $iWidthInput, $iHeightInput)

	;-------------------- BUTTONS ---------------------

	$iPositionLeft = 10

	Local $buttonBack = GUICtrlCreateButton("Закрыть", $iPositionLeft, _
		$previosPos[1] - ($iHeightButton - $previosPos[3]) / 2, $iWidthButton, $iHeightButton)

	$previosPos = ControlGetPos($detailedViewForm, "", $buttonBack)
	Local $buttonReset = GUICtrlCreateButton("Сбросить и уведомить", $iWidthGui - $iWidthButton * 1.5 - $iGap, _
		$previosPos[1], $iWidthButton * 1.5, $iHeightButton, $BS_MULTILINE)

	$previosPos = ControlGetPos($detailedViewForm, "", $buttonReset)
	$buttonApply = GUICtrlCreateButton("Применить", $previosPos[0] - $iGap - $iWidthButton, _
		$previosPos[1], $iWidthButton, $iHeightButton)


	GUISetState(@SW_SHOW)

;~ 	_ArrayDisplay(_AD_ListDomainControllers())
;~ 	_ArrayDisplay(_AD_GetPasswordInfo($sAMAccountName))
;~ 	_ArrayDisplay($accountInfo)

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				_AD_Close()
				Exit
			Case $buttonBack
				$aAccountProperties = 0
				GUIDelete($detailedViewForm)
				Return
			Case $buttonApply
				ConsoleWrite("$buttonApply" & @CRLF)
		EndSwitch
	WEnd

EndFunc





Func _WindowProc($hWnd, $Msg, $wParam, $lParam)
	ConsoleWrite("window proc")
    Switch $hWnd
        Case $idListViewResults
            Switch $Msg
                Case $WM_GETDLGCODE
                    Switch $wParam
                        Case $VK_RETURN
                            ConsoleWrite("Enter key is pressed" & @LF)
                            Return 0
                    EndSwitch
            EndSwitch
    EndSwitch

    Return _WinAPI_CallWindowProc($wProcOld, $hWnd, $Msg, $wParam, $lParam)
EndFunc




Func MY_WM_NOTIFY($hWnd, $Msg, $wParam, $lParam)
	If  $wParam <> $idListViewResults Then Return

	Local $tagNMHDR, $event, $hwndFrom, $code
	$tagNMHDR = DllStructCreate("int;int;int", $lParam)
	If @error Then Return
	$event = DllStructGetData($tagNMHDR, 3)

;~ 	ConsoleWrite($event & @CRLF)

	If $event = $NM_CLICK Or $event = -12 Then
		CheckSelected()
	ElseIf $event = $NM_DBLCLK Then
		ButtonNextPressed()
	EndIf

	$tagNMHDR = 0
	$event = 0
	$lParam = 0
EndFunc


Func My_WM_COMMAND($hWnd, $imsg, $iwParam, $ilParam)
    $nNotifyCode = BitShift($iwParam, 16)
    $hCtrl = $ilParam
    If $nNotifyCode <> $EN_CHANGE Then Return
	If $hCtrl <> GUICtrlGetHandle($idInputNameToSearch) Then Return

	GUICtrlSetState($idButtonSearch, GUICtrlRead($idInputNameToSearch) ? $GUI_ENABLE : $GUI_DISABLE)
EndFunc


Func CheckSelected()
	Local $boolIsSelected = False
	For $i = 0 To _GUICtrlListView_GetItemCount($idListViewResults)
		If _GUICtrlListView_GetItemSelected($idListViewResults, $i) Then
			$boolIsSelected = True
			GUICtrlSetState($idButtonNext, $GUI_ENABLE)
			ExitLoop
		EndIf
	Next

	If Not $boolIsSelected Then GUICtrlSetState($idButtonNext, $GUI_DISABLE)
EndFunc


Func CreateAttributeArray($name, $attribute, $plainText = False)
	Local $tmp[1][2]
	$tmp[0][0] = $name
	$tmp[0][1] = $plainText ? $attribute : GetAttributeText($attribute)
	Return $tmp
EndFunc


Func CreateLine($width, $height, $horizontal = True)
	GUICtrlCreateLabel("", $iPositionLeft, $iPositionTop, $width, $height)
	GUICtrlSetBkColor(-1, 0xa0a0a0)
	If $horizontal Then
		$iPositionTop += $iGap * 2
	Else
		$iPositionLeft += $iGap * 2
	EndIf
EndFunc


Func CreateLabelAndInput($textLabel, $textInput, $style = "")
	GUICtrlCreateLabel($textLabel, $iPositionLeft, $iPositionTop + 3, $iWidthLabel, $iHeightLabel)
	Local $ret = GUICtrlCreateInput($textInput, $iPositionLeft + $iWidthLabel, $iPositionTop, $iWidthInput, $iHeightInput, $style)
	$iPositionTop += $iGap + $iHeightInput
	Return $ret
EndFunc


Func GetAttributeText($attributeName)
;~ 	_ArrayDisplay($aAccountProperties)
	Local $searchResults = _ArraySearch($aAccountProperties, $attributeName)
	Local $ret = $searchResults <> -1 ? $aAccountProperties[$searchResults][1] : ""
	If Not $ret Or $ret = "1601/01/01 00:00:00" Or _
		$ret = "0000/00/00 00:00:00" Then $ret = ""
	ConsoleWrite($attributeName & " : " & $ret & @CRLF)
	Return $ret
EndFunc