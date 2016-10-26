#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
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

Opt("GUICloseOnESC", 0)
Opt("MustDeclareVars", 1)

#Region ====== Variables =========
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
Local $idLabelName = -670

Local $aAccountProperties = 0
Local $idGuiMain = 0

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

Local $wProcOld = 0
Global Const $VK_RETURN = 0x0D ;Enter key

GUIRegisterMsg($WM_NOTIFY, "MY_WM_NOTIFY")
GUIRegisterMsg($WM_COMMAND, "My_WM_COMMAND")

Local $wProcHandle = DllCallbackRegister("_WindowProc", "int", "hwnd;uint;wparam;lparam")

Local $bNextPressed = False
#EndRegion

FormMainGui()





Func FormMainGui($isChiefSelect = False, $idParentGui = "")
	Local $sTitle = 'Сброс паролей пользователей клиники "Будь здоров"'

	If $isChiefSelect Then
		$sTitle = 'Выбор руководителя'
		GUISetState(@SW_DISABLE, $idParentGui)
	EndIf

	$iPositionLeft = 20
	$iPositionTop = 20

	$idGuiMain = GUICreate($sTitle, $iWidthGui, $iHeightGui, -1, -1, -1, -1, $idParentGui)

	$idLabelName = GUICtrlCreateLabel("Имя пользователя:", $iPositionLeft, $iPositionTop + 3)
	Local $aControlPosition = ControlGetPos($idGuiMain, "", $idLabelName)
	$iPositionLeft = ($iWidthGui - $aControlPosition[2] - $iWidthButton - $iWidthInput - $iGap * 3) / 2
	GUICtrlSetPos($idLabelName, $iPositionLeft, $iPositionTop + 3)
	$aControlPosition = ControlGetPos($idGuiMain, "", $idLabelName)

	$idInputNameToSearch = GUICtrlCreateInput("", $aControlPosition[0] + $aControlPosition[2] + $iGap, _
		$iPositionTop, $iWidthInput, $iHeightInput)
	$aControlPosition = ControlGetPos($idGuiMain, "", $idInputNameToSearch)
	GUICtrlSetTip(-1, "Маска поиска: текст*")

	$idButtonSearch = GUICtrlCreateButton("Поиск", $aControlPosition[0] + $aControlPosition[2] + $iGap * 2, _
		$iPositionTop - ($iHeightButton - $iHeightInput) / 2, $iWidthButton, $iHeightButton)
	$aControlPosition = ControlGetPos($idGuiMain, "", $idButtonSearch)

	$iPositionLeft = 20

	Local $widthList = $iWidthGui - $iPositionLeft * 2
	Local $heightList = $iHeightGui - $aControlPosition[1] * 2 - $aControlPosition[3] - $iPositionLeft
	$idListViewResults = GUICtrlCreateListView("ФИО|Филиал|Подразделение|Должность|Логин", $iPositionLeft, _
		$aControlPosition[1] * 2 + $aControlPosition[3], $widthList, $heightList, -1, _
		BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_INFOTIP))
	_GUICtrlListView_SetColumnWidth($idListViewResults, 0, $widthList * 0.20)
	_GUICtrlListView_SetColumnWidth($idListViewResults, 1, $widthList * 0.23)
	_GUICtrlListView_SetColumnWidth($idListViewResults, 2, $widthList * 0.29)
	_GUICtrlListView_SetColumnWidth($idListViewResults, 3, $widthList * 0.26)
	_GUICtrlListView_SetColumnWidth($idListViewResults, 4, 0)

	$wProcOld = _WinAPI_SetWindowLong($idListViewResults, $GWL_WNDPROC, DllCallbackGetPtr($wProcHandle))
	_WinAPI_SetWindowLong($idListViewResults, $GWL_WNDPROC, $wProcOld)

	$aControlPosition = ControlGetPos($idGuiMain, "", $idButtonSearch)
	$idButtonNext = GUICtrlCreateButton("Далее", $iWidthGui - $iWidthButton - $iPositionLeft, _
		$aControlPosition[1], $iWidthButton, $iHeightButton)
	GUICtrlSetState(-1, $GUI_DISABLE)


	$aControlPosition = ControlGetPos($idGuiMain, "", $idButtonSearch)
	Local $idButtonCancel = GUICtrlCreateButton("Отмена", $iPositionLeft, $aControlPosition[1], $iWidthButton, $iHeightButton)
	If Not $isChiefSelect Then GUICtrlSetState(-1, $GUI_HIDE)

	GUISetState(@SW_SHOW)

	Local $boolNeedToSearch = False
	While 1
		Local $nMsg = GUIGetMsg()
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
				$bNextPressed = True

			Case $idButtonCancel
;~ 				DllCallbackFree($wProcHandle)
				GUIDelete($idGuiMain)
				GUISetState(@SW_ENABLE, $idParentGui)
				Return
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

		If $bNextPressed Then
			$bNextPressed = False

			For $i = 0 To _GUICtrlListView_GetItemCount($idListViewResults) - 1
				If Not _GUICtrlListView_GetItemSelected($idListViewResults, $i) Then ContinueLoop

				Local $dataFromListView = _GUICtrlListView_GetItemTextArray($idListViewResults, $i)

				If $isChiefSelect Then
;~ 					DllCallbackFree($wProcHandle)
					GUIDelete($idGuiMain)
					GUISetState(@SW_ENABLE, $idParentGui)
					Return $dataFromListView[5]
				EndIf

				Local $aMainGuiControlsId[] = [ _
					$idInputNameToSearch, _
					$idButtonSearch, _
					$idButtonNext, _
					$idListViewResults, _
					$idLabelName]

				For $id in $aMainGuiControlsId
					GUICtrlSetState($id, $GUI_HIDE)
				Next

				FormDetailedView($dataFromListView[5])

				For $id in $aMainGuiControlsId
					GUICtrlSetState($id, $GUI_SHOW)
				Next

				ExitLoop
			Next
		EndIf

		Sleep(20)
	WEnd
EndFunc


Func FormDetailedView($sAMAccountName)
	$aAccountProperties = _AD_GetObjectProperties($sAMAccountName)
	If Not IsArray($aAccountProperties) Then
		MsgBox(0, "error", "no properties")
		Return
	EndIf

	Local $idProgress = GUICtrlCreateProgress($iPositionLeft, $iHeightGui / 2 + $iGap / 2, _
		$iWidthGui - $iPositionLeft * 2, $iHeightInput)

	Local $idLabelProgress = GUICtrlCreateLabel("Сбор информации из Active Directory", $iPositionLeft, _
		$iHeightGui / 2 - $iHeightLabel - $iGap / 2, $iWidthGui - $iPositionLeft * 2, $iHeightLabel, $SS_CENTER)

	Sleep(10)

	$iWidthLabel = 85
	$iHeightLabel = 17
	$iWidthInput = 270
	$iHeightInput = 21

	$iPositionLeft = 10
	$iPositionTop = 12

	Local $bAccountExpired = False

	Local $idDetailedViewForm = GUICreate("Form1", $iWidthGui, $iHeightGui, 0, 0, $WS_CHILD, -1, $idGuiMain)
	GUISetBkColor(0xFFFFFF)

	Local $tmp = 0
	Local $previosPos = 0
	Local $ctrl = 0

	Local $aAccountInfo[0][4]
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Выводимое имя", "displayName"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Описание:", "description"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Должность:", "title"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Отдел:", "department"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Организация:", "company"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Руководитель:", "manager"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Последний вход:", "lastLogon"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Последний выход:", "lastLogoff"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Последняя неудачная попытка входа:", "badPasswordTime"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Неудачных попыток входа:", "badPwdCount"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Время блокировки:", "lockoutTime"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Имя входа пользователя:", $sAMAccountName, True))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Мобильный:", "mobile"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Рабочий:", "telephoneNumber"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Домашний:", "homePhone"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Руководитель:", "", True))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Другой:", "", True))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Эл. почта:", "mail"))
	_ArrayAdd($aAccountInfo, CreateAttributeArray("Другой:", "", True))


	;-------------------- ICON AND DISPLAY NAME ---------------------
	GUICtrlCreateIcon("shell32.dll", 269, $iPositionLeft, $iPositionTop, $iIconSize, $iIconSize)
	$tmp = $iWidthInput + $iWidthLabel - $iIconSize * 2
	GUICtrlCreateLabel($aAccountInfo[0][1], $iPositionLeft + $iIconSize * 2, $iPositionTop, $tmp, $iIconSize, $SS_CENTERIMAGE)
	$iPositionTop += $iHeightInput + 25
	CreateLine($iWidthInput + $iWidthLabel, 1)


	;-------------------- MAIN ---------------------
	For $i = 1 To 4
		$aAccountInfo[$i][2] = CreateLabelAndInput($aAccountInfo[$i][0], $aAccountInfo[$i][1], "", $aAccountInfo[$i][3])
		GUICtrlSetData($idProgress, $i * 20)
	Next

	$aAccountInfo[5][2] = CreateLabelAndInput($aAccountInfo[5][0], $aAccountInfo[5][1], $ES_READONLY, "", True)
	$previosPos = ControlGetPos($idDetailedViewForm, "", $aAccountInfo[5][2])
	Local $idButtonChief = GUICtrlCreateButton("...", $previosPos[0] + $previosPos[2] + $iGap, $previosPos[1], $iGap * 3, $previosPos[3])

	$iPositionTop += $iGap * 2 - $iGap
	CreateLine($iWidthInput + $iWidthLabel, 1)


	;-------------------- EXPIRES ---------------------
	$ctrl = GUICtrlCreateGroup("Срок действия учетной записи", $iPositionLeft, $iPositionTop, _
		$iWidthInput + $iWidthLabel, 84)
	$iPositionTop += 22
	$iPositionLeft += $iGap * 2 - 5

	Local $accountExpiresDate = GetAttributeText("accountExpires")
	Local $aRadioExpire[2]

	$aRadioExpire[0] = GUICtrlCreateRadio("Никогда", $iPositionLeft, $iPositionTop, -1, $iHeightLabel)
	GUICtrlSetState(-1, BitOr($GUI_CHECKED, $GUI_DISABLE))
	$iPositionTop += $iGap + $iHeightInput

	$aRadioExpire[1] = GUICtrlCreateRadio("Истекает:", $iPositionLeft, $iPositionTop, 73, $iHeightLabel)
	GUICtrlSetState(-1, $GUI_DISABLE)

	Local $dateX = $iPositionLeft + 80
	$iWidthInput -= 1
	Local $idDateExpires = GUICtrlCreateDate("", $dateX, $iPositionTop - 4, $iWidthInput + $iWidthLabel - $dateX, 25)
	GUICtrlSetState(-1, $GUI_DISABLE)
	GUICtrlCreateGroup("", -99, -99, 1, 1)

	If $accountExpiresDate Then
		GUICtrlSetState($aRadioExpire[0], $GUI_UNCHECKED)
		GUICtrlSetState($aRadioExpire[1], $GUI_CHECKED)
		GUICtrlSetData($idDateExpires, $accountExpiresDate)

		If _DateDiff('d', _NowCalc(), $accountExpiresDate) < 0 Then $bAccountExpired = True
	EndIf

	$previosPos = ControlGetPos($idDetailedViewForm, "", $ctrl)
	$iPositionLeft = $previosPos[0]
	$iPositionTop = $previosPos[1] + $previosPos[3] + $iGap

	;-------------------- VERICAL LINE ---------------------
	Local $previosPos = ControlGetPos($idDetailedViewForm, "", $ctrl)
	$iPositionLeft = 10 + $iWidthLabel + $iWidthInput + $iGap * 2
	$iPositionTop = 12
	CreateLine(1, $previosPos[1] + $previosPos[3] - $iPositionTop, False)


	;-------------------- BAD PASSWORD ---------------------
	$iWidthInput = 155
	$iWidthLabel = 200

	For $i = 6 To 10
		CreateLabelAndInput($aAccountInfo[$i][0], $aAccountInfo[$i][1], $ES_READONLY)
	Next

	;-------------------- ACCOUN PARAMETERS ---------------------
	Local $aCheckboxes[0][4]
	_ArrayAdd($aCheckboxes, CreateAttributeArray("Требовать смены пароля при следующем входе в систему", _
		GetAttributeText("pwdLastSet") <> "" ? 0 : 1, True))
	_ArrayAdd($aCheckboxes, CreateAttributeArray("Запретить смену пароля пользователем", _
		IsUserCannotChangePassword(GetAttributeText("distinguishedName")), True))
	_ArrayAdd($aCheckboxes, CreateAttributeArray("Срок действия пароля не ограничен", _
		StringInStr(GetAttributeText("userAccountControl"), "DONT_EXPIRE_PASSWD"), True))
	_ArrayAdd($aCheckboxes, CreateAttributeArray("Отключить учетную запись", _
		_AD_IsObjectDisabled($sAMAccountName), True))
	_ArrayAdd($aCheckboxes, CreateAttributeArray("Разблокировать учтеную запись", _
		_AD_IsObjectLocked($sAMAccountName), True))

	$aCheckboxes[4][2] = GUICtrlCreateCheckbox($aCheckboxes[4][0], $iPositionLeft, $iPositionTop, $iWidthInput + $iWidthLabel, $iHeightLabel)
	GUICtrlSetState($aCheckboxes[4][2], $aCheckboxes[4][1] ? $GUI_ENABLE : $GUI_DISABLE)
	$iPositionTop += $iGap * 2 + $iHeightInput - 4

	GUICtrlCreateLabel("Параметры учетной записи:", $iPositionLeft, $iPositionTop + 3, $iWidthLabel + $iWidthInput, $iHeightLabel)
	$iPositionTop += $iGap + $iHeightLabel

	$ctrl = GUICtrlCreateLabel("", $iPositionLeft, $iPositionTop, $iWidthInput + $iWidthLabel, 110, Default, $WS_EX_CLIENTEDGE)
	GUICtrlSetBkColor(-1, 0xf0f0f0)
	$iPositionTop += $iGap
	GUICtrlSetState(-1, $GUI_DISABLE)
	$iPositionLeft += $iGap * 1.5

	For $i = 0 To UBound($aCheckboxes, $UBOUND_ROWS) - 2
		$aCheckboxes[$i][2] = GUICtrlCreateCheckbox($aCheckboxes[$i][0], $iPositionLeft, $iPositionTop, -1, $iHeightLabel)
 		GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
		If $aCheckboxes[$i][1] Then GUICtrlSetState(-1, $GUI_CHECKED)
		$iPositionTop += $iGap + $iHeightLabel
	Next

	;-------------------- VERICAL LINE ---------------------
	$previosPos = ControlGetPos($idDetailedViewForm, "", $ctrl)
	$iPositionLeft = $previosPos[0] + $previosPos[2] + $iGap * 2
	$iPositionTop = 12
	CreateLine(1, $previosPos[1] + $previosPos[3] - $iPositionTop, False)


	;-------------------- LOGIN ---------------------
	$iWidthLabel = 140
	$iWidthInput = 215
	CreateLabelAndInput($aAccountInfo[11][0], $aAccountInfo[11][1], $ES_READONLY)
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

	Local $aRadioPhone[5]
	For $i = 12 To 16
		$aRadioPhone[$i - 12] = GUICtrlCreateRadio($aAccountInfo[$i][0], $iPositionLeft, $iPositionTop, $iWidthLabel)
		Local $style = ($i = 15 ? $ES_READONLY : -1)
		$aAccountInfo[$i][2] = GUICtrlCreateInput($aAccountInfo[$i][1], $iPositionLeft + $iWidthLabel, _
			$iPositionTop + 1, $iWidthInput, $iHeightInput, $style)
		$iPositionTop += $iGap + $iHeightInput
	Next
	GUICtrlSetState($aRadioPhone[0], $GUI_CHECKED)

	$previosPos = ControlGetPos($idDetailedViewForm, "", $ctrl)
	$iPositionTop = $previosPos[1] + $previosPos[3]


	$iPositionTop += $iGap

	;-------------------- EMAIL ---------------------
	$ctrl = GuiCtrlCreateGroup("Адрес электронной почты для уведомления:", $iPositionLeft - $iGap * 2 + 5, $iPositionTop, _
		$iWidthTotal, 3 * $iHeightInput + 2 * $iGap + 4)
	$iPositionTop += 20


	Local $aRadioEmail[2]
	For $i = 17 To 18
		$aRadioEmail[$i - 17] = GUICtrlCreateRadio($aAccountInfo[$i][0], $iPositionLeft, $iPositionTop, $iWidthLabel)
		$aAccountInfo[$i][2] = GUICtrlCreateInput($aAccountInfo[$i][1], $iPositionLeft + $iWidthLabel, $iPositionTop + 1, _
			$iWidthInput, $iHeightInput)
		$iPositionTop += $iGap + $iHeightInput
	Next
	GUICtrlSetState($aRadioEmail[0], $GUI_CHECKED)


	;-------------------- HORIZONTAL LINE ---------------------
	$previosPos = ControlGetPos($idDetailedViewForm, "", $ctrl)
	$iPositionLeft = 10
	$iPositionTop = $previosPos[1] + $previosPos[3] + $iGap * 2
	CreateLine($iWidthGui - $iGap * 2, 1)

	;-------------------- PASSWORD ---------------------
	$iPositionLeft = 318 + 116
	$iWidthInput = 150

	Local $idLabelPass = GUICtrlCreateLabel("Новый пароль:", $iPositionLeft, $iPositionTop + 3, -1, $iHeightLabel)

	$previosPos = ControlGetPos($idDetailedViewForm, "", $idLabelPass)
	Local $idInputPassword = GUICtrlCreateInput(GetNewPassword(), $previosPos[0] + $previosPos[2] + $iGap, _
		$previosPos[1] - 3, $iWidthInput, $iHeightInput, BitOR($ES_PASSWORD, $ES_AUTOHSCROLL))
	Local $bInputPasswordLocked = True
	Local $sDefaultPassChar = GUICtrlSendMsg($idInputPassword, $EM_GETPASSWORDCHAR, 0, 0)

	;-------------------- BUTTONS ---------------------
	$iPositionLeft = 10

	Local $idButtonBack = GUICtrlCreateButton("Закрыть", $iPositionLeft, _
		$previosPos[1] - ($iHeightButton - $previosPos[3]) / 2, $iWidthButton, $iHeightButton)

	$previosPos = ControlGetPos($idDetailedViewForm, "", $idButtonBack)
	Local $idButtonReset = GUICtrlCreateButton("Сбросить и уведомить", $iWidthGui - $iWidthButton * 1.5 - $iGap, _
		$previosPos[1], $iWidthButton * 1.5, $iHeightButton, $BS_MULTILINE)

	$previosPos = ControlGetPos($idDetailedViewForm, "", $idButtonReset)
	Local $idButtonApply = GUICtrlCreateButton("Применить", $previosPos[0] - $iGap - $iWidthButton, _
		$previosPos[1], $iWidthButton, $iHeightButton)

	Local $previosPos2 = ControlGetPos($idDetailedViewForm, "", $idInputPassword)
	Local $idButtonShowPass = GUICtrlCreateButton("Показать", $previosPos2[0] + $previosPos2[2] + $iGap, _
		$previosPos[1], $iHeightButton, $iHeightButton, $BS_ICON)
	GUICtrlSetImage(-1, "shell32.dll", 23, 0)

	Local $bHasUnclockResetRights = _AD_HasUnlockResetRights($sAMAccountName)
	Local $bHasWritePropWrite = _AD_HasRequiredRights($sAMAccountName, $ADS_RIGHT_DS_WRITE_PROP)

	If $bAccountExpired Or Not $bHasUnclockResetRights Or Not $bHasWritePropWrite Then
		GUICtrlSetState($idButtonReset, $GUI_DISABLE)
		GUICtrlSetState($idButtonApply, $GUI_DISABLE)
		GUICtrlSetState($idInputPassword, $GUI_HIDE)
		GUICtrlSetState($idButtonShowPass, $GUI_HIDE)
		GUICtrlSetState($idLabelPass, $GUI_HIDE)
		Local $sTmp = "Истек срок действия учетной записи. Сброс пароля невозможен." & @CRLF & _
			"Продление срока действия учетной записи необходимо согласовать с руководством."
		If Not $bHasUnclockResetRights Or Not $bHasWritePropWrite Then _
			$sTmp = "У вашей учетной записи отсутствую права на редактирование" & @CRLF & _
			"выбранной учетной записи."
		GUICtrlCreateLabel($sTmp, 300, $previosPos[1], $iWidthGui - 600, $previosPos[3], $SS_CENTER)
		GUICtrlSetBkColor(-1, $COLOR_YELLOW)
	EndIf

	GUICtrlDelete($idProgress)
	GUICtrlDelete($idLabelProgress)

	UpdateChiefInfo($aAccountInfo[5][2], $aAccountInfo[15][2], $aAccountInfo[5][1])

	GUISetState(@SW_SHOW)

	Local $sNewChief = ""

	While 1
		Local $nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				_AD_Close()
				Exit
			Case $idButtonBack
				$aAccountProperties = 0
				GUIDelete($idDetailedViewForm)
				Return
			Case $idButtonReset
				Local $sPhoneNumber = GetSelectedPhoneNumber($aRadioPhone, $aAccountInfo, 12)
				Local $sEmailAddress = GetSelectedEmailAddress($aRadioEmail, $aAccountInfo, 17)

				Local $sErrorMessage = ""
				If Not $sPhoneNumber Then
					$sErrorMessage &= "Не выбран номер телефона для отправки СМС"
				ElseIf $sPhoneNumber = -1 Then
					$sErrorMessage &= "Выбранный номер телефона не имеет значения"
				ElseIf $sPhoneNumber = -2 Then
					$sErrorMessage &= "Выбранный номер телефона не соответствует требуемому формату: 9xx-xxx-xx-xx"
				EndIf

				If Not $sEmailAddress Then
					$sErrorMessage &= @CRLF & "Не выбран адрес почты для уведомления"
				ElseIf $sEmailAddress = -1 Then
					$sErrorMessage &= @CRLF & "Выбранный адрес почты не имеет значения"
				ElseIf $sEmailAddress = -2 Then
					$sErrorMessage &= @CRLF & "Выбранный адрес почты не соответствует требуемому формату: x@x.x"
				EndIf

				If $sErrorMessage Then
					$sErrorMessage &= @CRLF & @CRLF & "Изменения не были применены," & @CRLF & _
										"устраните ошибки и попробуйте снова"
					MsgBox(16, "Невозможно отправить уведомление", $sErrorMessage)
					ContinueLoop
				EndIf

				SetCheckboxDefaultState($aCheckboxes)
				ApplyChanges($idDetailedViewForm, $sNewChief, $sAMAccountName, $aAccountInfo, $aCheckboxes, $idInputPassword, _
					$sPhoneNumber, $sEmailAddress, True)
			Case $idButtonApply
				ApplyChanges($idDetailedViewForm, $sNewChief, $sAMAccountName, $aAccountInfo, $aCheckboxes, $idInputPassword)
			Case $idButtonChief
				Local $idTmp1 = $idInputNameToSearch
				Local $idTmp2 = $idButtonSearch
				Local $idTmp3 = $idButtonNext
				Local $idTmp4 = $idListViewResults
				Local $idTmp5 = $idLabelName
				Local $idTmp6 = $idGuiMain

				$sNewChief = FormMainGui(True, $idDetailedViewForm)
				If $sNewChief Then UpdateChiefInfo($aAccountInfo[5][2], $aAccountInfo[15][2], $sNewChief)

				$idInputNameToSearch = $idTmp1
				$idButtonSearch = $idTmp2
				$idButtonNext = $idTmp3
				$idListViewResults = $idTmp4
				$idLabelName = $idTmp5
				$idGuiMain = $idTmp6
			Case $idButtonShowPass
				If $bInputPasswordLocked Then
					GUICtrlSendMsg($idInputPassword, $EM_SETPASSWORDCHAR, 0, 0)
					$bInputPasswordLocked = False
				Else
					GUICtrlSendMsg($idInputPassword, $EM_SETPASSWORDCHAR, $sDefaultPassChar, 0)
					$bInputPasswordLocked = True
				EndIf
				GUICtrlSetState($idInputPassword, $GUI_FOCUS)
				GUICtrlSetState($idButtonShowPass, $GUI_FOCUS)
			Case $aCheckboxes[0][2]
				VerifyCheckboxes($aCheckboxes)
			Case $aCheckboxes[1][2]
				VerifyCheckboxes($aCheckboxes)
			Case $aCheckboxes[2][2]
				VerifyCheckboxes($aCheckboxes)
		EndSwitch
	WEnd

EndFunc




Func ApplyChanges($idParentGui, $sNewChief, $sAMAccountName, ByRef $aAccountInfo, ByRef $aCheckboxes, ByRef $idInputPassword, _
	$sPhoneNumber = False, $sEmailAddress = False, $bNotificate = False)
	GUISetState(@SW_DISABLE, $idParentGui)
	ProgressOn("", "Внесение изменений", "", -1, -1, $DLG_MOVEABLE)
	Local $sResult = ""
	Local $sErrors = ""

	ProgressSet(0, "Проверка атрибутов")
	For $i = 0 To UBound($aAccountInfo, $UBOUND_ROWS) - 1
		If Not $aAccountInfo[$i][2] Or Not $aAccountInfo[$i][3] Then ContinueLoop

		Local $sNewValue = GUICtrlRead($aAccountInfo[$i][2])
		If $aAccountInfo[$i][3] = "manager" Then
			If Not $sNewChief Then ContinueLoop
			$sNewValue = _AD_GetObjectAttribute($sNewChief, "distinguishedName")
		EndIf

		If $sNewValue = $aAccountInfo[$i][1] Then ContinueLoop

		Local $bSuccess = _AD_ModifyAttribute($sAMAccountName, $aAccountInfo[$i][3], $sNewValue, 2)
		If Not $bSuccess Then
			$sErrors &= "Не удалось обновить: '" & $aAccountInfo[$i][3] & "' : '" & _
				$sNewValue & "', код ошибки: " & @error & @CRLF
		Else
			$sResult &= "Обновлен атрибут '" & $aAccountInfo[$i][0] & "', старое значение: '" & _
				$aAccountInfo[$i][1] & "', новое значение: '" & $sNewValue & "'" & @CRLF
			$aAccountInfo[$i][1] = $sNewValue
		EndIf
	Next

	ProgressSet(30, "Проверка параметров")
	For $i = UBound($aCheckboxes, $UBOUND_ROWS) - 1 To 0 Step -1
		Local $bState = (GUICtrlRead($aCheckboxes[$i][2]) = $GUI_CHECKED ? 1 : 0)
		If $bState = $aCheckboxes[$i][1] Then ContinueLoop

		Local $bSuccess = False

		If $i = 0 Then
			$bSuccess = _AD_SetPasswordExpire($sAMAccountName, $bState ? 0 : -1)
		ElseIf $i = 1 Then
			$bSuccess = SetUserCannotChangePassword("budzdorov", $sAMAccountName, $bState)
		ElseIf $i = 2 Then
			$bSuccess = $bState ? _AD_DisablePasswordExpire($sAMAccountName) : _
				_AD_EnablePasswordExpire($sAMAccountName)
		ElseIf $i = 3 Then
			$bSuccess = $bState ? _AD_DisableObject($sAMAccountName) : _
				_AD_EnableObject($sAMAccountName)
		ElseIf $i = 4 Then
			$bSuccess = _AD_UnlockObject($sAMAccountName)
		EndIf

		If Not $bSuccess Then
			$sErrors &= "Не удается изменить атрибут '" & $aCheckboxes[$i][0] & "'" & @CRLF
		Else
			$sResult &= "Установлен атрибут '" & $aCheckboxes[$i][0] & "' в значение " & $bState & @CRLF
			$aCheckboxes[$i][1] = $bState
		EndIf
	Next

	Local $sPassword = GUICtrlRead($idInputPassword)
	If $sPassword Then
		ProgressSet(60, "Установка нового пароля")
		Local $bSuccess = _AD_SetPassword($sAMAccountName, $sPassword)
		If Not $bSuccess Then
			$sErrors &= "Не удается изменить пароль, код ошибки: " & @error & @CRLF
		Else
			$sResult &= "Установлен новый пароль" & @CRLF
			GUICtrlSetData($idInputPassword, "")
		EndIf
	EndIf

	If $sErrors Then
		GUISetState(@SW_ENABLE, $idParentGui)
		ProgressOff()
		If Not $sResult Then $sResult = "Изменений нет" & @CRLF
		$sErrors &= "Уведомления не будут высланы"
		Return MsgBox(0, "", $sErrors, 0, $idParentGui)
	EndIf

	Local $sCurrentUserName = _AD_GetObjectAttribute(@UserName, "displayName")
	Local $sUserName = _AD_GetObjectAttribute($sAMAccountName, "displayName")

	If $bNotificate And $sPhoneNumber Then
		ProgressSet(80, "Отправка СМС уведомления")
		Local $sSmsText = "Учетные данные для входа в систему: " & _
			"пользователь - " & $sAMAccountName & " / " & _
			"пароль - " & $sPassword
		Local $bSuccess = SendSmsNotificationToUser($sSmsText, $sPhoneNumber)
		If Not $bSuccess Then
			$sErrors &= "Не удалось отправить SMS уведомление на номер: " & $sPhoneNumber & @CRLF
		Else
			$sResult &= "SMS уведомление отправлено на номер: " & $sPhoneNumber & @CRLF
		EndIf
	EndIf

	If $bNotificate And $sEmailAddress And $sResult Then
		ProgressSet(90, "Отправка email умедомления")
		Local $sMailMessageToSend = "Уважаемый(ая) " & $sUserName & "," & @CRLF & _
			@CRLF & _
			"Следующие изменения были внесены в Вашу учетную запись:" & @CRLF & _
			$sResult & _
			@CRLF & _
			"Сотрудник, выполнивший данные действия: " & $sCurrentUserName & @CRLF & _
			@CRLF & _
			"Внимание! Если Вы не обращались в службу технической поддержки с данной заявкой, " & @CRLF & _
			"то просьба связаться с сотрудниками техподдержки по телефонам: " & @CRLF & _
			"603, для регионов 30-494"
		Local $sMailTitle = "Изменение данных Вашей учетной записи"

		Local $bSuccess = SendEmail($sMailMessageToSend, $sMailTitle, $sEmailAddress)
		If Not $bSuccess Then
			$sErrors &= "Не удалось отправить email уведомление на адрес: " & $sEmailAddress & @CRLF
		Else
			$sResult &= "Email уведомление отправлено на адрес: " & $sEmailAddress & @CRLF
		EndIf
	EndIf

	If $sResult Then
		ProgressSet(95, "Отправка заявки в СТП")
		Local $sMailMessageToSend = "Внесены изменения в учетную запись '" & $sUserName & "':" & _
			@CRLF & $sResult & @CRLF & $sErrors & @CRLF & "Ответственный сотрудник: " & $sCurrentUserName
		Local $sMailTitle = "Сброс пароля через приложение"
		Local $sMailTo = ""
		Local $bSuccess = SendEmail($sMailMessageToSend, $sMailTitle, $sMailTo)

		If Not $bSuccess Then _
			$sErrors &= "Не удалось отправить заявку на адрес: " & $sMailTo
	Else
		If Not $sResult Then $sResult = "Изменений нет" & @CRLF
	EndIf

	ProgressOff()
	GUISetState(@SW_ENABLE, $idParentGui)
	MsgBox(0, "", $sResult & $sErrors, 0, $idParentGui)
EndFunc


Func SendSmsNotificationToUser($sSmsText, $sPhoneNumber)
	Local $sMailMessageToSend = "<!godmode> " & $sPhoneNumber & " " & $sSmsText
	Local $sMailTitle = "Сброс пароля"
	Local $sMailTo = "msg@7828882.ru"
	Return SendEmail($sMailMessageToSend, $sMailTitle, $sMailTo, True)
EndFunc


Func SetCheckboxDefaultState($aCheckboxes)
	GUICtrlSetState($aCheckboxes[0][2], $GUI_CHECKED)
	GUICtrlSetState($aCheckboxes[1][2], $GUI_UNCHECKED)
	GUICtrlSetState($aCheckboxes[2][2], $GUI_UNCHECKED)
	GUICtrlSetState($aCheckboxes[3][2], $GUI_UNCHECKED)
	If GUICtrlGetState($aCheckboxes[4][2]) <> 144 Then _
		GUICtrlSetState($aCheckboxes[4][2], $GUI_CHECKED)
EndFunc


Func VerifyCheckboxes($aCheckboxes)
	If GUICtrlRead($aCheckboxes[0][2]) <> $GUI_CHECKED Then Return

	If GUICtrlRead($aCheckboxes[1][2]) = $GUI_CHECKED Then
		MsgBox($MB_ICONWARNING, "Доменные службы Active Directory", _
			'Одновременная установка флажков "Требовать смены пароля при следующем "' & _
			'входе в систему и "Запретить смену пароля" не докускается.')
		GUICtrlSetState($aCheckboxes[1][2], $GUI_UNCHECKED)
	EndIf

	If GUICtrlRead($aCheckboxes[2][2]) = $GUI_CHECKED Then
		MsgBox($MB_ICONWARNING, "Доменные службы Active Directory", _
			'Выбран вариант пароля пользователя без истечения срока годности. ' & _
			'Пользователю не потребуется изменять пароль при следующем входе в сеть.')
		GUICtrlSetState($aCheckboxes[0][2], $GUI_UNCHECKED)
	EndIf
EndFunc


Func GetNewPassword()
	Local $sResult = ""
	Local $sCapital = "ABCDEFGHJKLMNOPQRSTUVWXYZ"
	Local $sLower = "abcdefghijklmnopqrstuvwxyz"
	Local $sNumbers = "0123456789"

	Local $iLenght = 8

	$sResult &= StringMid($sCapital, Random(1, StringLen($sCapital), 1), 1)
	For $i = 2 To $iLenght - 1
		$sResult &= StringMid($sLower, Random(1, StringLen($sLower), 1), 1)
	Next
	$sResult &= StringMid($sNumbers, Random(1, StringLen($sNumbers), 1), 1)

	Return $sResult
EndFunc


Func GetSelectedPhoneNumber($aRadioPhone, $aAccountInfo, $iStartIndex)
	If Not IsArray($aRadioPhone) Or Not IsArray($aAccountInfo) Then Return

	For $i = 0 To UBound($aRadioPhone) - 1
		If GUICtrlRead($aRadioPhone[$i]) <> $GUI_CHECKED Then ContinueLoop

		Local $sEnteredValue = GUICtrlRead($aAccountInfo[$iStartIndex + $i][2])
		If Not $sEnteredValue Then Return -1
		$sEnteredValue = GetNormalizedPhoneNumber($sEnteredValue)
		Return $sEnteredValue ? $sEnteredValue : -2
	Next
EndFunc


Func GetSelectedEmailAddress($aRadioEmail, $aAccountInfo, $iStartIndex)
	If Not IsArray($aRadioEmail) Or Not IsArray($aAccountInfo) Then Return

	For $i = 0 To UBound($aRadioEmail) - 1
		If GUICtrlRead($aRadioEmail[$i]) <> $GUI_CHECKED Then ContinueLoop

		Local $sEnteredValue = GUICtrlRead($aAccountInfo[$iStartIndex + $i][2])
		If Not $sEnteredValue Then Return -1
		If StringInStr($sEnteredValue, "@") And StringInStr($sEnteredValue, ".") Then
			Return $sEnteredValue
		Else
			Return -2
		EndIf
	Next
EndFunc


Func UpdateChiefInfo($idEditName, $idEditPhone, $sAMAccountName)
	Local $aTmp = _AD_GetObjectProperties($sAMAccountName, "displayName")
	$aTmp = IsArray($aTmp) ? $aTmp[1][1] : ""
	GUICtrlSetData($idEditName, $aTmp)
	Local $aPhoneAttribute[] = [ _
		"facsimileTelephoneNumber", _
		"homePhone", _
		"ipPhone", _
		"mobile", _
		"otherTelephone", _
		"pager", _
		"telephoneNumber"]

	$aTmp = ""
	For $sAttribute in $aPhoneAttribute
		Local $aAttributes = _AD_GetObjectProperties($sAMAccountName, $sAttribute)
		If Not IsArray($aAttributes) Or UBound($aAttributes, $UBOUND_ROWS) < 2 Then ContinueLoop
		If GetNormalizedPhoneNumber($aAttributes[1][1]) Then
			$aTmp = $aAttributes[1][1]
			ExitLoop
		EndIf
	Next

	GUICtrlSetData($idEditPhone, $aTmp)
EndFunc


Func GetNormalizedPhoneNumber($sNumber)
	If StringLen($sNumber) < 10 Then Return

	Local $sResult = ""
	For $i = 1 To StringLen($sNumber)
		Local $sSymbol = StringMid($sNumber, $i, 1)

		If Not StringIsDigit($sSymbol) Then ContinueLoop
		If Not $sResult And $sSymbol <> "9" Then ContinueLoop

		$sResult &= $sSymbol
	Next

	If StringLen($sResult) = 10 Then Return $sResult
EndFunc


Func SetUserCannotChangePassword($sDomain, $sUser, $bCannotChange)
	Local $strPath = "WinNT://" & $sDomain & "/" & $sUser
	Local $oUser = ObjGet($strPath)
	If Not IsObj($oUser) Then Return

    Local $lUserFlags = $oUser.Get("userFlags")

    If $bCannotChange Then
        $lUserFlags = BitOr($lUserFlags, 64)
    Else
        $lUserFlags = BitAND($lUserFlags, BitNOT(64))
    EndIf

    $oUser.Put("userFlags", $lUserFlags)
    $oUser.SetInfo
	If Not @error Then Return True
EndFunc


Func IsUserCannotChangePassword($userDN)
	If Not $userDN Then Return -1

	$userDN = "LDAP://" & $userDN
	Local $oUser = ObjGet($userDN)
	If Not IsObj($oUser) Then Return -1

	Local $oSecDesc = $oUser.Get("ntSecurityDescriptor")
	If Not IsObj($oSecDesc) Then Return -1

	Local $oACL = $oSecDesc.DiscretionaryAcl
	If Not IsObj($oACL) Then Return -1

	Local $fSelf = False
	Local $fEveryone = False
	For $oACE in $oACL
		If $oACE.ObjectType <> "{AB721A53-1E2F-11D0-9819-00AA0040529B}" Then ContinueLoop

		If $oACE.Trustee = "Все" And $oACE.AceType = 6 Then $fEveryone = True
		If $oACE.Trustee = "NT AUTHORITY\SELF" And $oACE.AceType = 6 Then $fSelf = True
	Next

	If $fSelf And $fEveryone Then Return 1
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
;~ 		ButtonNextPressed()
		$bNextPressed = True
	EndIf

	$tagNMHDR = 0
	$event = 0
	$lParam = 0
EndFunc


Func My_WM_COMMAND($hWnd, $imsg, $iwParam, $ilParam)
    Local $nNotifyCode = BitShift($iwParam, 16)
    Local $hCtrl = $ilParam
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
	Local $aReturn[1][4]
	$aReturn[0][0] = $name
	$aReturn[0][1] = $plainText ? $attribute : GetAttributeText($attribute)
	$aReturn[0][2] = 0
	$aReturn[0][3] = $attribute
	Return $aReturn
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


Func CreateLabelAndInput($textLabel, $textInput, $style = "", $attribute = "", $isButtonPresent = False)
	GUICtrlCreateLabel($textLabel, $iPositionLeft, $iPositionTop + 3, $iWidthLabel, $iHeightLabel)

	Local $ret = 0

	If Not $attribute Then
		$ret = GUICtrlCreateInput($textInput, $iPositionLeft + $iWidthLabel, $iPositionTop, _
			$isButtonPresent ? ($iWidthInput - $iGap * 4) : $iWidthInput, $iHeightInput, BitOr($style, $ES_AUTOHSCROLL))
	Else
		$ret = GUICtrlCreateCombo("", $iPositionLeft + $iWidthLabel, $iPositionTop, _
		$iWidthInput, $iHeightInput);, BitOr($style, $CBS_AUTOHSCROLL))
		GUICtrlSetData(-1, _ArrayToString(GetPromptArray($attribute)), $textInput)
	EndIf

	$iPositionTop += $iGap + $iHeightInput
	Return $ret
EndFunc


Func GetPromptArray($name)
	Local $aReturn = _ArrayUnique(_AD_GetObjectsInOU("", _
		"(&(objectCategory=person)(objectClass=user)(name=*))", 2, $name))
	_ArraySort($aReturn)
	Return $aReturn
EndFunc


Func GetAttributeText($attributeName)
;~ 	_ArrayDisplay($aAccountProperties)
	Local $searchResults = _ArraySearch($aAccountProperties, $attributeName)
	Local $ret = $searchResults <> -1 ? $aAccountProperties[$searchResults][1] : ""
	If Not $ret Or $ret = "1601/01/01 00:00:00" Or _
		$ret = "0000/00/00 00:00:00" Then $ret = ""
;~ 	ConsoleWrite($attributeName & " : " & $ret & @CRLF)
	Return $ret
EndFunc


Func SendEmail($sMailMessageToSend, $sMailTitle, $sMailTo, $bIsSms = False)
	Local $sMailServer = ""
	Local $sMailFrom = "Система сброса паролей"
	Local $sMailLogin = ""
	Local $sMailPassword = ""

	If Not $bIsSms Then $sMailMessageToSend &= @CRLF & @CRLF & _
											"----------------------------------------------" & @CRLF & _
											"Данное сообщение было отправлено автоматически," & @CRLF & _
											"просьба не отвечать на него"

	Return _INetSmtpMailCom($sMailServer, $sMailFrom, $sMailLogin, $sMailTo, $sMailTitle, _
		$sMailMessageToSend, "", "", "", $sMailLogin, $sMailPassword)
EndFunc   ;==>SendEmail


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", _
	$as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Username = "", _
	$s_Password = "", $IPPort = 25, $ssl = 0)

	Local $objEmail = ObjCreate("CDO.Message")
	Local $i_Error = 0
	Local $i_Error_desciption = ""

	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress

	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

	$objEmail.Subject = $s_Subject

 	$objEmail.Textbody = $as_Body
	$objEmail.TextBodyPart.Charset = "utf-8"

	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf

	If $ssl Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	EndIf

	$objEmail.Configuration.Fields.Update
	$objEmail.Send

	If @error Then Return

	Return True
EndFunc   ;==>_INetSmtpMailCom
