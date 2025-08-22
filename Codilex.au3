#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Data\Image\Icon.ico
#AutoIt3Wrapper_Outfile=Release\Codilex.exe
#AutoIt3Wrapper_Outfile_x64=Release\Codilex_x64.exe
#AutoIt3Wrapper_Compile_Both=y
#AutoIt3Wrapper_Res_Description=Codilex
#AutoIt3Wrapper_Res_Fileversion=1.0
#AutoIt3Wrapper_Res_LegalCopyright=Copyright (C) 2020-2025 Zaibai Software Production
#AutoIt3Wrapper_Res_Language=1036
#AutoIt3Wrapper_Res_Field=Nom Interne|Codilex
#AutoIt3Wrapper_Res_Field=Créer par|Zaibai
#AutoIt3Wrapper_Res_Field=Email|erwan-91310@hotmail.fr
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;Include
#include "Data\Include\GuiConstructor\GuiConstructor.au3"
#include <ProgressConstants.au3>
#include <EditConstants.au3>
#include <GuiListView.au3>
#include <GuiToolTip.au3>
#include <WinAPIDlg.au3>
#include <GuiMenu.au3>
#include <Misc.au3>

;Option
#NoTrayIcon
OnAutoItExitRegister("_Quit")
Opt("WinTitleMatchMode", 3) ;Exact title match (pour If ProcessExists($PIDini) And WinExists($TITLE))
Opt("GUIOnEventMode", 1)

;Variable Constante
Dim $TITLE = "Codilex"
Dim $VERSION = "1.0"

;Variable Constante - Fichier
Dim $DIR_INSTALL = @AppDataDir & "\" & $TITLE & "\"
Dim $DIR_DATA = $DIR_INSTALL & "Data\"
Dim $DIR_IMAGE = $DIR_DATA & "Image\"
Dim $DIR_DICO = $DIR_DATA & "Dico\"
Dim $INI_SETTING = $DIR_INSTALL & "paramètres.ini"

;Check si une instance du programme est déjà en cours
Dim $bPIDExit = False
Local $PIDini = IniRead($INI_SETTING, "Paramètres", "PID", "null")
If ProcessExists($PIDini) And WinExists($TITLE) Then
	If MsgBox(4 + 32 + 256 + 262144, "Programme déjà lancé", 'Nous avons détecté que ' & $TITLE & ' est déjà lancé.' & @CRLF & $TITLE & ' est limité à une instance. Voulez-vous arrêter la première instance ?') = 7 Then
		$bPIDExit = True
		Exit
	EndIf
	ProcessClose($PIDini)
EndIf
IniWrite($INI_SETTING, "Paramètres", "PID", @AutoItPID)

;Check Version/Installation
Local $CheckVersion = IniRead($INI_SETTING, "Paramètres", "Version", "null")
If $CheckVersion <> $VERSION Or Not FileExists($DIR_INSTALL) Then _Install()

;Chargement du fichier $INI_SETTING
Dim $Ini_XPos = IniRead($INI_SETTING, "Paramètres", "XPos", -1)
Dim $Ini_YPos = IniRead($INI_SETTING, "Paramètres", "YPos", -1)
Dim $Ini_Flag = IniRead($INI_SETTING, "Paramètres", "Langue", "FR")

;Variable utilisé dans diverses fonctions
Global $g_bSortSense = True
Dim $bAccentClear = False, $bfInterrupt = False, $bLVContextMenu = False, $bLVFocus = False
Dim $iEtapeTT = 1

;Variable - Tableau GUI
Global $aGFlag[5][2] = [[0, "FR"], [0, "EN"], [0, "DE"], [0, "IT"], [0, "ES"]] ;Ctrl_ID_Flag | Code Pays
Global $aGCtrl[4][4] = [[0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0]] ;LabelID | InputID | IconErrorID | ToolTipErrorID
Global $aGIni[9][3] = [["", "LetterNBR", "Nombre de lettres:"], ["", "LetterFIRST", "Première lettre:"], ["", "LetterEXP", "Contient l'expression:"], ["", "LetterCONTAINS", "Peut contenir les lettres:"], _ ;Cettte ligne est lié aux LabelID de $aGCtrl.
						["", "ChecboxAccent", "Présence d'accent inconnu"], ["", "ChecboxAccentTips", "Augmente le temps de la recherche."], _
						["", "ButtonCancel", "Annuler"], _
						["", "MenuCopy", "Copier"], ["", "MenuAll", "Tout sélectionner"]] ;Valeur | Clé | Valeur par défaut
Global $aGToolTipIni[4][3] = [["", "LetterNBRErreur", "Vous ne pouvez entrer qu'un nombre ici."], ["", "LetterFIRSTErreur", "Vous ne pouvez entrer qu'une lettre ici."], _
								["", "LetterEXPErreur", "Vous ne pouvez entrer qu'une expression ici. (Minimum 2 lettres)"], _
								["", "LetterCONTAINSErreur", "Vous ne pouvez entrer que des lettres ici. (Minimum 2 lettres)"]] ;Valeur | Clé | Valeur par défaut
_Update_aGIni()
_Update_aGToolTipIni()

;Création de la GUI
Global $GUIP = GUICreate($TITLE, 349, 319, $Ini_XPos, $Ini_YPos)
GUISetFont(8.5 * _RatioFont()[0])
GUISetIcon($DIR_IMAGE & "Icon.ico")
GUISetOnEvent($GUI_EVENT_CLOSE, "_GuiClose")
GUICtrlCreateGraphic(-1, 40, 351, 4, $SS_SUNKEN)
GUICtrlSetState(-1, $GUI_DISABLE)
Global $ProgressBar = GUICtrlCreateProgress(-2, 40, 353, 3, $PBS_SMOOTH)

;GUI - Flag/Langage
For $i = 0 To UBound($aGFlag)-1
	$aGFlag[$i][0] = GUICtrlCreateIcon($DIR_IMAGE & $aGFlag[$i][1] & ".ico", -1, 8+(40*$i), 8, 24, 24)
	GUICtrlSetTip(-1, $aGFlag[$i][1])
	_DisableTabFocus(GUICtrlGetHandle(-1))
	GUICtrlSetOnEvent(-1, "_FlagClick")
Next

;GUI - Listview
Global $hListView = _GUICtrlListView_Create($GUIP, " ", 0, 44, 349, 238, BitOR($LVS_REPORT, $LVS_SHOWSELALWAYS, $LVS_SORTASCENDING))
_GUICtrlListView_SetExtendedListViewStyle($hListView, BitOR($LVS_EX_DOUBLEBUFFER, $LVS_EX_FULLROWSELECT))
Global $hHeaderG = _GUICtrlListView_GetHeader($hListView)
_LVSetState(True)

;GUI - Listview - Context Menu
Global Enum $eMenuCopy = 1000, $eMenuAll
Global $hContextMenu = _GUICtrlMenu_CreatePopup()
_GUICtrlMenu_InsertMenuItem($hContextMenu, 0, $aGIni[7][0], $eMenuCopy)
_GUICtrlMenu_InsertMenuItem($hContextMenu, 1, $aGIni[8][0], $eMenuAll)

;GUI - Listview - Dummy (Ctrl+c et Ctrl+a dans la LV)
Global $idDummyCtrlc = GUICtrlCreateDummy()
GUICtrlSetOnEvent(-1, "_DummyCtrlC")
Global $idDummyCtrla = GUICtrlCreateDummy()
GUICtrlSetOnEvent(-1, "_DummyCtrlA")

;GUI - Input and Cie
For $i = 0 To UBound($aGCtrl)-1
	$iTop = 56 * $i
	$aGCtrl[$i][0] = GUICtrlCreateLabel($aGIni[$i][0], 7, 64 + $iTop, 175, 45, $SS_RIGHT)
	GUICtrlSetFont(-1, 10, 400, 0, "MS Sans Serif")
	$aGCtrl[$i][1] = GUICtrlCreateInput("", 192, 64 + $iTop, 121, 21, BitOR($GUI_SS_DEFAULT_INPUT, $ES_UPPERCASE), -1)
	$aGCtrl[$i][2] = GUICtrlCreateIcon($DIR_IMAGE & "Error.ico", -1, 318, 60 + $iTop, 24, 24)
	GUICtrlSetState(-1, $GUI_HIDE)
	_DisableTabFocus(GUICtrlGetHandle(-1))
	$aGCtrl[$i][3] = _GUIToolTip_Create($GUIP, BitOR($_TT_ghTTDefaultStyle, $TTS_BALLOON))
Next
_Update_ToolTip()

;GUI - Checkbox
Global $ChecboxAccent = GUICtrlCreateCheckbox($aGIni[4][0], 7, 288, 255, 25)
GUICtrlSetTip(-1, $aGIni[5][0])
_DisableTabFocus(GUICtrlGetHandle(-1))
GUICtrlSetOnEvent(-1, "_CheckboxAccentClick")

;GUI - Bouton
Global $ButtonOK = GUICtrlCreateButton("OK", 264, 289, 75, 25)
GUICtrlSetOnEvent(-1, "_ButtonOKClick")
Global $ButtonCancel = GUICtrlCreateButton($aGIni[6][0], 264, 289, 75, 25)
GUICtrlSetOnEvent(-1, "_ButtonCancelClick")
GUICtrlSetState(-1, $GUI_HIDE)

_NoFocusLines_Set($ChecboxAccent)
_NoFocusLines_Set($ButtonOK)
_NoFocusLines_Set($ButtonCancel)

;Lancement de la GUI
GUISetState(@SW_SHOW)
GUICtrlSetState($aGCtrl[0][1], $GUI_FOCUS)
_GuiHotkey()
GUIRegisterMsg($WM_COMMAND, "_WMCommand")
GUIRegisterMsg($WM_NOTIFY, "_WMNotify")
GUIRegisterMsg($WM_MOVE, "_WMGuiMoveSave")

;Boucle Principale
While 1
	Sleep(20)
	If $bLVContextMenu Then _LVContextMenu()
	If $bLVFocus Then _LVFocus()
Wend

;Fonction
Func _FlagClick()
	Local $iIndex = _ArraySearch($aGFlag, @GUI_CtrlId)
	If Not @error Then _ChangeLangage($aGFlag[$iIndex][1])
EndFunc

Func _CheckboxAccentClick()
	$bAccentClear = False
	If GUICtrlRead($ChecboxAccent) = $GUI_CHECKED Then $bAccentClear = True
EndFunc

Func _ButtonOKClick()
	If Not ControlCommand($GUIP, "", $hListView, "IsEnabled") Then
		$bfInterrupt = False
		If Not _CheckInput() Then Return
		Local $aResult = _SearchWord()
		If Not IsArray($aResult) Then
			_ButtonShow() ;;; Changer ça ?
		Else
			_DisplayResult($aResult)
		EndIf
	Else
		GUICtrlSetData($ProgressBar, 0)
		_LVSetState()
		For $i = 0 To UBound($aGCtrl)-1
			ControlShow($GUIP, "", $aGCtrl[$i][0])
			ControlShow($GUIP, "", $aGCtrl[$i][1])
		Next
		GUICtrlSetState($aGCtrl[0][1], $GUI_FOCUS)
		_GUICtrlListView_DeleteAllItems($hListView)
	EndIf
EndFunc

Func _ButtonCancelClick()
	GUICtrlSetData($ProgressBar, 0)
	If Not ControlCommand($GUIP, "", $hListView, "IsVisible") Then _GUICtrlListView_DeleteAllItems($hListView)
	_ButtonShow()
EndFunc

Func _ButtonShow()
	Local $ButtonHide = $ButtonOK, $ButtonShow = $ButtonCancel
	If ControlCommand($GUIP, "", $ButtonCancel, "IsVisible") Then
		$ButtonHide = $ButtonCancel
		$ButtonShow = $ButtonOK
	EndIf
	GUICtrlSetState($ButtonHide, $GUI_HIDE)
	GUICtrlSetState($ButtonShow, $GUI_SHOW)
	_GuiHotkey()
EndFunc

Func _GuiHotkey()
	If ControlCommand($GUIP, "", $ButtonOK, "IsVisible") Then
		Dim $aAccelKeys[1][2] = [["{ENTER}", $ButtonOK]]
		If ControlCommand($GUIP, "", $hListView, "IsEnabled") Then Dim $aAccelKeys[3][2] = [["{ENTER}", $ButtonOK], ["^c", $idDummyCtrlc], ["^a", $idDummyCtrla]]
	Else
		Dim $aAccelKeys[1][2] = [["{ENTER}", $ButtonCancel]]
	EndIf
	GUISetAccelerators($aAccelKeys, $GUIP)
EndFunc

Func _LVContextMenu($sHotKey = "")
	If Not $sHotKey Then
		$bLVContextMenu = False
		Switch _GUICtrlMenu_TrackPopupMenu($hContextMenu, $hListView, -1, -1, 1, 1, 2)
			Case $eMenuCopy
				$sHotKey = "Copy"

			Case $eMenuAll
				$sHotKey = "SelectAll"
		EndSwitch
	EndIf

	;All
	If $sHotKey = "SelectAll" Then _GUICtrlListView_SetItemSelected($hListView, -1)

	;Copy
	If $sHotKey = "Copy" Then
		Local $sCopy = ""
		Local $iSelectedItem = _GUICtrlListView_GetSelectedCount($hListView)
		If Not $iSelectedItem Then Return
		If $iSelectedItem > 1000 Then
			GUICtrlSetData($ProgressBar, 0)
			_ButtonShow()
			_LVSetState(False, False)
		EndIf

		Local $aIndexSelectedItem = _GUICtrlListView_GetSelectedIndices($hListView, True)
		For $i = 1 To $aIndexSelectedItem[0]
			$sCopy &= _GUICtrlListView_GetItemTextString($hListView, $aIndexSelectedItem[$i]) & @CRLF
			If $iSelectedItem > 1000 Then
				If Not $bfInterrupt Then
					_ProgressBarLoad($i, $aIndexSelectedItem[0])
				Else
					_LVSetState()
					$bfInterrupt = False
					Return
				EndIf
			EndIf
		Next

		If $iSelectedItem > 1000 Then
			Local $hTimer, $iWait = 700
			$hTimer = TimerInit()
		EndIf
		ClipPut(StringStripWS($sCopy, 2))
		If $iSelectedItem > 1000 Then
			If TimerDiff($hTimer) < $iWait Then Sleep($iWait-TimerDiff($hTimer)) ;Cette pause permet d'attendre la fin de l'animation de la progressbar avant d'afficher le résultat.
			If Not $bfInterrupt Then
				_ButtonShow()
			Else
				ClipPut("")
				$bfInterrupt = False
			EndIf
			_LVSetState()
		EndIf
	EndIf
EndFunc

Func _DummyCtrlC()
	If ControlCommand($GUIP, "", $hListView, "IsEnabled") Then _LVContextMenu("Copy")
EndFunc

Func _DummyCtrlA()
	If ControlCommand($GUIP, "", $hListView, "IsEnabled") Then _LVContextMenu("SelectAll")
EndFunc

Func _LVSetState($bCreate = False, $bHide = True)
	If ControlCommand($GUIP, "", $hListView, "IsEnabled") Or $bCreate Then
		If $bHide Then ControlHide($GUIP, "", $hListView)
		ControlDisable($GUIP, "", $hListView)
	Else
		If Not ControlCommand($GUIP, "", $hListView, "IsVisible") Then ControlShow($GUIP, "", $hListView)
		ControlEnable($GUIP, "", $hListView)
		$bLVFocus = True
	EndIf
	If Not $bCreate Then _GuiHotkey()
EndFunc

Func _LVFocus()
	$bLVFocus = False
	If ControlGetHandle($GUIP, "", ControlGetFocus($GUIP)) <> $hListView Then ControlFocus($GUIP, "", $hListView)
EndFunc

Func _SetToolTip($hToolTip, $Title, $Text, $ControlID)
	_GUIToolTip_SetTitle($hToolTip, $Title, 3)
	_GUIToolTip_AddTool($hToolTip, $GUIP, $Text, GUICtrlGetHandle($ControlID))
	_GUIToolTip_SetDelayTime($hToolTip, $TTDT_AUTOPOP, 30000)
EndFunc

Func _ChangeLangage($Flag)
	If $Ini_Flag = $Flag Then Return
	$Ini_Flag = $Flag
	IniWrite($INI_SETTING, "Paramètres", "Langue", $Ini_Flag)

	;Update des array contenant les textes
	_Update_aGIni()
	_Update_aGToolTipIni()

	;Updates des ctrls
	For $i = 0 To UBound($aGIni)-1
		Switch $aGIni[$i][1]
			Case "ChecboxAccentTips"
				GUICtrlSetTip($ChecboxAccent, $aGIni[$i][0])
			Case "MenuCopy"
				_GUICtrlMenu_SetItemText($hContextMenu, 0, $aGIni[$i][0])
			Case "MenuAll"
				_GUICtrlMenu_SetItemText($hContextMenu, 1, $aGIni[$i][0])
			Case "ButtonCancel"
				ControlSetText($GUIP, "", $ButtonCancel, $aGIni[$i][0])
			Case "ChecboxAccent"
				ControlSetText($GUIP, "", $ChecboxAccent, $aGIni[$i][0])
			Case Else
				ControlSetText($GUIP, "", $aGCtrl[$i][0], $aGIni[$i][0])
		EndSwitch
	Next
	If ControlCommand($GUIP, "", $hListView, "IsVisible") Then _LVHeaderUpdate()
	_Update_ToolTip()
EndFunc

Func _Update_aGIni()
	For $i = 0 To UBound($aGIni)-1
		$aGIni[$i][0] = IniRead($INI_SETTING, $Ini_Flag, $aGIni[$i][1], $aGIni[$i][2])
	Next
EndFunc

Func _Update_aGToolTipIni()
	For $i = 0 To UBound($aGToolTipIni)-1
		$aGToolTipIni[$i][0] = IniRead($INI_SETTING, $Ini_Flag, $aGToolTipIni[$i][1], $aGToolTipIni[$i][2])
	Next
EndFunc

Func _Update_ToolTip()
	$TTTitle = IniRead($INI_SETTING, $Ini_Flag, "TitreErreur", "Caractère non autorisé")
	For $i = 0 To UBound($aGCtrl)-1
		_SetToolTip($aGCtrl[$i][3], $TTTitle, $aGToolTipIni[$i][0], $aGCtrl[$i][2])
	Next
EndFunc

Func _LVHeaderUpdate()
	Local $iW = 332
	If _GUICtrlListView_GetItemCount($hListView) < 11 Then $iW = 349
	_GUICtrlListView_SetColumn($hListView, 0, IniRead($INI_SETTING, $Ini_Flag, "ColumnHeaderText", "Résultat") & " " & _GUICtrlListView_GetItemCount($hListView), $iW, 2)
EndFunc

Func _ProgressBarLoad($iNbr, $iMax, $iEtape = 0, $iEtapeTT = 1)
	Local $iProgressPercent = GUICtrlRead($ProgressBar)
	Local $iNewProgressPercent = Floor((($iNbr + ($iMax * $iEtape)) * 100) / ($iMax * $iEtapeTT))
	If $iNewProgressPercent > $iProgressPercent Then GUICtrlSetData($ProgressBar, $iNewProgressPercent)
	;ConsoleWrite($iProgressPercent & " -> " & $iNewProgressPercent & @CRLF)
EndFunc

Func _InputEXPError()
	AdlibUnRegister("_InputEXPError")
	GUICtrlSetState($aGCtrl[2][2], $GUI_SHOW)
EndFunc

Func _CheckInput()
	$bNoInput = True
	For $i = 0 To UBound($aGCtrl)-1
		If ControlCommand($GUIP, "", $aGCtrl[$i][2], "IsVisible") Or $bfInterrupt Then Return 0
		If GUICtrlRead($aGCtrl[$i][1]) <> "" Then $bNoInput = False
	Next

	If $bNoInput Then Return 0
	Return 1
EndFunc

Func _ErrorMsgResult()
	ConsoleWrite("ErrorMsgResult" & @CRLF)
	;;;Rajouter un texte d'erreur par langue !
	;;;MsgBox(16, $TITLE & "Error, "error pendant search. dico.ini corrompu ??, 0, $GUIP)
	;;;Remettre l'interface de base ! (label/input etc)
EndFunc

Func _SearchWord()
	Local $fDico = $DIR_DICO & $Ini_Flag & ".ini"
	Local $aResult[1] = [-1], $aTemp, $aSectionName
	Local $iEtape = -1
	$iEtapeTT = 1
	_ButtonShow()

	;Nombre d'étapes total pour le calcul de la progressbar
	For $i = 0 To UBound($aGCtrl)-1
		If GUICtrlRead($aGCtrl[$i][1]) <> "" Then $iEtapeTT += 1
	Next

	For $i = 0 To UBound($aGCtrl)-1
		If $bfInterrupt Then Return 0

		$Input = GUICtrlRead($aGCtrl[$i][1])
		If $Input = "" Then ContinueLoop
		$iEtape += 1

		If $aGCtrl[$i][1] = $aGCtrl[0][1] Then ;Nombre de lettres
			$aResult = _IniReadSectionEx($fDico, $Input) ;Retourne 0 en cas d'erreur ou un tableau avec le nombre de mot trouvé à 0 si la section n'existe pas.
			If Not IsArray($aResult) Then _ErrorMsgResult()
		ElseIf $aResult[0] = -1 Then
			$aResult[0] = 0
			$aSectionName = IniReadSectionNames($fDico)
			If Not IsArray($aSectionName) Then _ErrorMsgResult()

			For $j = 1 To $aSectionName[0]
				$aTemp = _IniReadSectionEx($fDico, $aSectionName[$j])
				If Not IsArray($aTemp) Then _ErrorMsgResult()
				$aTemp = _SearchWordSubRoutine($aGCtrl[$i][1], $Input, $aTemp)
				If UBound($aTemp) > 0 Then
					$aResult[0] += $aTemp[0]
					_ArrayDelete($aTemp, 0)
					_ArrayAdd($aResult, $aTemp)
				EndIf

				If $bfInterrupt Then Return 0
				_ProgressBarLoad($j, $aSectionName[0], $iEtape, $iEtapeTT)
			Next
		Else
			$aResult = _SearchWordSubRoutine($aGCtrl[$i][1], $Input, $aResult)
		EndIf

		If $bfInterrupt Then Return 0
		If $aResult[0] = 0 Then Return $aResult
		_ProgressBarLoad($aResult[0], $aResult[0], $iEtape, $iEtapeTT)
	Next

	If $bfInterrupt Then Return 0
	Return $aResult
EndFunc

Func _SearchWordSubRoutine($Ctrlid, $Input, $aArray)
	Local $aResult[0]
	For $i = 1 To $aArray[0]
		$sWord = $aArray[$i]
		If $bAccentClear Then $sWord = _ReplaceAlphaAccents($sWord)

		Switch $Ctrlid
			Case $aGCtrl[1][1] ;Première lettre
				If StringLeft($sWord, 1) = $Input Then _ArrayAdd($aResult, $aArray[$i])

			Case $aGCtrl[2][1] ;Contient l'expression
				If StringRegExp($sWord, "(?i)(" & $Input & ")") Then _ArrayAdd($aResult, $aArray[$i])

			Case $aGCtrl[3][1] ;Peut contenir les lettres
				If _SearchWordLetterContains($sWord, $Input) Then _ArrayAdd($aResult, $aArray[$i])

		EndSwitch

		If $bfInterrupt Then Return 0
	Next
	_ArrayInsert($aResult, 0, UBound($aResult))
	Return $aResult
EndFunc

Func _DisplayResult($aResult)
	Local $hTimer, $iWait = 700, $iEtape = $iEtapeTT - 1

	If $aResult[0] > 0 Then
		For $i = 1 To $aResult[0]
			If $bfInterrupt Then Return 0
			_GUICtrlListView_AddItem($hListView, $aResult[$i])
			_ProgressBarLoad($i, $aResult[0], $iEtape, $iEtapeTT)
		Next
	EndIf

	$hTimer = TimerInit()
	_LVHeaderUpdate()
	_GUICtrlHeader_SetItemFormat($hHeaderG, 0, BitOR(_GUICtrlHeader_GetItemFormat($hHeaderG, 0), $HDF_SORTUP))
	If TimerDiff($hTimer) < $iWait Then Sleep($iWait-TimerDiff($hTimer)) ;Cette pause permet d'attendre la fin de l'animation de la progressbar avant d'afficher le résultat.
	If $bfInterrupt Then Return 0

	_ButtonShow()
	For $i = 0 To UBound($aGCtrl)-1
		ControlHide($GUIP, "", $aGCtrl[$i][0])
		ControlHide($GUIP, "", $aGCtrl[$i][1])
	Next
	_LVSetState()
EndFunc

Func _IniReadSectionEx($File, $vSection)
	Local $iSize, $aResult[1] = [0], $aTemp
	$iSize = FileGetSize($File) / 1024 ;Result in kilobyte (KiloOctet)
	If @error Then Return 0
	If $iSize <= 31 Then
		$aTemp = IniReadSection($File, $vSection)
		If @error Then Return $aResult
		$aTemp[0][1] = $aTemp[0][0]
		_ArrayColDelete($aTemp, 0, True)
	Else
		Local $hFile, $sFRead
		$hFile = FileOpen($File, 0)
		If $hFile = -1 Then Return 0
		$sFRead = @CRLF & FileRead($hFile) & @CRLF & '['
		FileClose($hFile)
		$vSection = StringStripWS($vSection, 7)
		$aTemp = StringRegExp($sFRead, '(?s)(?i)\n\s*\[\s*' & $vSection & '\s*\]\s*\r\n(.*?)\[', 3)
		If Not IsArray($aTemp) Then Return $aResult
		$aTemp = StringRegExp(@LF & $aTemp[0], '\n\s*.*?\s*=(.*?)\r', 3)
		_ArrayInsert($aTemp, 0, UBound($aTemp))
	EndIf
	$aResult = $aTemp
	Return $aResult
EndFunc

Func _ReplaceAlphaAccents($sWord)
	Local $aAccentAlpha[20][2] = [ _
			["a", "ÁáÀàÂâǍǎĂăÃãẢảẠạÄäÅåĀāĄąẤấẦầẪẫẨẩẬậẮắẰằẴẵẲẳẶặǺǻ"], _
			["c", "ĆćĈĉČčĊċÇç"], _
			["d", "ĎďĐđÐ"], _
			["e", "ÉéÈèÊêĚěĔĕẼẽẺẻĖėËëĒēĘęẾếỀềỄễỂểẸẹỆệ"], _
			["g", "ĞğĜĝĠġĢģ"], _
			["h", "ĤĥĦħ"], _
			["i", "ÍíÌìĬĭÎîǏǐÏïĨĩĮįĪīỈỉỊị"], _
			["j", "Ĵĵ"], _
			["k", "Ķķ"], _
			["l", "ĹĺĽľĻļŁłĿŀ"], _
			["n", "ŃńŇňÑñŅņ"], _
			["o", "ÓóÒòŎŏÔôỐốỒồỖỗỔổǑǒÖöŐőÕõØøǾǿŌōỎỏƠơỚớỜờỠỡỞởỢợỌọỘộ"], _
			["p", "ṔṕṖṗ"], _
			["r", "ŔŕŘřŖŗ"], _
			["s", "ŚśŜŝŠšŞş"], _
			["t", "ŤťŢţŦŧ"], _
			["u", "ÚúÙùŬŭÛûǓǔŮůÜüǗǘǛǜǙǚǕǖŰűŨũŲųŪūỦủƯưỨứỪừỮữỬửỰựỤụ"], _
			["w", "ẂẃẀẁŴŵẄẅ"], _
			["y", "ÝýỲỳŶŷŸÿỸỹỶỷỴỵ"], _
			["z", "ŹźŽžŻż"]]
	For $i = 0 To UBound($aAccentAlpha)-1
		$sWord = StringRegExpReplace($sWord, "[" & $aAccentAlpha[$i][1] & "]", $aAccentAlpha[$i][0])
	Next
	Return $sWord
EndFunc

Func _SearchWordLetterContains($sWord, $Input)
	If $sWord = "" Then Return 0
	Local $sLetter

	;Remplace les caractères spéciaux par rien
	$sWord = StringRegExpReplace($sWord, "[-_' ]", "") ;Générique
	If $Ini_Flag = "ES" Then $sWord = StringRegExpReplace($sWord, "[—‡œ’]", "") ;ES

	For $i = 1 To StringLen($sWord)
		$sLetter = StringMid($sWord, $i, 1) ;Récup la première lettre du mot, puis la deuxième etc.
		If StringInStr($Input, $sLetter) Then ;Cherche si dans les lettres proprosées il y a la lettre $sLetter
			$Input = StringReplace($Input, $sLetter, "", 1) ;Si la condition est vrai, remplace la lettre $sLetter par rien dans les lettres proposés ($Input)
		Else
			Return 0
		EndIf
	Next
	Return 1 ;Si toute les lettres du mot en cours sont validé, la fonction Return 1
EndFunc

Func _Install()
	;Récupération de valeur de $INI_SETTING (utile en cas de MAJ)
	Local $IniR_XPos = IniRead($INI_SETTING, "Paramètres", "XPos", -1)
	Local $IniR_YPos = IniRead($INI_SETTING, "Paramètres", "YPos", -1)
	Local $IniR_Flag = IniRead($INI_SETTING, "Paramètres", "Langue", "FR")

	;Suppression de donnée
	DirRemove($DIR_INSTALL, 1)
	Sleep(500)

	;Création de dossier
	DirCreate($DIR_INSTALL)
	DirCreate($DIR_DATA)
	DirCreate($DIR_IMAGE)
	DirCreate($DIR_DICO)

	;Installation de fichier
	FileInstall(".\Data\Image\Icon.ico", $DIR_IMAGE & "Icon.ico")
	FileInstall(".\Data\Image\Error.ico", $DIR_IMAGE & "Error.ico")
	FileInstall(".\Data\Image\FR.ico", $DIR_IMAGE & "FR.ico")
	FileInstall(".\Data\Image\EN.ico", $DIR_IMAGE & "EN.ico")
	FileInstall(".\Data\Image\DE.ico", $DIR_IMAGE & "DE.ico")
	FileInstall(".\Data\Image\IT.ico", $DIR_IMAGE & "IT.ico")
	FileInstall(".\Data\Image\ES.ico", $DIR_IMAGE & "ES.ico")
	FileInstall(".\Data\Dico\FR.ini", $DIR_DICO & "FR.ini")
	FileInstall(".\Data\Dico\EN.ini", $DIR_DICO & "EN.ini")
	FileInstall(".\Data\Dico\DE.ini", $DIR_DICO & "DE.ini")
	FileInstall(".\Data\Dico\IT.ini", $DIR_DICO & "IT.ini")
	FileInstall(".\Data\Dico\ES.ini", $DIR_DICO & "ES.ini")

	;Restauration de valeur de $INI_SETTING (utile en cas de MAJ)
	IniWrite($INI_SETTING, "Paramètres", "Version", $VERSION)
	IniWrite($INI_SETTING, "Paramètres", "XPos", $IniR_XPos)
	IniWrite($INI_SETTING, "Paramètres", "YPos", $IniR_YPos)
	IniWrite($INI_SETTING, "Paramètres", "Langue", $IniR_Flag)

	;Langue - FR
	IniWrite($INI_SETTING, "FR", "LetterNBR", "Nombre de lettres:")
	IniWrite($INI_SETTING, "FR", "LetterFIRST", "Première lettre:")
	IniWrite($INI_SETTING, "FR", "LetterEXP", "Contient l'expression:")
	IniWrite($INI_SETTING, "FR", "LetterCONTAINS", "Peut contenir les lettres:")
	IniWrite($INI_SETTING, "FR", "ColumnHeaderText", "Résultat:")
	IniWrite($INI_SETTING, "FR", "ChecboxAccent", "Présence d'accent inconnu")
	IniWrite($INI_SETTING, "FR", "ChecboxAccentTips", "Augmente le temps de la recherche.")
	IniWrite($INI_SETTING, "FR", "ButtonCancel", "Annuler")
	IniWrite($INI_SETTING, "FR", "MenuCopy", "Copier")
	IniWrite($INI_SETTING, "FR", "MenuAll", "Tout sélectionner")
	IniWrite($INI_SETTING, "FR", "TitreErreur", "Caractère non autorisé")
	IniWrite($INI_SETTING, "FR", "LetterNBRErreur", "Vous ne pouvez entrer qu'un nombre ici.")
	IniWrite($INI_SETTING, "FR", "LetterFIRSTErreur", "Vous ne pouvez entrer qu'une lettre ici.")
	IniWrite($INI_SETTING, "FR", "LetterEXPErreur", "Vous ne pouvez entrer qu'une expression ici. (Minimum 2 lettres)")
	IniWrite($INI_SETTING, "FR", "LetterCONTAINSErreur", "Vous ne pouvez entrer que des lettres ici. (Minimum 2 lettres)")

	;Langue - EN
	IniWrite($INI_SETTING, "EN", "LetterNBR", "Number of letter:")
	IniWrite($INI_SETTING, "EN", "LetterFIRST", "First letter:")
	IniWrite($INI_SETTING, "EN", "LetterEXP", "Contain the expression:")
	IniWrite($INI_SETTING, "EN", "LetterCONTAINS", "Can contain these letters:")
	IniWrite($INI_SETTING, "EN", "ColumnHeaderText", "Result:")
	IniWrite($INI_SETTING, "EN", "ChecboxAccent", "Unknown accent presence")
	IniWrite($INI_SETTING, "EN", "ChecboxAccentTips", "Increases the time of the search.")
	IniWrite($INI_SETTING, "EN", "ButtonCancel", "Cancel")
	IniWrite($INI_SETTING, "EN", "MenuCopy", "Copy")
	IniWrite($INI_SETTING, "EN", "MenuAll", "Select all")
	IniWrite($INI_SETTING, "EN", "TitreErreur", "Character not allowed")
	IniWrite($INI_SETTING, "EN", "LetterNBRErreur", "You can only enter a number here.")
	IniWrite($INI_SETTING, "EN", "LetterFIRSTErreur", "You can enter only a letter here.")
	IniWrite($INI_SETTING, "EN", "LetterEXPErreur", "You can only enter an expression here. (Two letters minimum)")
	IniWrite($INI_SETTING, "EN", "LetterCONTAINSErreur", "You can enter only letters here. (Two letters minimum)")

	;Langue - DE
	IniWrite($INI_SETTING, "DE", "LetterNBR", "Anzahl der buchstaben:")
	IniWrite($INI_SETTING, "DE", "LetterFIRST", "Anfangsbuchstaben:")
	IniWrite($INI_SETTING, "DE", "LetterEXP", "Enthält ausdrücke:")
	IniWrite($INI_SETTING, "DE", "LetterCONTAINS", "Kann buchstaben enthalten:")
	IniWrite($INI_SETTING, "DE", "ColumnHeaderText", "Ergebnis:")
	IniWrite($INI_SETTING, "DE", "ChecboxAccent", "Unknown akzent presence")
	IniWrite($INI_SETTING, "DE", "ChecboxAccentTips", "Erhöht die zeit der suche.")
	IniWrite($INI_SETTING, "DE", "ButtonCancel", "Stornieren")
	IniWrite($INI_SETTING, "DE", "MenuCopy", "Kopie")
	IniWrite($INI_SETTING, "DE", "MenuAll", "Alle auswählen")
	IniWrite($INI_SETTING, "DE", "TitreErreur", "Unzulässiges zeichen")
	IniWrite($INI_SETTING, "DE", "LetterNBRErreur", "Sie können nur geben sie eine zahl hier.")
	IniWrite($INI_SETTING, "DE", "LetterFIRSTErreur", "Sie können einen buschtaben hier eingeben.")
	IniWrite($INI_SETTING, "DE", "LetterEXPErreur", "Sie können nicht einen ausdruck hier eingeben. (Minimum 2 Buchstaben)")
	IniWrite($INI_SETTING, "DE", "LetterCONTAINSErreur", "Sie können nur buchstaben eingeben. (Minimum 2 Buchstaben)")

	;Langue - IT
	IniWrite($INI_SETTING, "IT", "LetterNBR", "Numero di lettera:")
	IniWrite($INI_SETTING, "IT", "LetterFIRST", "Prima lettera:")
	IniWrite($INI_SETTING, "IT", "LetterEXP", "Contiene espressioni:")
	IniWrite($INI_SETTING, "IT", "LetterCONTAINS", "Può contenere lettere:")
	IniWrite($INI_SETTING, "IT", "ColumnHeaderText", "Risultato:")
	IniWrite($INI_SETTING, "IT", "ChecboxAccent", "Unknown presenza accento")
	IniWrite($INI_SETTING, "IT", "ChecboxAccentTips", "Aumenta il tempo della ricerca.")
	IniWrite($INI_SETTING, "IT", "ButtonCancel", "Cancellare")
	IniWrite($INI_SETTING, "IT", "MenuCopy", "Copia")
	IniWrite($INI_SETTING, "IT", "MenuAll", "Seleziona tutto")
	IniWrite($INI_SETTING, "IT", "TitreErreur", "Carattere non valido")
	IniWrite($INI_SETTING, "IT", "LetterNBRErreur", "È possibile inserire solo un numero qui.")
	IniWrite($INI_SETTING, "IT", "LetterFIRSTErreur", "È possibile immettere una lettera qui.")
	IniWrite($INI_SETTING, "IT", "LetterEXPErreur", "Non è possibile immettere un'espressione qui. (Minimo 2 lettere)")
	IniWrite($INI_SETTING, "IT", "LetterCONTAINSErreur", "È possibile inserire solo lettere qui. (Minimo 2 lettere)")

	;Langue - ES
	IniWrite($INI_SETTING, "ES", "LetterNBR", "Número de letra:")
	IniWrite($INI_SETTING, "ES", "LetterFIRST", "Primera letra:")
	IniWrite($INI_SETTING, "ES", "LetterEXP", "Contiene expresiones:")
	IniWrite($INI_SETTING, "ES", "LetterCONTAINS", "Puede contener letras:")
	IniWrite($INI_SETTING, "ES", "ColumnHeaderText", "Resultado:")
	IniWrite($INI_SETTING, "ES", "ChecboxAccent", "Presencia acento desconocido")
	IniWrite($INI_SETTING, "ES", "ChecboxAccentTips", "Aumenta el tiempo de la búsqueda.")
	IniWrite($INI_SETTING, "ES", "ButtonCancel", "Cancelar")
	IniWrite($INI_SETTING, "ES", "MenuCopy", "Copia")
	IniWrite($INI_SETTING, "ES", "MenuAll", "Seleccionar todo")
	IniWrite($INI_SETTING, "ES", "TitreErreur", "Carácter ilegal")
	IniWrite($INI_SETTING, "ES", "LetterNBRErreur", "Sólo se puede introducir un número aquí.")
	IniWrite($INI_SETTING, "ES", "LetterFIRSTErreur", "Puede introducir aquí una carta.")
	IniWrite($INI_SETTING, "ES", "LetterEXPErreur", "No se puede introducir una expresión aquí. (Mínimo 2 letras)")
	IniWrite($INI_SETTING, "ES", "LetterCONTAINSErreur", "Puede introducir sólo letras aquí. (Mínimo 2 letras)")
EndFunc

Func _WMCommand($hWnd, $imsg, $iwParam, $ilParam)
	If BitAND($iwParam, 0x0000FFFF) = $ButtonCancel Then
		$bfInterrupt = True
		Return $GUI_RUNDEFMSG
	EndIf

	Local $nNotifyCode = BitShift($iwParam, 16)
	If $nNotifyCode = $EN_CHANGE Then

		Local $hCtrl = $ilParam, $Ctrlid = _WinAPI_GetDlgCtrlID($ilParam)
		Local $bAccentDetect = False, $bInputEXPErrorRegister = False
		Local $Input, $InputLessAccent

		;Détection d'accent dans les saisies
		For $i = 0 To UBound($aGCtrl)-1
			$Input = GUICtrlRead($aGCtrl[$i][1])
			$InputLessAccent = _ReplaceAlphaAccents($Input)
			If $InputLessAccent <> $Input Then
				$bAccentDetect = True
				ExitLoop
			EndIf
		Next
		If $bAccentDetect Then
			ControlDisable($GUIP, "", $ChecboxAccent)
		ElseIf Not ControlCommand($GUIP, "", $ChecboxAccent, "IsEnabled") Then
			ControlEnable($GUIP, "", $ChecboxAccent)
		EndIf

		;Règle de saisie
		$Input = ControlGetText($GUIP, "", $hCtrl)
		Switch $Ctrlid
			Case $aGCtrl[0][1] ;Nombre de lettres
				If Not $Input = "" And Not StringIsDigit($Input) Then
					GUICtrlSetState($aGCtrl[0][2], $GUI_SHOW)
				ElseIf ControlCommand($GUIP, "", $aGCtrl[0][2], "IsVisible") Then
					GUICtrlSetState($aGCtrl[0][2], $GUI_HIDE)
				EndIf

			Case $aGCtrl[1][1] ;Première lettre
				If Not $Input = "" And Not StringIsAlpha($Input) Or StringLen($Input) > 1 Then
					GUICtrlSetState($aGCtrl[1][2], $GUI_SHOW)
				ElseIf ControlCommand($GUIP, "", $aGCtrl[1][2], "IsVisible") Then
					GUICtrlSetState($aGCtrl[1][2], $GUI_HIDE)
				EndIf

			Case $aGCtrl[2][1] ;Contient l'expression
				If Not $Input = "" And Not StringIsAlpha($Input) Then
					GUICtrlSetState($aGCtrl[2][2], $GUI_SHOW)
				ElseIf Not $Input = "" And StringLen($Input) < 2 Then
					AdlibRegister("_InputEXPError", 1500)
					$bInputEXPErrorRegister = True
				ElseIf ControlCommand($GUIP, "", $aGCtrl[2][2], "IsVisible") Then
					GUICtrlSetState($aGCtrl[2][2], $GUI_HIDE)
				EndIf
				If Not $bInputEXPErrorRegister Then AdlibUnRegister("_InputEXPError")

			Case $aGCtrl[3][1] ;Peut contenir les lettres
				If Not $Input = "" And Not StringIsAlpha($Input) Then
					GUICtrlSetState($aGCtrl[3][2], $GUI_SHOW)
				ElseIf ControlCommand($GUIP, "", $aGCtrl[3][2], "IsVisible") Then
					GUICtrlSetState($aGCtrl[3][2], $GUI_HIDE)
				EndIf
		EndSwitch
	EndIf

	Return $GUI_RUNDEFMSG
EndFunc

Func _WMNotify($hWndGUI, $MsgID, $wParam, $lParam)
	#forceref $hWndGUI, $MsgID, $wParam, $lParam

	Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
    Local $hWndFrom = HWnd(DllStructGetData($tNMHDR, 1))
    Local $iCode = DllStructGetData($tNMHDR, 3)
	Local Const $tagNMLVKEYDOWN = $tagNMHDR & ";word VKey;uint Flags" ;Pour le fonctionnement sous 64 bits

	Select
		Case $wParam = $hListView
			Select
				Case $iCode = $LVN_BEGINDRAG
					Return 1 ; Disable Drag action

				Case $iCode = $NM_RCLICK

			EndSelect
	EndSelect

	Switch $hWndFrom
		Case $hHeaderG ;Column Listview ---------------------------------------------------------
			Switch $iCode
				Case $HDN_BEGINTRACK, $HDN_BEGINTRACKW
					Return True ;Disable resizing
			EndSwitch

		Case $hListView ;Listview ---------------------------------------------------------------
			Switch $iCode
				Case $LVN_BEGINDRAG
					Return 1 ;Disable Drag action

				Case $LVN_COLUMNCLICK
					Local $iFormat = _GUICtrlHeader_GetItemFormat($hHeaderG, 0)
					Local $iNewFormat = BitXOR($iFormat, $HDF_SORTUP)
					_GUICtrlListView_SimpleSort($hListView, $g_bSortSense, 0)
					_GUICtrlHeader_SetItemFormat($hHeaderG, 0, BitOR($iNewFormat, $HDF_SORTDOWN))

				Case $NM_DBLCLK
					Local $aSelectedIndices = _GUICtrlListView_GetSelectedIndices($hListView, True)
					If $aSelectedIndices[0] > 0 Then
						ClipPut(_GUICtrlListView_GetItemText($hListView, $aSelectedIndices[1]))
						Return 1
					EndIf

				Case $NM_RCLICK
					$bLVContextMenu = True
					Return 1

			EndSwitch
	EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc

Func _WMGuiMoveSave($hWnd, $nMsg, $wParam, $lParam)
	#forceref $hWnd, $nMsg, $wParam, $lParam

	;Enregistre l'emplacement de la $GUIP
	If $hWnd <> $GUIP Or StringRegExp($lParam, '(83008300)') Then Return $GUI_RUNDEFMSG ;83008300 correspond à "l'emplacement" quand on réduit la GUI (donc emplacement à ne pas enregistrer).
    Local $aWPos = WinGetPos($hWnd)
    If IsArray($aWPos) Then
		IniWrite($INI_SETTING, "Paramètres", "XPos", $aWPos[0])
		IniWrite($INI_SETTING, "Paramètres", "YPos", $aWPos[1])
        $Ini_XPos = $aWPos[0]
        $Ini_XPos = $aWPos[1]
    EndIf
	Return $GUI_RUNDEFMSG
EndFunc

Func _GuiClose()
	Exit
EndFunc

Func _Quit()
	If Not $bPIDExit Then IniWrite($INI_SETTING, "Paramètres", "PID", "null")
	_NoFocusLines_Clear()
EndFunc

;Amélioration possible:
;gui style/exstyle less flinkering LV/BUTTON ?
;LV police size un peu plus grosse ?
;replacer Caractère non autorisé par saisie non autorisé
;refaire la trad avec deepL
