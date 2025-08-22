#include-once
#include <StaticConstants.au3>
#include <GUIConstantsEx.au3>
#include "GuiCtrlOnHover.au3"
#include "NoFocusLines.au3"
#include <GDIPlus.au3>
#include "Icons.au3"
#include <Array.au3>

Global $__3SP_Array
__3StatPic_InitMainArray()
#cs
	$__3SP_Array[$i][0] = Handle du control (GuiCtrlPic)
	$__3SP_Array[$i][1] = Chemin de l'image *normale*
	$__3SP_Array[$i][2] = Chemin de l'image *survol*
	$__3SP_Array[$i][3] = Chemin de l'image *clique*
#ce

Func _3StatPic_Create($sImg1, $iImgX, $iImgY, $sImg2 = "", $sImg3 = "", $iImgW = "", $iImgH = "", $sTips = "", $iStyle = Default, $iExtStyle = Default, $sLabel = "", $iLabelSize = 8.5, $iLabelX = "", $iLabelY = "", $iLabelW = "", $iLabelH = "")

	;Paramètre par défaut
	Local $hCtrlLabel = ""
	If $sImg2 = "" Then $sImg2 = $sImg1
	If $sImg3 = "" Then $sImg3 = $sImg1

	;Charge l'image 1, récupère ça taille si nécéssaire, libère les ressources
	_GDIPlus_Startup()
	Local $Fimg = _GDIPlus_BitmapCreateFromFile($sImg1)
	If $iImgW = "" Then $iImgW = _GDIPlus_ImageGetWidth($Fimg)
	If $iImgH = "" Then $iImgH = _GDIPlus_ImageGetHeight($Fimg)
	_GDIPlus_ImageDispose($Fimg)
	_GDIPlus_Shutdown()

	;Création du contrôle de l'image + Tips si existant (bulle infos)
	Local $CtrlPic = GUICtrlCreatePic("", $iImgX, $iImgY, $iImgW, $iImgH, $iStyle, $iExtStyle)
	If $sTips <> "" Then GUICtrlSetTip(-1, $sTips)

	;Création du label OnTop
	If $sLabel <> "" Then
		If $iLabelX = "" Then $iLabelX = $iImgX
		If $iLabelY = "" Then $iLabelY = $iImgY
		If $iLabelW = "" Then $iLabelW = $iImgW
		If $iLabelH = "" Then $iLabelH = $iImgH
		Local $hCtrlLabel = GUICtrlCreateLabel($sLabel, $iLabelX, $iLabelY, $iLabelW, $iLabelH, BitOR($GUI_SS_DEFAULT_LABEL, $SS_CENTER, $SS_CENTERIMAGE))
		GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
		GUICtrlSetFont(-1, $iLabelSize * _RatioFont()[0])
	EndIf

	If $sImg2 <> $sImg1 Then
		_GUICtrl_OnHoverRegister($CtrlPic, "__3StatPic_OnHoverFunction", "__3StatPic_OnHoverFunction", _
									"__3StatPic_OnClickFunction", "__3StatPic_OnClickFunction", 0)
	ElseIf $sImg3 <> $sImg1 Then
		_GUICtrl_OnHoverRegister($CtrlPic, "", "", _
									"__3StatPic_OnClickFunction", "__3StatPic_OnClickFunction", 0)
	EndIf

	;Affecte l'image au contrôle
	_SetImage($CtrlPic, $sImg1)

	__3StatPic_ArrayAdd($CtrlPic, $sImg1, $sImg2, $sImg3)

	Return $CtrlPic
EndFunc

Func _3StatPic_ChangeImg($hCtrl, $sImg1 = Default, $sImg2 = Default, $sImg3 = Default)
	Local $iCtrlIndex = __3StatPic_hCtrlToIndex($hCtrl)
	If $iCtrlIndex = -1 Then Return 0
	If $sImg1 <> Default Then $__3SP_Array[$iCtrlIndex][1] = $sImg1
	If $sImg2 <> Default Then $__3SP_Array[$iCtrlIndex][2] = $sImg2
	If $sImg3 <> Default Then $__3SP_Array[$iCtrlIndex][3] = $sImg3
	_SetImage($hCtrl, $sImg1)
	Return 1
EndFunc

Func _3StatPic_Delete($hCtrl)
	__3StatPic_ArrayDel($hCtrl)
	Return GuiCtrlDelete($hCtrl)
EndFunc

Func _RatioFont($iDPIDef = 96) ;Gère la taille de l'écriture en fonction du DPI windows utilisé. (_GDIPlus_GraphicsGetDPIRatio)
    #forcedef $__g_hGDIPDll, $ghGDIPDll
	_GDIPlus_Startup()
    Local $hGfx = _GDIPlus_GraphicsCreateFromHWND(0)
    If @error Then Return SetError(1, @extended, 0)
    Local $aResult = DllCall($__g_hGDIPDll, "int", "GdipGetDpiX", "handle", $hGfx, "float*", 0)
    If @error Then Return SetError(2, @error, 0)
	Local $iDPI = $aResult[2]
	Local $aResults[2] = [$iDPIDef / $iDPI, $iDPI / $iDPIDef]
	_GDIPlus_GraphicsDispose($hGfx)
    _GDIPlus_Shutdown()
    Return $aResults
EndFunc

Func _DisableTabFocus($hCtrl) ;Désactive le focus du controle via la touche tabulation (désactive $WS_TABSTOP)
	_WinAPI_SetWindowLong($hCtrl, $GWL_STYLE, BitAND(_WinAPI_GetWindowLong($hCtrl, $GWL_STYLE), BitNOT($WS_TABSTOP)))
EndFunc

; ###############################################################
; ################## Internal use only ##########################
; ###############################################################

Func __3StatPic_OnClickFunction($hCtrl, $iClickMode)
	Local $iCtrlIndex = __3StatPic_hCtrlToIndex($hCtrl)
	If $iCtrlIndex = -1 Then Return 0

	Switch $iClickMode
		Case 1 ; pressed
			_SetImage($hCtrl, $__3SP_Array[$iCtrlIndex][3])
		Case 2 ; released
			_SetImage($hCtrl, $__3SP_Array[$iCtrlIndex][2])
	EndSwitch
EndFunc

Func __3StatPic_OnHoverFunction($hCtrl, $iHoverMode, $hWin)
	Local $iCtrlIndex = __3StatPic_hCtrlToIndex($hCtrl)
	If $iCtrlIndex = -1 Then Return 0

	Switch $iHoverMode
		Case 1 ; hover
			_SetImage($hCtrl, $__3SP_Array[$iCtrlIndex][2])
		Case 2 ; leave
			_SetImage($hCtrl, $__3SP_Array[$iCtrlIndex][1])
	EndSwitch
EndFunc

Func __3StatPic_ArrayAdd($hCtrl, $sImg1, $sImg2, $sImg3)
	Local $iNbr = $__3SP_Array[0][0]
	Redim $__3SP_Array[$iNbr + 2][4]
	$__3SP_Array[$iNbr + 1][0] = $hCtrl
	$__3SP_Array[$iNbr + 1][1] = $sImg1
	$__3SP_Array[$iNbr + 1][2] = $sImg2
	$__3SP_Array[$iNbr + 1][3] = $sImg3
	$__3SP_Array[0][0] = $iNbr + 1
EndFunc

Func __3StatPic_ArrayDel($hCtrl)
	For $i = 1 To $__3SP_Array[0][0]
		If $__3SP_Array[$i][0] = $hCtrl Then
			_ArrayDelete($__3SP_Array, $i)
			If Not IsArray($__3SP_Array) Then __3StatPic_InitMainArray()
			Return 1
		EndIf
	Next
	Return 0
EndFunc

Func __3StatPic_hCtrlToIndex($hCtrl)
	Local $iCtrlIndex = -1
	For $i = 1 To $__3SP_Array[0][0]
		If $__3SP_Array[$i][0] = $hCtrl Then
			$iCtrlIndex = $i
			ExitLoop
		EndIf
	Next
	Return $iCtrlIndex
EndFunc

Func __3StatPic_InitMainArray()
	Global $__3SP_Array[1][4]
	$__3SP_Array[0][0] = 0
EndFunc
