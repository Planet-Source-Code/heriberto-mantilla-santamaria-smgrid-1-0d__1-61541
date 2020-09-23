VERSION 5.00
Begin VB.UserControl SMGrid 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.HScrollBar HScrollUpD 
      Height          =   270
      Left            =   300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   825
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.VScrollBar VScrollUpD 
      Height          =   1170
      Left            =   300
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "SMGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2005       *'
'******************************************************'
'*                     Version 1.0a                   *'
'******************************************************'
'* Control:       SMGrid Control                      *'
'******************************************************'
'* Author:        Heriberto Mantilla Santamaría       *'
'******************************************************'
'* Description:   Emulate a FlexGrid Control          *'
'******************************************************'
'* Started on:    Wednesday, 10-nov-2004.             *'
'* Release date:  Friday, 18-mar-2005.                *'
'******************************************************'
'* Comments: This control was necessary to develop it *'
'*           for a program of a thesis of grade of my *'
'*           University, its evolution was stopped by *'
'*           a lot of time, although it is not comple-*'
'*           tely ended, but it's a beginning.        *'
'******************************************************'
'*----------------------------------------------------*'
'* Credits:  Richard Mewett (GetColFromX Function)    *'
'*           [CodeId = 61438]                         *'
'*----------------------------------------------------*'
'*
'******************************************************'
'*                                                    *'
'* Note:     Comments, suggestions, doubts or bug     *'
'*           reports are wellcome to these e-mail     *'
'*           addresses:                               *'
'*                                                    *'
'*                  heri_05-hms@mixmail.com or        *'
'*                  hcammus@hotmail.com               *'
'******************************************************'
'* Now my website is available but alone the version  *'
'* in Spanish.                                        *'
'-----------------------------------------------------*'
'* WebSite:  http://hackprotm.webcindario.com/        *'
'*           http://www.geocities.com/hackprotm/      *'
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2005       *'
'******************************************************'
Option Explicit
 
 Public Enum EnumValue
  ComboBox = &H0
  TextBox = &H1
  Button = &H2
  None = &H3
 End Enum
 
 Private Type Tamano
  Height As Long
  Width  As Long
 End Type
 
 Private Type isStyleI
  Colors()     As OLE_COLOR
  Item()       As String
  Style()      As String
  TextColors() As OLE_COLOR
  TotalI       As Integer
  Values()     As Boolean
 End Type
 
 Private Type isStyle
  Col          As Long
  Item         As isStyleI
  Row          As Long
  TextList     As String
 End Type
 
 Private Type POINTAPI
  X           As Long
  Y           As Long
 End Type
 
 Private Type RECT
  Left         As Long
  Top          As Long
  Right        As Long
  Bottom       As Long
 End Type
   
 Private Const BF_BOTTOM = &H8
 Private Const BF_LEFT = &H1
 Private Const BF_RIGHT = &H4
 Private Const BF_TOP = &H2
 Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
 Private Const COLOR_BTNFACE = 15
 Private Const COLOR_WINDOW = 5
 Private Const DFC_BUTTON = 4
 Private Const DFCS_BUTTONCHECK = &H0
 Private Const DFCS_BUTTONRADIO = &H4
 Private Const DFCS_BUTTON3STATE = &H10
 Private Const DFCS_CHECKED As Long = &H400
 Private Const BDR_RAISEDINNER = &H4
 Private Const BDR_RAISEDOUTER = &H1
 Private Const BDR_SUNKENINNER = &H8
 Private Const BDR_SUNKENOUTER = &H2
 Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
 'Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
 Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
 Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
 Private Const DT_BOTTOM        As Long = &H8
 Private Const DT_CALCRECT      As Long = &H400
 Private Const DT_CENTER        As Long = &H1
 Private Const DT_RIGHT         As Long = &H2
 Private Const DT_TOP           As Long = &H0
 Private Const DT_VCENTER       As Long = &H4
 Private Const DT_WORD_ELLIPSIS As Long = &H40000
 Private Const DT_WORDBREAK     As Long = &H10
 
 Private bDrawTheme      As Boolean
 Private ColItem         As Integer
 Private Columns         As Long
 Private Headers()       As Tamano
 Private hTheme          As Long '* hTheme Handle.
 Private isEnabled       As Boolean
 Private Items()         As isStyle
 Private ItemsT()        As isStyleI
 Private m_lBackColor    As OLE_COLOR
 Private m_lHeadersColor As OLE_COLOR
 Private m_StateG        As Integer
 Private m_sTextHeaders  As String
 Private m_txtRect       As RECT
 Private m_Buttons       As RECT
 Private NoChanged       As Boolean
 Private RowItem         As Integer
 Private ShowItems       As Integer
 Private tmpC1           As Integer
 Private tmpC2           As Integer
 Private TheForm         As Object
 Private TheText         As Object
 Private TotalItems      As Long
 
 Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
 Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
 Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
 Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
 Private Declare Function DrawThemeEdge Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pDestRect As RECT, ByVal uEdge As Long, ByVal uFlags As Long, pContentRect As RECT) As Long
 Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
 Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
 Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
 Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
 
 Public Event Click(ByVal Row As Long, ByVal Col As Long, ByVal Style As EnumValue)

'---------------------------------------------------------------------------------------
' Procedure : BackColor
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Set the BackColor.
'---------------------------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR
 BackColor = m_lBackColor
End Property

'---------------------------------------------------------------------------------------
' Procedure : BackColor
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Set the BackColor.
'---------------------------------------------------------------------------------------
Public Property Let BackColor(ByVal lBackColor As OLE_COLOR)
 m_lBackColor = ConvertSystemColor(lBackColor)
 Call UserControl.PropertyChanged("BackColor")
 Call Refresh
End Property

Public Property Get ColPos()
 ColPos = ColItem
End Property

'---------------------------------------------------------------------------------------
' Procedure : HeadersColor
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : BackColor of the Headers.
'---------------------------------------------------------------------------------------
Public Property Get HeadersColor() As OLE_COLOR
 HeadersColor = m_lHeadersColor
End Property

'---------------------------------------------------------------------------------------
' Procedure : HeadersColor
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : BackColor of the Headers.
'---------------------------------------------------------------------------------------
Public Property Let HeadersColor(ByVal lHeadersColor As OLE_COLOR)
 m_lHeadersColor = ConvertSystemColor(lHeadersColor)
 Call UserControl.PropertyChanged("HeadersColor")
 Call Refresh
End Property
 
'---------------------------------------------------------------------------------------
' Procedure : Enabled
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Enabled/Disabled control.
'---------------------------------------------------------------------------------------
Public Property Get Enabled() As Boolean
 Enabled = isEnabled
End Property

'---------------------------------------------------------------------------------------
' Procedure : Enabled
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Enabled/Disabled control.
'---------------------------------------------------------------------------------------
Public Property Let Enabled(ByVal isValue As Boolean)
 isEnabled = isValue
 UserControl.Enabled = isEnabled
 VScrollUpD.Enabled = isEnabled
 Call Refresh
 Call PropertyChanged("Enabled")
End Property

'---------------------------------------------------------------------------------------
' Procedure : ListCount
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Total items of grid.
'---------------------------------------------------------------------------------------
Public Property Get ListCount() As Long
Attribute ListCount.VB_MemberFlags = "400"
 ListCount = TotalItems
End Property

Public Property Get RowPos()
 RowPos = RowItem
End Property

'---------------------------------------------------------------------------------------
' Procedure : TextHeaders
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Text of the Headers of grid.
'---------------------------------------------------------------------------------------
Public Property Get TextHeaders() As String
 TextHeaders = m_sTextHeaders
End Property

'---------------------------------------------------------------------------------------
' Procedure : TextHeaders
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Text of the Headers of grid.
'---------------------------------------------------------------------------------------
Public Property Let TextHeaders(ByVal sTextHeaders As String)
 m_sTextHeaders = sTextHeaders
 Call UserControl.PropertyChanged("TextHeaders")
 Call Refresh
End Property

'---------------------------------------------------------------------------------------
' Procedure : AddItem
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Add a new item.
'---------------------------------------------------------------------------------------
Public Sub AddItem(ByVal Item As String, Optional ByVal Colors As String = "", Optional ByVal TextColors As String = "", Optional ByVal Style As String = "", Optional ByVal TextList As String = "", Optional ByVal Values As String)
 Dim iCarac As Variant, i As Integer
 
 TotalItems = TotalItems + 1
 ReDim Preserve Items(TotalItems)
 iCarac = Split(m_sTextHeaders, "|")
 Items(TotalItems).Col = UBound(iCarac) + 1
 iCarac = Split(Item, "|")
 If (UBound(iCarac) >= 0) Then
  ReDim Preserve Items(TotalItems).Item.Item(UBound(iCarac))
  ReDim Preserve Items(TotalItems).Item.Colors(UBound(iCarac))
  ReDim Preserve Items(TotalItems).Item.TextColors(UBound(iCarac))
  ReDim Preserve Items(TotalItems).Item.Style(Items(TotalItems).Col)
  ReDim Preserve Items(TotalItems).Item.Values(Items(TotalItems).Col)
 End If
 For i = 0 To UBound(iCarac)
  Items(TotalItems).Item.Item(i) = iCarac(i)
  Items(TotalItems).Item.Values(i) = False
  Items(TotalItems).Item.Colors(i) = ConvertSystemColor(UserControl.BackColor)
  Items(TotalItems).Item.Style(i) = ""
  Items(TotalItems).Item.TextColors(i) = &H0
 Next
 iCarac = Split(Colors, "|")
 For i = 0 To UBound(iCarac)
  ReDim Preserve Items(TotalItems).Item.Colors(i)
  Items(TotalItems).Item.Colors(i) = IIf(iCarac(i) = "", ConvertSystemColor(UserControl.BackColor), iCarac(i))
 Next
 iCarac = Split(TextColors, "|")
 For i = 0 To UBound(iCarac)
  ReDim Preserve Items(TotalItems).Item.TextColors(i)
  Items(TotalItems).Item.TextColors(i) = IIf(iCarac(i) = "", &H0, iCarac(i))
 Next
 iCarac = Split(Style, "|")
 For i = 0 To UBound(iCarac)
  ReDim Preserve Items(TotalItems).Item.Style(i)
  Items(TotalItems).Item.Style(i) = iCarac(i)
 Next
 iCarac = Split(Values, "|")
 For i = 0 To UBound(iCarac)
  ReDim Preserve Items(TotalItems).Item.Values(i)
  If (iCarac(i) <> "T") Or (iCarac(i) <> "F") Or (iCarac(i) = "F") Then
   iCarac = False
  ElseIf (iCarac(i) = "T") Then
   iCarac = True
  End If
  Items(TotalItems).Item.Values(i) = iCarac(i)
 Next
 Items(TotalItems).Row = TotalItems
 Items(TotalItems).TextList = TextList
 iCarac = ""
End Sub

'---------------------------------------------------------------------------------------
' Procedure : APIFillRect
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Pinta el rectángulo de un objeto.
'---------------------------------------------------------------------------------------
Private Sub APIFillRect(ByVal hDC As Long, ByRef RC As RECT, ByVal Color As Long)
 Dim NewBrush As Long
 
 NewBrush& = CreateSolidBrush(Color&)
 Call FillRect(hDC&, RC, NewBrush&)
 Call DeleteObject(NewBrush&)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : APILine
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Pinta líneas de forma sencilla y rápida.
'---------------------------------------------------------------------------------------
Private Sub APILine(ByVal whDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal lColor As Long)
 Dim PT As POINTAPI, hPen As Long, hPenOld As Long
 
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(whDC, hPen)
 Call MoveToEx(whDC, x1, y1, PT)
 Call LineTo(whDC, x2, y2)
 Call SelectObject(whDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChangedItem
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Cambia el valor de un item.
'---------------------------------------------------------------------------------------
Public Sub ChangedItem(ByVal Col As Integer, ByVal Row As Integer, ByVal Item As String)
 Items(Col).Item.Item(Row) = Item
 Call Refresh
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ConvertSystemColor
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Convierte un Long en un color del sistema.
'---------------------------------------------------------------------------------------
Private Function ConvertSystemColor(ByVal theColor As Long) As Long
 Call OleTranslateColor(theColor, 0, ConvertSystemColor)
End Function

'---------------------------------------------------------------------------------------
' Procedure : DrawCaption
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Paint Text of the object.
'---------------------------------------------------------------------------------------
Private Sub DrawCaption(ByVal lCaption As String, Optional ByVal lColor As OLE_COLOR = &H0)
 Dim iAlign As String, RemF As Boolean, fAlign As Long
 
 iAlign = Mid$(lCaption, 1, 1)
 RemF = True
 fAlign = 0
 Select Case iAlign
  Case "^" '* Center Align.
   fAlign = DT_CENTER
  Case "~" '* Right Align.
   fAlign = DT_RIGHT
   m_txtRect.Right = m_txtRect.Right - 3
  Case Else
   RemF = False
   lCaption = " " & lCaption
 End Select
 If (RemF = True) Then lCaption = Mid$(Trim$(lCaption), 2, Len(lCaption))
 Call SetTextColor(hDC, lColor)
 Call DrawText(hDC, lCaption, Len(lCaption), m_txtRect, DT_TOP Or DT_VCENTER Or DT_WORD_ELLIPSIS Or fAlign)
 If (fAlign = DT_RIGHT) Then m_txtRect.Right = m_txtRect.Right + 3
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DrawHeaders
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Paint the text headers.
'---------------------------------------------------------------------------------------
Private Function DrawHeaders(Optional ByVal Value As Long = 0, Optional ByVal State As Integer = 0) As Long
 Dim m_Text As Variant, i  As Long
 Dim xRight As Long, lBack As OLE_COLOR
 
 m_Text = Split(m_sTextHeaders, "|")
 m_txtRect.Right = 0
 m_txtRect.Top = 2
 m_txtRect.Bottom = 22
 m_txtRect.Left = 2
 If (UBound(m_Text) > 0) Then ReDim Headers(UBound(m_Text))
 For i = Value To UBound(m_Text)
  m_txtRect.Right = m_txtRect.Right + UserControl.TextWidth(m_Text(i)) + 10
  ReDim Preserve Headers(i)
  Headers(i).Width = m_txtRect.Right
  Headers(i).Height = m_txtRect.Bottom
  lBack = ConvertSystemColor(UserControl.BackColor)
  Call APIFillRect(hDC, m_txtRect, m_lHeadersColor)
  bDrawTheme = DrawTheme("Header", 1, State, m_txtRect)
  If (bDrawTheme = False) Then
   Call DrawEdge(hDC, m_txtRect, EDGE_RAISED, BF_RECT)
  Else
   Call DrawThemeEdge(hTheme, hDC, 1, 0, m_txtRect, EDGE_RAISED, BF_RECT, m_txtRect)
  End If
  xRight = m_txtRect.Right
  m_txtRect.Top = 4
  Call DrawCaption(m_Text(i))
  m_txtRect.Left = xRight
  m_txtRect.Top = 2
 Next
 DrawHeaders = i
 Call CloseThemeData(hTheme)
End Function

'---------------------------------------------------------------------------------------
' Procedure : DrawPoint
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Pinta tres puntos.
'---------------------------------------------------------------------------------------
Private Sub DrawPoint(Optional ByVal iColor3 As OLE_COLOR = &H0)
 iColor3 = ConvertSystemColor(iColor3)
 tmpC1 = m_Buttons.Right - m_Buttons.Left - 3
 tmpC2 = m_Buttons.Bottom - m_Buttons.Top + 1
 tmpC1 = m_Buttons.Left + tmpC1 / 2 + 1
 tmpC2 = m_Buttons.Top + tmpC2 / 2 - 1
 Call APILine(UserControl.hDC, tmpC1 - 3, tmpC2, tmpC1 - 3, tmpC2 + 1, iColor3)
 Call APILine(UserControl.hDC, tmpC1, tmpC2, tmpC1, tmpC2 + 1, iColor3)
 Call APILine(UserControl.hDC, tmpC1 + 3, tmpC2, tmpC1 + 3, tmpC2 + 1, iColor3)
End Sub

'---------------------------------------------------------------------------------------
' Function  : DrawTheme
' DateTime  : 03/08/05 13:38
' Author    : HACKPRO TM
' Purpose   : Try to open Uxtheme.dll.
'---------------------------------------------------------------------------------------
Private Function DrawTheme(sClass As String, ByVal iPart As Long, ByVal iState As Long, rtRect As RECT, Optional ByVal CloseTheme As Boolean = False) As Boolean
 Dim lResult As Long '* Temp Variable.
 
 '* If a error occurs then or we are not running XP or the visual style is Windows Classic.
On Error GoTo NoXP
 '* Get out hTheme Handle.
 hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))
 '* Did we get a theme handle?.
 If (hTheme) Then
  '* Yes! Draw the control Background.
  lResult = DrawThemeBackground(hTheme, UserControl.hDC, iPart, iState, rtRect, rtRect)
  '* If drawing was successful, return true, or false If not.
  DrawTheme = IIf(lResult, False, True)
 Else
  '* No, we couldn't get a hTheme, drawing failed.
  DrawTheme = False
 End If
 '* Close theme.
 If (CloseTheme = True) Then Call CloseThemeData(hTheme)
 '* Exit the function now.
 Exit Function
NoXP:
 '* An Error was detected, drawing Failed.
 DrawTheme = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : DrawXpArrow
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Dibuja la flecha estilo Xp.
'---------------------------------------------------------------------------------------
Private Sub DrawXpArrow(Optional ByVal iColor3 As OLE_COLOR = &H0)
 Dim tmpC1 As Long, tmpC2 As Long
 
 iColor3 = ConvertSystemColor(iColor3)
 tmpC1 = m_Buttons.Right - m_Buttons.Left - 1
 tmpC2 = m_Buttons.Bottom - m_Buttons.Top + 1
 tmpC1 = m_Buttons.Left + tmpC1 / 2 + 1
 tmpC2 = m_Buttons.Top + tmpC2 / 2
 Call APILine(UserControl.hDC, tmpC1 - 5, tmpC2 - 2, tmpC1, tmpC2 + 3, iColor3)
 Call APILine(UserControl.hDC, tmpC1 - 4, tmpC2 - 2, tmpC1, tmpC2 + 2, iColor3)
 Call APILine(UserControl.hDC, tmpC1 - 4, tmpC2 - 3, tmpC1, tmpC2 + 1, iColor3)
 Call APILine(UserControl.hDC, tmpC1 + 3, tmpC2 - 2, tmpC1 - 2, tmpC2 + 3, iColor3)
 Call APILine(UserControl.hDC, tmpC1 + 2, tmpC2 - 2, tmpC1 - 2, tmpC2 + 2, iColor3)
 Call APILine(UserControl.hDC, tmpC1 + 2, tmpC2 - 3, tmpC1 - 2, tmpC2 + 1, iColor3)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FormShow
' DateTime  : 09/07/05 13:35
' Author    : HACKPRO TM
' Purpose   : Establece los controles para Button.
'---------------------------------------------------------------------------------------
Private Sub FormShow(ByRef isObject As Object, ByRef isText As Object)
On Error Resume Next
 isText.Text = ItemText(RowItem, ColItem)
 Call isObject.Show(1)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetColFromX
' DateTime  : 09/07/05 09:19
' Author    : Richard Mewett (Thanks)
' Purpose   : Position Col.
'---------------------------------------------------------------------------------------
Private Function GetColFromX(ByVal X As Single) As Integer
 Dim lX As Long, nCol As Integer
    
 GetColFromX = -1
 For nCol = LBound(Headers) To UBound(Headers)
  If (Headers(nCol).Width >= X) Then
   GetColFromX = nCol
   Exit For
  End If
 Next
End Function

Public Function ObjectForm(ByRef isObject As Object, ByRef isText As Object)
On Error Resume Next
 Set TheForm = isObject
 Set TheText = isText
End Function

'---------------------------------------------------------------------------------------
' Procedure : Wait
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Espera un x tiempo.
'---------------------------------------------------------------------------------------
Private Sub Wait(ByVal Segundos As Single)
 Dim ComienzoSeg As Single, sumSeg As Long
 Dim FinSeg      As Single
 
 ComienzoSeg = Timer
 FinSeg = ComienzoSeg + Segundos
 sumSeg = 0
 Do While (FinSeg > Timer)
  DoEvents
  If (ComienzoSeg > Timer) Then FinSeg = FinSeg - 24 * 60 * 60
  If (sumSeg > 20) Then Exit Do
  sumSeg = sumSeg + 1
 Loop
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ItemText
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Muestra el item actual.
'---------------------------------------------------------------------------------------
Public Function ItemText(ByVal Col As Integer, ByVal Row As Integer) As String
On Error Resume Next
 ItemText = Items(Row).Item.Item(Col)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Refresh
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Refresca el control.
'---------------------------------------------------------------------------------------
Public Sub Refresh(Optional ByVal UpDown As Boolean = False)
 Dim CountItems As Integer, i         As Long, tmpRect As RECT
 Dim IPos       As String, xRight     As Long, j       As Long
 Dim lColor     As OLE_COLOR, mColor  As Long, isOpt   As Integer
 
 UserControl.BackColor = GetSysColor(COLOR_WINDOW)
 HScrollUpD.Visible = False
 Call VScrollUpD.Move(UserControl.ScaleWidth - 19, 2, 17, UserControl.ScaleHeight - 22)
 Call HScrollUpD.Move(2, UserControl.ScaleHeight - 20, UserControl.ScaleWidth - 4, 18)
 VScrollUpD.SmallChange = 3
 VScrollUpD.LargeChange = 3
 ShowItems = Int(UserControl.ScaleHeight / 18)
 If (TotalItems > ShowItems) Then
  VScrollUpD.Max = TotalItems - ShowItems + 2
  VScrollUpD.Visible = True
 Else
  VScrollUpD.Max = TotalItems
  ShowItems = TotalItems
  VScrollUpD.Visible = False
 End If
 UserControl.Cls
 UserControl.BackColor = m_lBackColor
 Call SetRect(m_txtRect, 4, 4, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4)
 CountItems = DrawHeaders(HScrollUpD.Value)
On Error Resume Next
 If (TotalItems > 0) Then
  If ((CountItems * Headers(UBound(Headers)).Width) > UserControl.ScaleWidth) Then
   HScrollUpD.Visible = True
   HScrollUpD.Max = UserControl.ScaleWidth / Headers(UBound(Headers)).Width
  End If
 End If
 m_Buttons.Top = 25
 m_Buttons.Bottom = 40
 For i = HScrollUpD.Value + VScrollUpD.Value + 1 To ShowItems + VScrollUpD.Value
  m_txtRect.Left = 2
  m_txtRect.Right = 0
  Call OffsetRect(m_txtRect, 0, 20)
  For j = HScrollUpD.Value To CountItems - 1
  On Error Resume Next
   m_txtRect.Right = Headers(j).Width
   m_Buttons.Left = Headers(j).Width - 17
   m_Buttons.Right = Headers(j).Width - 2
   lColor = ConvertSystemColor(UserControl.BackColor)
   mColor = &H0
   If (j <= UBound(Items(i).Item.Colors)) And (UBound(Items(i).Item.Colors) > 0) Then lColor = ConvertSystemColor(Items(i).Item.Colors(j))
   If (j <= UBound(Items(i).Item.TextColors)) And (UBound(Items(i).Item.TextColors) > 0) Then mColor = ConvertSystemColor(Items(i).Item.TextColors(j))
   Call APIFillRect(UserControl.hDC, m_txtRect, lColor)
   Call DrawEdge(hDC, m_txtRect, EDGE_BUMP, BF_RECT)
   xRight = m_txtRect.Right
   m_txtRect.Top = m_txtRect.Top + 3
   Call DrawCaption(Items(i).Item.Item(j), mColor)
   IPos = ""
   If (j <= UBound(Items(i).Item.Style)) And (UBound(Items(i).Item.Style) > 0) Then IPos = Items(i).Item.Style(j)
   bDrawTheme = False
   isOpt = 0
   If (Items(i).Item.Values(j) = True) And (UpDown = True) Then isOpt = 3
   If (IPos = "B") Then '* Button.
    bDrawTheme = DrawTheme("Button", 1, isOpt, m_Buttons)
    If (bDrawTheme = False) Then
     Call APIFillRect(hDC, m_Buttons, GetSysColor(COLOR_BTNFACE))
     If (ColItem = j) And (UpDown = True) Then
      Call DrawEdge(hDC, m_Buttons, EDGE_SUNKEN, BF_RECT)
     Else
      Call DrawEdge(hDC, m_Buttons, EDGE_RAISED, BF_RECT)
     End If
    End If
    Call DrawPoint(IIf(isEnabled = False, &H80000011, &H80000012))
   ElseIf (IPos = "C") Then '* ComboBox.
    bDrawTheme = DrawTheme("Edit", 2, isOpt, m_Buttons)
    Call SetRect(tmpRect, m_Buttons.Left - 1, m_Buttons.Top - 1, m_Buttons.Right + 1, m_Buttons.Bottom + 1)
    bDrawTheme = DrawTheme("ComboBox", 1, 0, tmpRect)
    If (bDrawTheme = False) Then
     Call APIFillRect(hDC, m_Buttons, GetSysColor(COLOR_BTNFACE))
     If (ColItem = j) And (UpDown = True) Then
      Call DrawEdge(hDC, m_Buttons, EDGE_SUNKEN, BF_RECT)
     Else
      Call DrawEdge(hDC, m_Buttons, EDGE_RAISED, BF_RECT)
     End If
     Call DrawXpArrow(IIf(isEnabled = False, &H80000011, &H80000012))
    End If
   ElseIf (IPos = "T") Then '* TextBox.
      
   ElseIf (IPos = "Ch") Then '* CheckBox.
    isOpt = (Items(i).Item.Values(j) * -5)
    bDrawTheme = DrawTheme("Button", 3, isOpt, m_Buttons)
    If (bDrawTheme = False) Then
     m_Buttons.Top = m_Buttons.Top + 1
     m_Buttons.Bottom = m_Buttons.Bottom - 1
     m_Buttons.Left = m_Buttons.Left + 1
     If (ColItem = j) And (isOpt = 0) Then
      Call DrawFrameControl(hDC, m_Buttons, DFC_BUTTON, DFCS_BUTTONCHECK) 'Or DFCS_CHECKED)
     Else
      Call DrawFrameControl(hDC, m_Buttons, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED)
     End If
     m_Buttons.Bottom = m_Buttons.Bottom + 1
     m_Buttons.Top = m_Buttons.Top - 1
     m_Buttons.Left = m_Buttons.Left - 1
    End If
   ElseIf (IPos = "O") Then '* OptionBox.
    isOpt = IIf(Items(i).Item.Values(j) = True, 6, 1)
    bDrawTheme = DrawTheme("Button", 2, isOpt, m_Buttons)
    If (ColItem = j) And (bDrawTheme = False) Then
     m_Buttons.Top = m_Buttons.Top + 1
     m_Buttons.Bottom = m_Buttons.Bottom - 1
     m_Buttons.Left = m_Buttons.Left + 1
     Call DrawFrameControl(hDC, m_Buttons, DFC_BUTTON, DFCS_BUTTONRADIO) 'Or DFCS_CHECKED)
     m_Buttons.Bottom = m_Buttons.Bottom + 1
     m_Buttons.Top = m_Buttons.Top - 1
     m_Buttons.Left = m_Buttons.Left - 1
    End If
   ElseIf (IPos = "Bk") Then '* Background Picture.
    
   End If
   m_txtRect.Left = xRight
   m_txtRect.Top = m_txtRect.Top - 3
  Next
  m_Buttons.Top = m_Buttons.Top + 20
  m_Buttons.Bottom = m_Buttons.Top + 15
 Next
 m_txtRect.Left = 0
 m_txtRect.Top = 0
 m_txtRect.Right = UserControl.ScaleWidth
 m_txtRect.Bottom = UserControl.ScaleHeight
 Call DrawEdge(hDC, m_txtRect, EDGE_SUNKEN, BF_RECT)
 Columns = CountItems
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShiftColorOXP
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Shift a color.
'---------------------------------------------------------------------------------------
Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
 Dim cRed   As Long, cBlue  As Long
 Dim Delta  As Long, cGreen As Long

 cBlue = ((theColor \ &H10000) Mod &H100)
 cGreen = ((theColor \ &H100) Mod &H100)
 cRed = (theColor And &HFF)
 Delta = &HFF - Base
 cBlue = Base + cBlue * Delta \ &HFF
 cGreen = Base + cGreen * Delta \ &HFF
 cRed = Base + cRed * Delta \ &HFF
 If (cRed > 255) Then cRed = 255
 If (cGreen > 255) Then cGreen = 255
 If (cBlue > 255) Then cBlue = 255
 ShiftColorOXP = cRed + 256& * cGreen + 65536 * cBlue
End Function

Private Sub UserControl_InitProperties()
 TotalItems = 0
 m_sTextHeaders = ""
 ReDim Items(1)
 isEnabled = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim isStyle As Integer, iStyle As String
 Dim isCol   As Long, IPos      As Long
 
 RowItem = (GetColFromX(X) + HScrollUpD.Value) - IIf(HScrollUpD.Value > 0, 1, 0) '* Col.
 ColItem = ((Y / Headers(RowItem).Height) + VScrollUpD.Value)  '* Row.
 If (ColItem = 0) Then ColItem = 1
On Error Resume Next
 iStyle = Items(ColItem).Item.Style(RowItem)
 If (iStyle = "B") Then
  Items(ColItem).Item.Values(RowItem) = True
 ElseIf (iStyle = "O") Then '* Option.
  For isCol = 1 To TotalItems
   If (RowItem < UBound(Items(isCol).Item.Style) + 1) Then
    If (Items(isCol).Item.Style(RowItem) = "O") Then Items(isCol).Item.Values(RowItem) = False
   End If
  Next
  Items(ColItem).Item.Values(RowItem) = True
 ElseIf (iStyle = "Ch") Then
  Items(ColItem).Item.Values(RowItem) = Not (Items(ColItem).Item.Values(RowItem))
 End If
 Call Refresh(True)
 '* Button.
 If (iStyle = "B") Then
  Call Wait(0.04)
  Items(ColItem).Item.Values(RowItem) = False
  Call Refresh(False)
  Call FormShow(TheForm, TheText)
 End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 Enabled = PropBag.ReadProperty("Enabled", True)
 BackColor = PropBag.ReadProperty("BackColor", GetSysColor(COLOR_BTNFACE))
 HeadersColor = PropBag.ReadProperty("HeadersColor", GetSysColor(COLOR_BTNFACE))
 TextHeaders = PropBag.ReadProperty("TextHeaders", "")
End Sub

Private Sub UserControl_Resize()
 Call Refresh
End Sub

Private Sub UserControl_Show()
 If (isEnabled = False) Then m_StateG = -1 Else m_StateG = 1
 m_StateG = 0
 Call Refresh
End Sub

Private Sub UserControl_Terminate()
 TotalItems = 0
 ReDim Items(1)
 ReDim ItemsT(1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  Call .WriteProperty("Enabled", isEnabled, True)
  Call .WriteProperty("BackColor", m_lBackColor, GetSysColor(COLOR_BTNFACE))
  Call .WriteProperty("HeadersColor", m_lHeadersColor, GetSysColor(COLOR_BTNFACE))
  Call .WriteProperty("TextHeaders", m_sTextHeaders, "")
 End With
End Sub

Private Sub VScrollUpD_Change()
 Call Refresh
End Sub

Private Sub VScrollUpD_Scroll()
 Call VScrollUpD_Change
End Sub

Private Sub HScrollUpD_Change()
 Call Refresh
End Sub

Private Sub HScrollUpD_Scroll()
 Call HScrollUpD_Change
End Sub
