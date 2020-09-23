VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo SMGrid 1.0b by HACKPRO TM © 2005"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin GridControl.SMGrid SMGrid 
      Height          =   3645
      Left            =   105
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   6429
      BackColor       =   16777215
      TextHeaders     =   "^Demo Version|~by HACKPRO TM"
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemo.frx":058A
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   105
      TabIndex        =   1
      Top             =   3825
      Width           =   6150
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2005       *'
'******************************************************'
'* Comments: This control was necessary to develop it *'
'*           for a program of a thesis of grade of my *'
'*           University, its evolution was stopped by *'
'*           a lot of time, although it is not comple-*'
'*           tely ended, but it's a beginning.        *'
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

 Private i As Integer

Private Sub Form_Load()
 With SMGrid
  .TextHeaders = "Column Header1|^Column Header2|Column Header3|Column Header4|~Column Header5"
  Call .AddItem("Col1 Row1|Col2 Row1", "&HEED5C4|&HE0E0E0|&HDEEDEF", , "T|B")
  Call .AddItem("Col1 Row2|Col2 Row2", , , "B|C")
  Call .AddItem("~Right Align|Col2 Row3", , , "N|C")
  Call .AddItem("Col1 Row4|Col2 Row4", , , "N|B")
  Call .AddItem("^Center Align|Col2 Row5", , , "T|T")
  For i = 6 To 8
   Call .AddItem("Col1 Row" & i & "|Col2 Row" & i & "|Col3 Row" & i, , , "T|Ch|O|O")
  Next
  For i = 9 To 16
   Call .AddItem("||||^Col5 Row" & i, , "||||&HC56A31", "C")
  Next
  Call .ObjectForm(frmPopUp, frmPopUp.txtValue)
 End With
End Sub
