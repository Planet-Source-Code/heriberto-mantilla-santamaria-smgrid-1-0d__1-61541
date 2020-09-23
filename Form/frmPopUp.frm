VERSION 5.00
Begin VB.Form frmPopUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PopUp Demo - SMGrid"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frmPopUp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtValue 
      Height          =   2550
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblTitle1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Write your opinion for this control."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2910
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2005       *'
'******************************************************'
'* Comments: PopUp form for SMGrid                    *'
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
