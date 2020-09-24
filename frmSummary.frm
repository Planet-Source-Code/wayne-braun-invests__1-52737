VERSION 5.00
Begin VB.Form frmSummary 
   BackColor       =   &H00CBB674&
   Caption         =   "                                                Portfolio Summary"
   ClientHeight    =   7695
   ClientLeft      =   1710
   ClientTop       =   870
   ClientWidth     =   8025
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   8025
   Visible         =   0   'False
   Begin VB.ListBox lstResults 
      BackColor       =   &H00E6F2FF&
      Height          =   5460
      ItemData        =   "frmSummary.frx":0000
      Left            =   1035
      List            =   "frmSummary.frx":0002
      TabIndex        =   1
      Top             =   540
      Width           =   5910
   End
   Begin VB.CommandButton cmdCloseSummary 
      BackColor       =   &H00243DE8&
      Caption         =   "Close Window"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   3420
      Picture         =   "frmSummary.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6795
      Width           =   1095
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCloseSummary_Click()
   Unload Me
End Sub

