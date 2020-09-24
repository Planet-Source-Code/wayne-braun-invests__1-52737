VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Investment Log In"
   ClientHeight    =   6825
   ClientLeft      =   765
   ClientTop       =   2175
   ClientWidth     =   9495
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4032.435
   ScaleMode       =   0  'User
   ScaleWidth      =   8915.291
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   6435
      Picture         =   "frmLogin.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   13
      Top             =   2160
      Width           =   555
   End
   Begin VB.TextBox txtRndm 
      DataField       =   "RandomStr"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   180
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3060
      Top             =   6390
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change Password"
      Height          =   465
      Left            =   5535
      TabIndex        =   6
      Top             =   5220
      Width           =   2490
   End
   Begin VB.CommandButton cmdChangeOwnerName 
      Caption         =   "Change Owner Name"
      Height          =   465
      Left            =   5535
      TabIndex        =   5
      Top             =   4590
      Width           =   2490
   End
   Begin VB.TextBox txtConfirmPassword 
      BackColor       =   &H00E6F2FF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   5535
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3420
      Width           =   2310
   End
   Begin VB.TextBox txtOwner 
      BackColor       =   &H00E6F2FF&
      Height          =   345
      Left            =   5535
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1665
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00ADC4AA&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   525
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3870
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00243DE8&
      Cancel          =   -1  'True
      Caption         =   "Cancel   (Terminate Program without saving changes)"
      Height          =   525
      Left            =   6660
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3870
      Width           =   2265
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00E6F2FF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5535
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   2325
   End
   Begin VB.Label lblChanges 
      Caption         =   $"frmLogin.frx":0884
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   270
      TabIndex        =   11
      Top             =   4455
      Width           =   5100
   End
   Begin VB.Label lblPassword 
      Caption         =   $"frmLogin.frx":0A46
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   540
      TabIndex        =   10
      Top             =   2565
      Width           =   4920
   End
   Begin VB.Label lblOwner 
      Caption         =   "Owner Name:  (Use up to 20 letters and/or digits) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   900
      TabIndex        =   9
      Top             =   1710
      Width           =   4560
   End
   Begin VB.Label lblConfirmPassword 
      Caption         =   "Confirm Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3735
      TabIndex        =   8
      Top             =   3465
      Width           =   1680
   End
   Begin VB.Label lblNewOwner 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frmLogin.frx":0AE4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1365
      Left            =   270
      TabIndex        =   7
      Top             =   135
      Visible         =   0   'False
      Width           =   8520
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---Login is case sensitive:
'---Owner Name and Password are stored in 5th record of Random table
'---Owner Name length is Asc(1st char)-36
'---Owner Name is stored beginning in 2nd char and is decoded by
'---shifting up 2 ascii characters.
'---Password length is at 22nd char and begins at 23rd.
'---Decoding same as Owner Name
'---1st character of 5th record changed from d when
'---owner name and password are intiialized
Option Explicit
Public intLenOwner As Integer  '---length of owner name
Public intLenPassword As Integer  '---length of password
Public strBuildName As String   '---to build decoded owner name
Public strBuildPassword           '---to build decoded password
Public strBuildCoded As String   '---to build coded owner name and password
Public PasswordAccepted As Boolean  '---flag true if entered password
                                    '---matched saved password
Public PasswordLegit As Boolean    '---flag true if password & confirm password
                                   '---match & formatted ok
Dim I As Integer

Private Sub Form_Load()
   '---get Random records for Login use
   Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
   App.Path & "\Invests.mdb; Persist Security Info=False; Jet OLEDB:Database Password=fred"
   Adodc1.CursorLocation = adUseClient
   Adodc1.CommandType = adCmdText
   Adodc1.CursorType = adOpenStatic '---this may need to be adOpenStatic
   Adodc1.LockType = adLockOptimistic '---Guarantee that record being edited can be saved
   Adodc1.RecordSource = "Select * From Random"
   Adodc1.Refresh
   Adodc1.Recordset.MoveFirst
   Adodc1.Recordset.MoveNext
   Adodc1.Recordset.MoveNext
   Adodc1.Recordset.MoveNext
   Adodc1.Recordset.MoveNext   '---now at 5th record
   strBuildCoded = txtRndm.Text '---move codes to modifiable string
   If Mid$(strBuildCoded, 1, 1) = "d" Then   '---if Owner Name & password not initialized
      lblNewOwner.Visible = True
      lblChanges.Visible = False
      cmdChangeOwnerName.Enabled = False
      cmdChangePassword.Enabled = False
   Else
      lblNewOwner.Visible = False
      lblChanges.Visible = True
      cmdChangeOwnerName.Enabled = True
      txtOwner.Enabled = False  '---cannot change owner name w/o click change
      cmdChangePassword.Enabled = True
      lblConfirmPassword.Visible = False
      txtConfirmPassword.Visible = False
      If Mid$(strBuildCoded, 1, 1) <> "d" Then   '---if OwnerName inititalized
         intLenOwner = Asc(Mid$(strBuildCoded, 1, 1)) - 36 '---get length of OwnerName
         strBuildName = ""   '---decode and
         For I = 1 To intLenOwner   '---fill in OwnerName
            strBuildName = strBuildName + Chr$(Asc(Mid$(txtRndm.Text, I + 1, 1)) + 2)
         Next I
         OwnerName = strBuildName   '--- and set Owner Name for Program to use
         txtOwner.Text = OwnerName
      End If
   End If
End Sub

Private Sub cmdCancel_Click()
   MsgBox ("You must complete Log In to access the program.  The program will now terminate.")
   End
End Sub

Private Sub cmdOK_Click()
   strBuildCoded = txtRndm.Text '---move codes to new copy of modifiable string
   If cmdOK.Caption = "Save New Owner Name" Then
      If txtOwner = "" Then
         MsgBox "Owner Name must be established."
         txtOwner.SetFocus
         Exit Sub    '---exit without saving OwnerName
      Else
         Call CodeOwnerName
         Call SaveNameAndPassword
         Exit Sub
      End If
   End If
   If cmdOK.Caption = "Save New Password" Then
      Call EstablishPassword(PasswordLegit)
      If PasswordLegit = False Then
         Exit Sub
      End If
   End If
   '---now get the OwnerName for the first time
   If Mid$(strBuildCoded, 1, 1) = "d" Then   '---if OwnerName & password not inititalized
      If txtOwner = "" Then
         MsgBox "Owner Name must be established."
         txtOwner.SetFocus
         Exit Sub    '---exit without saving OwnerName
      Else
         Call CodeOwnerName
      End If
      '---now get the password for the first time
      Call EstablishPassword(PasswordLegit)
      If PasswordLegit = False Then
         Exit Sub
      Else
         Call SaveNameAndPassword
      End If
   Else
      '---code follows for OwnerName and password previously initialized
      Call DecodeOwnerName
      '---now decode and compare the saved password
      Call CheckPassword(PasswordAccepted)
      If PasswordAccepted = False Then
         Exit Sub
      Else
         Me.Hide    '---log in was succcessful
         frmMain.Show  '---proceed with program
      End If
   End If
End Sub

Private Sub cmdChangeOwnerName_Click()
   cmdOK.Caption = "Save New Owner Name"
   Call CheckPassword(PasswordAccepted)
   If PasswordAccepted = False Then
      cmdOK.Caption = "Ok"
      Exit Sub
   Else
      txtOwner.Enabled = True
      txtOwner = ""
      txtOwner.SetFocus
   End If
End Sub
Private Sub cmdChangePassword_Click()
   cmdOK.Caption = "Save New Password"
   Call CheckPassword(PasswordAccepted)
   If PasswordAccepted = False Then
      MsgBox "The Password entered was not valid.  The Log In will now start over."
      cmdOK.Caption = "Ok"
      Exit Sub
   Else
      txtPassword.Enabled = True
      txtPassword = ""
      txtConfirmPassword.Visible = True
      txtConfirmPassword = ""
      txtPassword.SetFocus
   End If
End Sub

Private Sub CodeOwnerName()
   Mid$(strBuildCoded, 1, 1) = Chr$(Len(txtOwner) + 36) '---set coded length of OwnerName
   For I = 1 To Len(txtOwner)   '---fill in coded OwnerName
      Mid$(strBuildCoded, I + 1, 1) = Chr$(Asc(Mid$(txtOwner, I, 1)) - 2)
   Next I
   OwnerName = txtOwner '--- and set Owner Name for Program to use
End Sub

Private Sub EstablishPassword(Ok As Boolean)   'get a newpassword &
                                         '---confirm, save it
   Ok = True
   If Len(txtPassword) < 4 Then
      MsgBox "A password of at least 4 characters must be established.  The Log In will now start over."
      txtOwner = ""
      txtPassword = ""
      Ok = False
      Exit Sub  '---exit without saving OwnerName or password
   Else
      intLenPassword = Len(txtPassword)
      Mid$(strBuildCoded, 22, 1) = Chr$(intLenPassword + 36) '---put the coded length in strBuildCoded
      For I = 1 To intLenPassword  '---put the coded password into strBuildCoded
         Mid$(strBuildCoded, I + 22, 1) = Chr$(Asc(Mid$(txtPassword, I, 1)) - 2)
      Next I
   End If
   If txtConfirmPassword = "" Or txtConfirmPassword <> txtPassword Then
      MsgBox "The Password did not match the Confirm Password.  The Log In will now start over."
      txtOwner = ""
      txtPassword = ""
      txtConfirmPassword = ""
      txtOwner.SetFocus
      Ok = False
      Exit Sub   '---exit without saving Owner Name & password
   Else
      Call SaveNameAndPassword
   End If
End Sub

Private Sub SaveNameAndPassword()
   txtRndm.Text = strBuildCoded   '---save the new coded OwnerName & password
   Adodc1.Recordset.MoveNext
   Me.Hide   '---the log in was successful
   frmMain.Show   '---proceed with program
End Sub

Private Sub DecodeOwnerName()
   intLenOwner = Asc(Mid$(strBuildCoded, 1, 1)) - 36
   OwnerName = "" '---build the decoded owner name
   For I = 1 To intLenOwner
      OwnerName = OwnerName + Chr$(Asc(Mid$(strBuildCoded, I + 1, 1)) + 2)
   Next I
   txtOwner.Text = OwnerName
End Sub

Private Sub CheckPassword(Passed As Boolean)  '---compare entered password to
                                              '---saved password
   Passed = True
   intLenPassword = Asc(Mid$(strBuildCoded, 22, 1)) - 36
   strBuildPassword = ""
   For I = 1 To intLenPassword
      strBuildPassword = strBuildPassword + Chr$(Asc(Mid$(strBuildCoded, I + 22, 1)) + 2)
   Next I
   If txtPassword = "" Or txtPassword <> strBuildPassword Then
      MsgBox "The Password entered was not valid.  The Log In will now start over."
      txtPassword = ""
      txtPassword.SetFocus
      Passed = False
   End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
   '---allow txtPassword to take only Ascii CHR$(32) to Chr$(122)
   '---else beep and leave cursor where it is.
   Select Case KeyAscii
      Case 8   '---backspace key
      Case 32 To 122  '---Ascii CHR$(32) to Chr$(122)
      Case Else:
         Beep
         KeyAscii = 0
   End Select
End Sub

Private Sub txtConfirmPassword_KeyPress(KeyAscii As Integer)
   '---allow txtConfirmPassword to take only Ascii CHR$(32) to Chr$(122)
   '---else beep and leave cursor where it is.
   Select Case KeyAscii
      Case 8   '---backspace key
      Case 32 To 122  '---Ascii CHR$(32) to Chr$(122)
      Case Else:
         Beep
         KeyAscii = 0
   End Select
End Sub

Private Sub txtOwner_KeyPress(KeyAscii As Integer)
   '---allow txtOwner to take only Ascii CHR$(32), A-Z, a-z, 0-9
   '---else beep and leave cursor where it is.
   Select Case KeyAscii
      Case 8   '---backspace key
      Case 32   '---space
      Case 48 To 57  '---0 to 9
      Case 65 To 90   '---A to Z
      Case 97 To 122   '---a to z
      Case Else:
         Beep
         KeyAscii = 0
   End Select
End Sub
