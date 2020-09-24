VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please enter your login information!"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   0
      Picture         =   "frmLogin.frx":1782
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   0
      Width           =   540
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Please Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password: "
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblUsername 
      Caption         =   "Username: "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim db As Database
    Dim rs As Recordset

Private Sub Form_Load()
    On Error GoTo ErrHandler
    Dim Cnt
    Set db = OpenDatabase(App.Path & "\data.mdb")
    Set rs = db.OpenRecordset("tblLogin", dbOpenDynaset)
    Cnt = rs.RecordCount
    If Cnt = 0 Then
        MsgBox "Enter your username and password to create an account.", vbOKOnly + vbExclamation, "Enter your new information"
        Exit Sub
    End If
    Exit Sub
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdLogin_Click()
    On Error GoTo ErrHandler
    Dim Username, Password, Cnt
    Cnt = rs.RecordCount
    If Cnt = 0 Then
        Username = Encode(txtUsername.Text)
        Password = Encode(txtPassword.Text)
        With rs
            .AddNew
                !Username = Username
                !Password = Password
            .Update
        End With
        frmMain.Show
    Else
        Authenticate
    End If
    Unload Me
    Exit Sub
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo ErrHandler
    Unload Me
    Unload frmMain
    Exit Sub
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Sub
End Sub

Function Authenticate()
    On Error GoTo ErrHandler
    Dim Username1, Username2, Password1, Password2, Cnt, x
    Dim ResU, ResP, Result, Title, Text
    rs.MoveFirst
    Username1 = txtUsername.Text
    Password1 = txtPassword.Text
    Username2 = DeCode(rs.Fields("Username"))
    Password2 = DeCode(rs.Fields("Password"))
    ResU = 0
    ResP = 0
    Result = 0
    If Username1 = Username2 Then ResU = 1
    If Password1 = Password2 Then ResP = 1
    Result = ResU + ResP
    Select Case Result
        Case 0
            Title = "Authentication Error!"
            Text = "The Username and Password you entered do not match."
            Text = Text & vbCrLf & "Please run the program again and enter the correct information!"
            MsgBox Text, vbCritical + vbOKOnly, Title
            Unload Me
            Unload frmMain
        Case 1
            Title = "Authentication Error!"
            Text = "The Username and Password you entered do not match."
            Text = Text & vbCrLf & "Please run the program again and enter the correct information!"
            MsgBox Text, vbCritical + vbOKOnly, Title
            Unload Me
            Unload frmMain
        Case 2
            frmMain.Show
            Unload Me
    End Select
    Exit Function
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Function
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    rs.Close
    db.Close
    Exit Sub
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Sub
End Sub

