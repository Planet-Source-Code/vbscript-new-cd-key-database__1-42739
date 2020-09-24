VERSION 5.00
Begin VB.Form frmView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Serial Numbers"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   4710
   Begin VB.TextBox txtGenre 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   855
      Left            =   3360
      Picture         =   "frmView.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtSerial 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   3255
   End
   Begin VB.ComboBox cmbProd 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblGenre 
      Caption         =   "Genre"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblSerial 
      Caption         =   "Serial Number"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblCDKey 
      Caption         =   "CD-Key"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblProduct 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim db As Database
    Dim rs As Recordset

Private Sub Form_Load()
    On Error GoTo ErrHandler
    Set db = OpenDatabase(App.Path & "\data.mdb")
    Set rs = db.OpenRecordset("tblData", dbOpenDynaset)
    GetData
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

Private Sub cmdClose_Click()
    On Error GoTo ErrHandler
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
    
Public Sub cmbProd_Click()
    On Error GoTo ErrHandler
    rs.MoveFirst
Again:
    If cmbProd.Text = DeCode(rs.Fields("Prod")) Then
        txtGenre.Text = rs.Fields("Genre")
        txtKey.Text = DeCode(rs.Fields("Key"))
        txtSerial.Text = DeCode(rs.Fields("Serial"))
        txtName.Text = DeCode(rs.Fields("Name"))
        txtNotes.Text = DeCode(rs.Fields("Notes"))
    Else
        rs.MoveNext
        GoTo Again
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

Private Sub GetData()
    On Error GoTo ErrHandler
    rs.MoveFirst
    While Not rs.EOF
       cmbProd.AddItem DeCode(rs.Fields("Prod"))
       rs.MoveNext
    Wend
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
