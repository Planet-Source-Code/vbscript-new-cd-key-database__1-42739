VERSION 5.00
Begin VB.Form frmDelete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Record"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frmDelete.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   4695
   Begin VB.ComboBox cmbProd 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   855
      Left            =   3000
      Picture         =   "frmDelete.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Record"
      Height          =   855
      Left            =   1200
      Picture         =   "frmDelete.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblProduct 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmDelete"
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

Private Sub cmdDelete_Click()
    On Error GoTo ErrHandler
    Set db = OpenDatabase(App.Path & "\data.mdb")
    Set rs = db.OpenRecordset("tblData", dbOpenDynaset)
    rs.MoveFirst
    Do Until cmbProd.Text = DeCode(rs.Fields("Prod"))
        rs.MoveNext
    Loop
    rs.Delete
    GetEntries
    Unload Me
    Exit Sub
ErrHandler:
    Unload Me
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
