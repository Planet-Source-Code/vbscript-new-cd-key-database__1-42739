VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Entry"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   4695
   Begin VB.ComboBox cmbGenre 
      Height          =   315
      Left            =   1320
      TabIndex        =   11
      Top             =   120
      Width           =   3255
   End
   Begin VB.ComboBox cmbProd 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtSerial 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Entry"
      Height          =   855
      Left            =   2040
      Picture         =   "frmUpdate.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   855
      Left            =   3360
      Picture         =   "frmUpdate.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lblGenre 
      Caption         =   "Genre"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblProduct 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblCDKey 
      Caption         =   "CD-Key"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblSerial 
      Caption         =   "Serial Number"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim db As Database
    Dim rs As Recordset
    Dim rsGenre As Recordset

Private Sub Form_Load()
    On Error GoTo ErrHandler
    Set db = OpenDatabase(App.Path & "\data.mdb")
    Set rs = db.OpenRecordset("tblData", dbOpenDynaset)
    Set rsGenre = db.OpenRecordset("tblGenre", dbOpenDynaset)
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

Private Sub cmdUpdate_Click()
    On Error GoTo ErrHandler
    With rs
        .Edit
            !Genre = cmbGenre.Text
            !Prod = Encode(cmbProd.Text)
            !Key = Encode(txtKey.Text)
            !Serial = Encode(txtSerial.Text)
            !Name = Encode(txtName.Text)
            !Notes = Encode(txtNotes.Text)
        .Update
    End With
    ClearForm
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
    
Private Sub cmbProd_Click()
    On Error GoTo ErrHandler
    rs.MoveFirst
Again:
    If cmbProd.Text = DeCode(rs.Fields("Prod")) Then
        cmbGenre.Text = rs.Fields("Genre")
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
    rsGenre.MoveFirst
    While Not rsGenre.EOF
       cmbGenre.AddItem rsGenre.Fields("Genre")
       rsGenre.MoveNext
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

Private Sub ClearForm()
    On Error GoTo ErrHandler
    cmbProd.Clear
    cmbGenre.Clear
    txtKey.Text = ""
    txtSerial.Text = ""
    txtName.Text = ""
    txtNotes.Text = ""
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
