VERSION 5.00
Begin VB.Form frmAddNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Entry"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   855
      Left            =   3360
      Picture         =   "frmAddNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Entry"
      Height          =   855
      Left            =   2040
      Picture         =   "frmAddNew.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1920
      Width           =   4455
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtSerial 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtKey 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtProd 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblGenre 
      Caption         =   "Genre"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
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
   Begin VB.Label lblCDKey 
      Caption         =   "CD-Key"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblProduct 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddNew"
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
    If frmSearch.Visible = True Then Unload frmSearch
    Set db = OpenDatabase(App.Path & "\data.mdb")
    Set rs = db.OpenRecordset("tblData", dbOpenDynaset)
    Set rsGenre = db.OpenRecordset("tblGenre", dbOpenDynaset)
    ClearForm
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

Private Sub cmdAdd_Click()
    On Error GoTo ErrHandler
    Dim Prod, Key, Serial, Name, Notes, Genre
    If txtProd.Text = "" And txtKey.Text = "" And _
        txtSerial.Text = "" And txtName.Text = "" And _
        txtNotes.Text = "" Then
        MsgBox "You must enter information in for at least the Porduct Name" _
        , vbOKOnly + vbCritical, "Data Required"
        ClearForm
        Exit Sub
    End If
    Genre = cmbGenre.Text
    Prod = Encode(txtProd.Text)
    Key = Encode(txtKey.Text)
    Serial = Encode(txtSerial.Text)
    Name = Encode(txtName.Text)
    Notes = Encode(txtNotes.Text)
    With rs
        .AddNew
            !Genre = Genre
            !Prod = Prod
            !Key = Key
            !Serial = Serial
            !Name = Name
            !Notes = Notes
        .Update
    End With
    ClearForm
    txtProd.SetFocus
    GetEntries
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
    ClearForm
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

Function ClearForm()
    On Error GoTo ErrHandler
    rsGenre.MoveFirst
    While Not rsGenre.EOF
       cmbGenre.AddItem rsGenre.Fields("Genre")
       rsGenre.MoveNext
    Wend
    txtProd.Text = ""
    txtKey.Text = ""
    txtSerial.Text = ""
    txtName.Text = ""
    txtNotes.Text = ""
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
    rsGenre.Close
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
