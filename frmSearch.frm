VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   Caption         =   "Search by Software Type"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   8790
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid grdResults 
      Height          =   3375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.ComboBox cmbGenre 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Type ""ALL"" to return all entries!"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblGenre 
      Caption         =   "Software Type"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSearch"
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
    ClearSearch
    rsGenre.MoveFirst
    While Not rsGenre.EOF
       cmbGenre.AddItem rsGenre.Fields("Genre")
       rsGenre.MoveNext
    Wend
    cmbGenre.Text = "All"
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

Private Sub cmdClear_Click()
    On Error GoTo ErrHandler
    ClearSearch
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

Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler
    Dim Genre As String, strSearch As String
    ClearSearch
    Genre = cmbGenre.Text
    If Genre = "All" Or Genre = "all" Then
        ReturnAll
    Else
        strSearch = "[Genre] Like '" & Genre & "'"
        With rs
            .FindFirst strSearch
            If .NoMatch Then
                Exit Sub
            Else
                grdResults.Row = 1
                grdResults.Col = 0
                grdResults.Text = DeCode(rs.Fields("Genre"))
                grdResults.Col = 1
                grdResults.Text = DeCode(rs.Fields("Prod"))
                grdResults.Col = 2
                grdResults.Text = DeCode(rs.Fields("Key"))
                grdResults.Col = 3
                grdResults.Text = DeCode(rs.Fields("Serial"))
                grdResults.Col = 4
                grdResults.Text = DeCode(rs.Fields("Name"))
                grdResults.Col = 5
                grdResults.Text = DeCode(rs.Fields("Notes"))
                Again strSearch
            End If
        End With
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

Private Sub Again(strSearch As String)
    On Error GoTo ErrHandler
    Dim TheRow As Integer
    With rs
        .FindNext strSearch
        If .NoMatch Then
            ' Exit Sub
        Else
            grdResults.Rows = grdResults.Rows + 1
            grdResults.Row = grdResults.Rows - 1
            grdResults.Col = 0
            grdResults.Text = DeCode(rs.Fields("Genre"))
            grdResults.Col = 1
            grdResults.Text = DeCode(rs.Fields("Prod"))
            grdResults.Col = 2
            grdResults.Text = DeCode(rs.Fields("Key"))
            grdResults.Col = 3
            grdResults.Text = DeCode(rs.Fields("Serial"))
            grdResults.Col = 4
            grdResults.Text = DeCode(rs.Fields("Name"))
            grdResults.Col = 5
            grdResults.Text = DeCode(rs.Fields("Notes"))
            Again strSearch
        End If
    End With
    grdResults.ColSel = 2
    grdResults.Sort = 1
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

Private Sub ReturnAll()
    On Error GoTo ErrHandler
    rs.MoveFirst
    Do While Not rs.EOF
        grdResults.Row = grdResults.Rows - 1
        grdResults.Col = 0
        grdResults.Text = DeCode(rs.Fields("Genre"))
        grdResults.Col = 1
        grdResults.Text = DeCode(rs.Fields("Prod"))
        grdResults.Col = 2
        grdResults.Text = DeCode(rs.Fields("Key"))
        grdResults.Col = 3
        grdResults.Text = DeCode(rs.Fields("Serial"))
        grdResults.Col = 4
        grdResults.Text = DeCode(rs.Fields("Name"))
        grdResults.Col = 5
        grdResults.Text = DeCode(rs.Fields("Notes"))
        grdResults.Rows = grdResults.Rows + 1
        rs.MoveNext
    Loop
    grdResults.Rows = grdResults.Rows - 1
    grdResults.ColSel = 1
    grdResults.Sort = 1
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

Private Sub ClearSearch()
    On Error GoTo ErrHandler
    grdResults.Clear
    grdResults.Rows = 2
    grdResults.Row = 0
    grdResults.Col = 0
    grdResults.CellFontBold = True
    grdResults.Text = "Genre"
    grdResults.Col = 1
    grdResults.CellFontBold = True
    grdResults.Text = "Product"
    grdResults.Col = 2
    grdResults.CellFontBold = True
    grdResults.Text = "Product Key"
    grdResults.Col = 3
    grdResults.CellFontBold = True
    grdResults.Text = "Serial Number"
    grdResults.Col = 4
    grdResults.CellFontBold = True
    grdResults.Text = "Owner Name"
    grdResults.Col = 5
    grdResults.CellFontBold = True
    grdResults.Text = "Comments"
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

Private Sub Form_Resize()
    On Error GoTo ErrHandler
    Dim frmHeight As Integer, frmWidth As Integer
    Dim grdHeight As Integer, grdWidth As Integer
    Dim colWidth As Integer, x As Integer
    frmHeight = Me.Height
    frmWidth = Me.Width
    grdHeight = frmHeight - 870
    grdWidth = frmWidth - 135
    grdResults.Height = grdHeight
    grdResults.Width = grdWidth
    colWidth = (grdWidth / 6) - 25
    For x = 0 To 5
        grdResults.ColAlignment(x) = 4
        grdResults.colWidth(x) = colWidth
    Next
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

Private Sub grdResults_DblClick()
    On Error GoTo ErrHandler
    Dim Product As String
    frmView.Show
    grdResults.Col = 1
    Product = grdResults.Text
    frmView.cmbProd = Product
    frmView.cmbProd_Click
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
