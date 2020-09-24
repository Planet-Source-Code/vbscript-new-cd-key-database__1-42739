Attribute VB_Name = "Utilities"
Option Explicit
    Dim db As Database
    Dim rs As Recordset

Sub Main()
    On Error GoTo ErrHandler
    frmLogin.Show
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
    
Public Function GetEntries()
    On Error GoTo ErrHandler
    Dim Count As Integer
    Set db = OpenDatabase(App.Path & "\data.mdb")
    Set rs = db.OpenRecordset("tblData", dbOpenDynaset)
    rs.MoveFirst
    Count = 0
    Do Until rs.EOF
        Count = Count + 1
        rs.MoveNext
    Loop
    frmMain.stbMain.Panels(1).Text = "There are currently " & Count & " entries in the database."
    rs.Close
    db.Close
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

