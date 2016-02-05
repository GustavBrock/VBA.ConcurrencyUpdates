Attribute VB_Name = "ConcurrencyTest"
Option Compare Database
Option Explicit

' Functions to demonstrate handling of concurrent updates with DAO and
' the two functions SetEdit and GetUpdate.
'
' Function UpdateNoConcurrency is a normal function that updates a field.
' In case of a concurrent update, this will fail and pop a message for the user.
'
' Function UpdateConcurrency is the modified function that updates a field.
' In case of a concurrent update, this will not fail and not pop a message for the user
' but silently retry the update until success without bothering the user.'
'
' 2016-01-31. Gustav Brock, Cactus Data ApS, CPH.

' Common constants.

    ' Name of linked table to update (like Products of Northwind.
    Const TableName As String = "Products"
    ' Name of the field in TableName to update.
    Const FieldName As String = "Discontinued"
    ' Name of the primary key field in TableName.
    Const KeyName   As String = "ProductID"
    ' Value of the (arbitrary) primary key to look up and update.
    Const KeyValue  As Long = 20
    
    ' Duration in seconds for the concurrency test to run.
    Const Duration  As Integer = 3

Public Sub ConcurrencyTest()
    
    Dim db          As DAO.Database
    Dim rs          As DAO.Recordset
    Dim fd          As DAO.Field
    
    Dim StopTime    As Single
    Dim Delay       As Single
    Dim Attempts    As Long
    Dim LoopStart   As Single
    Dim LoopEnd     As Single
    Dim Loops       As Long
    
    Dim SQL         As String
    Dim Criteria    As String
    Dim NewValue    As Boolean
    
    SQL = "Select * From " & TableName & ""
    Criteria = KeyName & " = " & CStr(KeyValue) & ""
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(SQL, dbOpenDynaset, dbSeeChanges)
    
    rs.FindFirst Criteria
    Set fd = rs.Fields(FieldName)
    
    ' Set time for the test to stop.
    StopTime = Timer + Duration
    ' Let SetEdit and GetUpdate print debug information.
    DebugMode = True
    
    ' At random intervals, call updates of the field until StopTime is reached.
    While Timer < StopTime
        
        ' Postpone the next update.
        Delay = Timer + Rnd / 100
        While Timer < Delay
            DoEvents
        Wend
        Loops = Loops + 1
        LoopStart = Timer
        Debug.Print Loops, LoopStart
        
        ' Perform update.
        NewValue = Not fd.Value
        Do
            ' Count the attempts to update in this loop.
            Attempts = Attempts + 1
            ' Attempt edit and update until success.
            SetEdit rs
                fd.Value = NewValue
        Loop Until GetUpdate(rs)
        
        LoopEnd = Timer
        ' Print loop duration in milliseconds and edit attempts.
        Debug.Print , LoopEnd, Int(1000 * (LoopEnd - LoopStart)), Attempts
        Attempts = 0
        
    Wend
    rs.Close
    
    DebugMode = False
    Set fd = Nothing
    Set rs = Nothing
    Set db = Nothing
    
End Sub

Public Sub NoConcurrencyTest()
    
    Dim db          As DAO.Database
    Dim rs          As DAO.Recordset
    Dim fd          As DAO.Field
    
    Dim StopTime    As Single
    Dim Delay       As Single
    Dim LoopStart   As Single
    Dim LoopEnd     As Single
    Dim Loops       As Long
    
    Dim SQL         As String
    Dim Criteria    As String
    Dim NewValue    As Boolean
    
    SQL = "Select * From " & TableName & ""
    Criteria = KeyName & " = " & CStr(KeyValue) & ""
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(SQL, dbOpenDynaset, dbSeeChanges)
    
    rs.FindFirst Criteria
    Set fd = rs.Fields(FieldName)
    
    ' Set time for the test to stop.
    StopTime = Timer + Duration
    ' Let SetEdit and GetUpdate print debug information.
    DebugMode = True
    
    ' At random intervals, call updates of the field until StopTime is reached.
    While Timer < StopTime
        
        ' Postpone the next update.
        Delay = Timer + Rnd / 100
        While Timer < Delay
            DoEvents
        Wend
        Loops = Loops + 1
        LoopStart = Timer
        Debug.Print Loops, LoopStart
        
        ' Perform update.
        NewValue = Not fd.Value
        ' Attempt edit and update.
        rs.Edit
            fd.Value = NewValue
        rs.Update
        
        LoopEnd = Timer
        ' Print loop duration in milliseconds and edit attempts.
        Debug.Print , LoopEnd, Int(1000 * (LoopEnd - LoopStart))
        
    Wend
    rs.Close
    
    DebugMode = False
    Set fd = Nothing
    Set rs = Nothing
    Set db = Nothing
    
End Sub

Public Sub UpdateNoConcurrency()
    
    Dim db          As DAO.Database
    Dim rs          As DAO.Recordset
    Dim fd          As DAO.Field
    
    Dim SQL         As String
    Dim Criteria    As String
    Dim NewValue    As Boolean
    
    SQL = "Select * From " & TableName & ""
    Criteria = KeyName & " = " & CStr(KeyValue) & ""
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(SQL, dbOpenDynaset, dbSeeChanges)
    
    rs.FindFirst Criteria
    Set fd = rs.Fields(FieldName)
    NewValue = Not fd.Value
    
    rs.Edit
        fd.Value = NewValue
    rs.Update
    
    rs.Close
    
    Set fd = Nothing
    Set rs = Nothing
    Set db = Nothing
    
End Sub

Public Sub UpdateConcurrency()
    
    Dim db          As DAO.Database
    Dim rs          As DAO.Recordset
    Dim fd          As DAO.Field
    
    Dim SQL         As String
    Dim Criteria    As String
    Dim NewValue    As Boolean
    
    SQL = "Select * From " & TableName & ""
    Criteria = KeyName & " = " & CStr(KeyValue) & ""
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(SQL, dbOpenDynaset, dbSeeChanges)
    
    rs.FindFirst Criteria
    Set fd = rs.Fields(FieldName)
    NewValue = Not fd.Value
    
    Do
        SetEdit rs
            fd.Value = NewValue
    Loop Until GetUpdate(rs)
    
    rs.Close
    
    Set fd = Nothing
    Set rs = Nothing
    Set db = Nothing
    
End Sub
