Attribute VB_Name = "Concurrency"
Option Compare Database
Option Explicit

' Function to replace the Edit and Update methods of a DAO.Recordset
' to avoid errors from concurrent updates by two or more processes.
'
' Version 1.0.2
'
' 2016-02-06. Gustav Brock, Cactus Data ApS, CPH.


' Public variables.

    ' If True : Print error.
    ' If False: Silently handle error.
    Public DebugMode    As Boolean

' Function to replace the Edit method of a DAO.Recordset.
' To be used followed by GetUpdate to automatically handle
' concurrent updates.
'
' 2016-02-06. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub SetEdit(ByRef rs As DAO.Recordset)

    On Error GoTo Err_SetEdit

    ' Attempt to set rs into edit mode.
    Do While rs.EditMode <> dbEditInProgress
        rs.Edit
        If rs.EditMode = dbEditInProgress Then
            ' rs is ready for edit.
            Exit Do
        End If
    Loop

Exit_SetEdit:
    Exit Sub

Err_SetEdit:
    If DebugMode Then Debug.Print "    Edit", Timer, Err.Description
    If Err.Number = 3197 Then
        ' Concurrent edit.
        ' Continue in the loop.
        ' Will normally happen ONCE only for each call of SetEdit.
        Resume Next
    Else
        ' Other error, like deleted record.
        ' Pass error handling to the calling procedure.
        Resume Exit_SetEdit
    End If

End Sub

' Function to replace the Update method of a DAO.Recordset.
' To be used following SetEdit to automatically handle
' concurrent updates.
'
' 2016-01-31. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function GetUpdate(ByRef rs As DAO.Recordset) As Boolean

    On Error GoTo Err_GetUpdate
    
    ' Attempt to update rs and terminate edit mode.
    rs.Update
    
    GetUpdate = True

Exit_GetUpdate:
    Exit Function

Err_GetUpdate:
    If DebugMode Then Debug.Print "    Update", Timer, Err.Description
    ' Update failed.
    ' Cancel and return False.
    rs.CancelUpdate
    Resume Exit_GetUpdate

End Function
