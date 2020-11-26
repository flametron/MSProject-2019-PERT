Sub PERT()
Dim tskT As Task
Dim tskFirst As Task
Dim FoundBadWeights As Boolean
Dim UseOneSetofWeights As Boolean
Dim FL As MSProject.TableField
Dim found As Boolean

found = False

For Each FL In ActiveProject.TaskTables("Entry").TableFields
    If GetFieldName(FL) = "Duration1" Then
        found = True
        Exit For
    End If
Next

If found = False Then
    found = Check()
    MsgBox "Fields were not present, hence Inserted", Buttons:=vbInformation, Title:="PERT Field Insertion"
    Exit Sub
End If

If ActiveProject.Tasks.Count = 0 Then
    MsgBox Prompt:="There are no Tasks in Entry Table to work upon", Buttons:=vbCritical, Title:="PERT No Task Error"
    Exit Sub
End If

UseOneSetofWeights = True
FoundBadWeights = False

CustomFieldRename FieldID:=pjCustomTaskDuration1, NewName:="Optimistic Duration"
CustomFieldRename FieldID:=pjCustomTaskDuration2, NewName:="Most Likely Duration"
CustomFieldRename FieldID:=pjCustomTaskDuration3, NewName:="Pessimistic Duration"
CustomFieldRename FieldID:=pjCustomTaskNumber1, NewName:="Optimistic Weight"
CustomFieldRename FieldID:=pjCustomTaskNumber2, NewName:="Most Likely Weight"
CustomFieldRename FieldID:=pjCustomTaskNumber3, NewName:="Pessimistic Weight"
CustomFieldRename FieldID:=pjCustomTaskText30, NewName:="PERT State"

If UseOneSetofWeights = True Then
    Set tskFirst = ActiveProject.Tasks(1)
    For Each tskT In ActiveProject.Tasks
      If Not (tskT Is Nothing) Then
        If tskT.PercentComplete = 0 And tskT.PercentWorkComplete = 0 Then
          If (tskFirst.Number1 + tskFirst.Number2 + tskFirst.Number3) = 6 Then
            tskT.Duration = ((((tskT.Duration1) * tskFirst.Number1) _
            + ((tskT.Duration2) * tskFirst.Number2) _
            + ((tskT.Duration3) * tskFirst.Number3)) / 6)
            tskT.Text30 = "Duration Calc'd: " & Now()
          Else
            tskT.Text30 = "Not Calc'd: Weights <> 6"
            FoundBadWeights = True
          End If
        Else
          tskT.Text30 = "Not Calc'd: Task In Progress or Complete"
        End If
      End If
    Next tskT
Else
    For Each tskT In ActiveProject.Tasks
      If Not (tskT Is Nothing) Then
        If tskT.PercentComplete = 0 And tskT.PercentWorkComplete = 0 Then
          If (tskT.Number1 + tskT.Number2 + tskT.Number3) = 6 Then
            tskT.Duration = ((((tskT.Duration1) * tskT.Number1) _
            + ((tskT.Duration2) * tskT.Number2) _
            + ((tskT.Duration3) * tskT.Number3)) / 6)
            tskT.Text30 = "Duration Calc'd: " & Now()
          Else
            tskT.Text30 = "Not Calc'd: Weights <> 6"
            FoundBadWeights = True
          End If
        Else
          tskT.Text30 = "Not Calc'd: Task In Progress or Complete"
        End If
      End If
    Next tskT
End If
If FoundBadWeights = True Then
    MsgBox Prompt:="Some Tasks Weight Values were found to be incorrect." & _
    Chr(13) & "Check the PERT State field for details.", Buttons:=vbCritical, _
    Title:="PERT Weights Error"
    Exit Sub
End If
MsgBox "PERT Calculated!", Buttons:=vbInformation, Title:="PERT Analysis Successful"
End Sub



Private Function Check() As Boolean
Dim pos As Integer
Dim FL As MSProject.TableField

pos = 0

For Each FL In ActiveProject.TaskTables("Entry").TableFields

    If GetFieldName(FL) = "Duration" Then
        Exit For
    End If
    pos = pos + 1
Next

TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Duration1", Title:="Optimistic Duration", Width:=15, ShowInMenu:=True, ColumnPosition:=pos
TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Duration2", Title:="Most Likely Duration", Width:=15, ShowInMenu:=True, ColumnPosition:=pos + 1
TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Duration3", Title:="Pessimistic Duration", Width:=15, ShowInMenu:=True, ColumnPosition:=pos + 2
TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Number1", Title:="Optimistic Weight", Width:=15, ShowInMenu:=True, ColumnPosition:=pos + 3
TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Number2", Title:="Most Likely Weight", Width:=15, ShowInMenu:=True, ColumnPosition:=pos + 4
TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Number3", Title:="Pessimistic Weight", Width:=15, ShowInMenu:=True, ColumnPosition:=pos + 5
TableEditEx Name:="Entry", TaskTable:=True, NewFieldName:="Text30", Title:="PERT State", Width:=15, ShowInMenu:=True, ColumnPosition:=-1
TableApply "Entry"

Check = True

End Function

Private Function GetFieldName(ByVal objField As MSProject.TableField) As String
  Dim lngFieldID As Long
  Dim strResult As String

  lngFieldID = objField.Field

  With objField.Application
    strResult = Trim(.FieldConstantToFieldName(lngFieldID))
    On Error GoTo ErrorIfMinus1 ' CustomField does not handle lngFieldID= -1
    If Len(Trim(CustomFieldGetName(lngFieldID))) > 0 Then strResult2 = " (" & Trim(CustomFieldGetName(lngFieldID)) & ")" Else strResult2 = ""
  End With

  GetFieldName = strResult
Exit Function

ErrorIfMinus1:
  strResult2 = ""
  Resume Next
End Function
