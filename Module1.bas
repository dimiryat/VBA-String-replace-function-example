Attribute VB_Name = "Module1"
Option Explicit

Public Function ConvertToSINF(Value As String) As String

   Dim iter As Long
   Dim Result, Bin1, NonBin1, EmptyCell As String
   
   Result = "RowData:"
   Bin1 = "000 "
   NonBin1 = "031 "
   EmptyCell = "___ "
   
   For iter = 1 To Len(Value) Step 1
      If Mid(Value, iter, 1) = "." Then
         Result = Result + EmptyCell
      ElseIf Mid(Value, iter, 1) = "1" Then
         Result = Result + Bin1
      ElseIf Mid(Value, iter, 1) = "X" Then
         Result = Result + NonBin1
      Else
         Exit For
      End If
   Next iter
   
   ConvertToSINF = Result

End Function

