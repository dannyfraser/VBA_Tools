Attribute VB_Name = "ListBuilder"
Option Compare Database
Option Explicit

Public Enum SelectionStatus
    Selected
    AllItems
End Enum

Public Function GetList(ListBox As ListBox, Column As Long, ItemSelection As SelectionStatus) As String
    
    Select Case ItemSelection
        Case Is = SelectionStatus.AllItems
            GetList = GetAllItems(ListBox, Column)
        Case Is = SelectionStatus.Selected
            GetList = GetSelectedItems(ListBox, Column)
    End Select
    
End Function

Private Function GetSelectedItems(ListBox As ListBox, Column As Long) As String
    Dim SelectedItem As Variant
    For Each SelectedItem In ListBox.ItemsSelected
        If Not IsNull(ListBox.Column(Column, SelectedItem)) Then
            GetSelectedItems = AddToNumberList(GetSelectedItems, ListBox.Column(Column, SelectedItem))
        End If
    Next SelectedItem
End Function

Private Function GetAllItems(ListBox As ListBox, Column As Long) As String
    Dim ListItem As Long
    For ListItem = 0 To ListBox.ListCount - 1
        If Not IsNull(ListBox.Column(Column, ListItem)) Then
            GetAllItems = AddToNumberList(GetAllItems, ListBox.Column(Column, ListItem))
        End If
    Next ListItem
End Function


Private Function AddToNumberList(List As String, NewItem As Variant) As String

    If Len(List) > 0 Then
        AddToNumberList = List & "," & CStr(NewItem)
    Else
        AddToNumberList = List & CStr(NewItem)
    End If
    
End Function
