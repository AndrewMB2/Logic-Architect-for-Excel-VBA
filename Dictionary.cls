VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Logic Architect for Excel/VBA - Toolset for manipulating tabular data efficiently in Excel VBA
'Copyright (C) 2025  Andrew Burnett
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <https://www.gnu.org/licenses/>.
'Version 2.0

Public Enum CompareMethod
    BinaryCompare = 0
    TextCompare = 1
    DatabaseCompare = 2
End Enum

Dim DictValues As New Collection
Dim DictKeys As New Collection
Private Const Dummy As String = "||||||"

'Only TextCompare is supported
Property Get CompareMode() As CompareMethod
    CompareMode = TextCompare
End Property

Property Let CompareMode(CompareModeValue As CompareMethod)
    If CompareModeValue <> TextCompare Then MsgBox "Only TextCompare is supported"
End Property

Sub Add(Key, Value)
    DictValues.Add Value, Key
    DictKeys.Add Key, Key
End Sub

Function Count()
    Count = DictValues.Count
End Function

Function Exists(Key) As Boolean
    Dim Vtype As Long
    
    On Error Resume Next
    Vtype = VarType(DictValues(Key))
    If Err = 0 Then
        On Error GoTo 0
        Exists = True
    Else
        On Error GoTo 0
        Exists = False
    End If
End Function

Property Get Item(Key)
Attribute Item.VB_UserMemId = 0
'This is the default method.
'Attribute Item.VB_UserMemId = 0
    If IsObject(DictValues(Key)) Then
        Set Item = DictValues(Key)
    Else
        Item = DictValues(Key)
    End If
End Property

Property Let Item(Key, Value)
    DictValues.Add Value, Dummy, , Key
    DictValues.Remove Key
    DictValues.Add Value, Key, , Dummy
    DictValues.Remove Dummy
End Property

Property Set Item(Key, Value)
    DictValues.Add Value, Dummy, , Key
    DictValues.Remove Key
    DictValues.Add Value, Key, , Dummy
    DictValues.Remove Dummy
End Property

Function Items()
    Dim AllItems()
    Dim I As Long
    
    If DictValues.Count > 0 Then
        ReDim AllItems(DictValues.Count - 1)
        For I = 1 To DictValues.Count
            If IsObject(DictValues(I)) Then
                Set AllItems(I - 1) = DictValues(I)
            Else
                AllItems(I - 1) = DictValues(I)
            End If
        Next I
        Items = AllItems
    Else
        Items = VBA.Array()
    End If
End Function

Property Let Key(OldKey, NewKey)
    DictValues.Add DictValues(OldKey), NewKey, , OldKey
    DictValues.Remove OldKey
    DictKeys.Add NewKey, NewKey, , OldKey
    DictKeys.Remove OldKey
End Property

Function Keys()
    Dim AllKeys()
    Dim I As Long
    
    If DictKeys.Count > 0 Then
        ReDim AllKeys(DictKeys.Count - 1)
        For I = 1 To DictKeys.Count
            AllKeys(I - 1) = DictKeys(I)
        Next I
        Keys = AllKeys
    Else
        Keys = VBA.Array()
    End If
End Function

Function Remove(Key)
    DictValues.Remove Key
    DictKeys.Remove Key
End Function

Function RemoveAll()
    Set DictValues = New Collection
    Set DictKeys = New Collection
End Function
