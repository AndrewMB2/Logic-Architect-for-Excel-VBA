VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Public Function ObjType() As String
End Function

Public Property Get HasError() As Boolean
End Property

Public Property Let HasError(RHS As Boolean)
End Property

Public Function GetType(Index As Variant) As Variant
End Function

Public Function DataRowCount() As Long
End Function

Public Function Column(Field As Variant) As Integer
End Function

Public Function GetHeader(Index As Variant) As Variant
End Function

Public Function Columns() As Integer
End Function

Public Function Exists(Field As Variant) As Boolean
End Function

Public Function RecStart() As Boolean
End Function

Public Function RecNext() As Boolean
End Function

Public Property Get RecPos() As Long
End Property

Public Property Let RecPos(Pos As Long)
End Property

Public Property Get RecValue(Field) As Variant
End Property

Public Property Let RecValue(Field, FieldValue As Variant)
End Property

'Only available as part of this interface
Public Property Get RecValue2(Field) As Variant
End Property

Public Property Let RecValue2(Field, FieldValue As Variant)
End Property

Public Function RecAdd() As Boolean
End Function

Public Function RecDelete() As Boolean
End Function

Public Function RecClose(Optional Cancel As Boolean = False) As Boolean
End Function

Public Function CreateIndex(ParamArray Fields()) As Boolean
End Function

Public Function RemoveIndex() As Boolean
End Function

Public Function Lookup(ParamArray Values()) As Variant
End Function


