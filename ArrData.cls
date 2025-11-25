VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Implements XData

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

Const RepChar1 As String = "\1" 'replaces \
Const RepChar2 As String = "\2" 'replaces ,
Const RepChar3 As String = "\3" 'replaces ;
Const RepChar4 As String = "\4" 'replaces "

Private Type FieldType
    Name As String
    VarType As VbVarType
    Indexed As Boolean
    Data As Variant '1-based
    CalcProc As Variant 'Calculation procedure or formula
    CalcFields As Variant 'Fields in CalcProc
    CalcOffset As Variant 'True to use previous value
    Params As Variant 'Array of parameters for calculation procedure
    PrevVarType As VbVarType 'Previous VarType before change in Group (0 if none as VbEmpty is not used)
End Type

Private Type FieldType2 'for SaveBin
    ThisObjName As String
    Name As String
    VarType As Integer
    Data As Variant '1-based
End Type

Private Type BinDataType
    ThisObjName As String
    DataRowCount As Long 'for SaveBin
    DataCount As Long
    Pointer As Variant
    FldCount As Long
End Type

Private Enum TestType
    TypEquals = 1
    TypGreater = 2
    TypGE = 3
    TypLess = 4
    TypLE = 5
    TypNE = 6
    TypIn = 7
    TypIgnore = 8
End Enum

Public Enum MatchType
    MatchMaxLE = 1
    MatchEqual = 0
    MatchMinGE = -1
End Enum

Public Enum RowIndexType
    IndexMultiple = 0
    IndexUnique = 1
End Enum

Public HasError As Boolean
Public ThisObjName As String
Public IndexType As RowIndexType
Public DataRowCount As Long 'number of logical records in Pointer
Public GroupData As Variant '1-based array of values for callback
Public Xs As New XShared
Dim DataCount As Long 'number of physical records in Data
Dim RecordPos As Long
Dim FieldList As Dictionary
Dim Fld() As FieldType '1-based
Dim Pointer() As Long
Dim RowIndex As Dictionary
Const SepChar As Long = &H25C6
Const GranSize As Long = 100

Public Function ObjType() As String
    ObjType = "ArrData"
End Function

Private Function XData_ObjType() As String
    XData_ObjType = "ArrData"
End Function

Private Function XData_DataRowCount() As Long
    If HasError Then Exit Function
    XData_DataRowCount = DataRowCount
End Function

Private Property Get XData_HasError() As Boolean
    XData_HasError = HasError
End Property

Private Property Let XData_HasError(RHS As Boolean)
    HasError = RHS
End Property

Public Function Create(Headers, ParamArray RecordTypes()) As Boolean
'Create a new dataset.
'Headers and RecordTypes must be matching fields/arrays.
    
    Dim Vtype As Integer
    Dim I As Long
    Dim Headers2 As Variant
    Dim RecordTypes2 As Variant
    Dim RecordTypesx As Variant
    
    If HasError Then Exit Function
    ThisObjName = vbNullString
    IndexType = IndexMultiple
    DataRowCount = 0
    DataCount = 0
    RecordPos = 0
    Set RowIndex = Nothing
    
    Headers2 = Xs.ExtractArray(Headers)
    RecordTypesx = RecordTypes
    RecordTypes2 = Xs.ExtractArray(RecordTypesx)
    If UBound(Headers2) < LBound(Headers2) Then
        Msg "No fields"
        Exit Function
    ElseIf UBound(RecordTypes2) <> UBound(Headers2) Or _
        LBound(RecordTypes2) <> LBound(Headers2) Then
        Msg "RecordTypes and Headers must match"
        Exit Function
    End If
    
    Set FieldList = New Dictionary
    FieldList.CompareMode = TextCompare
    ReDim Fld(1 To UBound(Headers2) - LBound(Headers2) + 1)
    ReDim Pointer(1 To GranSize)
    
    For I = LBound(Headers2) To UBound(Headers2)
        If FieldList.Exists(Headers2(I)) Then
            Msg "Duplicate header " & Headers2(I)
            HasError = True
            Exit Function
        Else
            FieldList.Add Headers2(I), I - LBound(Headers2) + 1
        End If
        
        With Fld(I - LBound(Headers2) + 1)
            .Name = Headers2(I)
            .VarType = RecordTypes2(I)
        End With
        If Not DimData(I - LBound(Headers2) + 1, GranSize) Then Exit Function
    Next I
    
    Xs.Show Me
    Create = True
End Function

Public Function AddField(Field As String, RecordType As VbVarType, Optional ByVal FieldValue As Variant) As Boolean
'Add a new field.

    Dim I As Long
    
    If HasError Then Exit Function
    If Exists(Field) Then
        Msg "Duplicate field name " & Field
        Exit Function
    End If
    
    ReDim Preserve Fld(1 To UBound(Fld) + 1)
    With Fld(UBound(Fld))
        .Name = Field
        .VarType = RecordType
    End With
    FieldList.Add Field, UBound(Fld)
    If DataCount = 0 Then
        If Not DimData(UBound(Fld), GranSize) Then Exit Function
    Else
        If Not DimData(UBound(Fld), UBound(Fld(1).Data)) Then Exit Function
    End If
    
    If Not IsMissing(FieldValue) Then
        With Fld(UBound(Fld))
            If Not ConvertType(FieldValue, .VarType) Then Exit Function
            For I = 1 To DataRowCount
                .Data(Pointer(I)) = FieldValue
            Next I
        End With
    End If
    
    Xs.Show Me
    AddField = True
End Function

Public Function AddFields(Xd As XData, Optional DataType As VbVarType) As Boolean
'Adds all the fields in xd not here already.
'If Xd is a TableData object, uses DataType.
    
    Dim UseType As VbVarType
    Dim I As Long
    
    If HasError Then Exit Function
    If Xd.HasError Then
        HasError = True
        Exit Function
    End If
    For I = 1 To Xd.Columns
        If Not Exists(Xd.GetHeader(I)) Then
            UseType = Xd.GetType(I)
            If UseType = -1 Then UseType = DataType
            Xs.NoShow = True
            If Not AddField(Xd.GetHeader(I), UseType) Then Exit Function
        End If
    Next I
    
    Xs.Show Me
    AddFields = True
End Function

Public Function DropField(ParamArray Field() As Variant) As Boolean
'Drops one or more fields.

    Dim Fieldx As Variant
    Dim Field2 As Variant
    Dim I As Long, J As Long
    
    If HasError Then Exit Function
    Fieldx = Field
    Field2 = Xs.ExtractArray(Fieldx)
    
    For I = LBound(Field2) To UBound(Field2)
        If Not FieldCheck(Field2(I)) Then Exit Function
        With Fld(FieldCol(Field2(I)))
            If .Indexed Then RemoveIndex
            .Data = Null
        End With
    Next I
    
    For I = 1 To UBound(Fld)
        If Not IsNull(Fld(I).Data) Then
            J = J + 1
            If J <> I Then
                Fld(J) = Fld(I)
            End If
        End If
    Next I
    
    If J = 0 Then
        Msg "No fields remain"
        Exit Function
    End If
    ReDim Preserve Fld(1 To J)
    
    Set FieldList = New Dictionary
    FieldList.CompareMode = TextCompare
    For I = 1 To J
        FieldList.Add Fld(I).Name, I
        With Fld(I)
            .CalcFields = Empty
            .CalcOffset = Empty
            .CalcProc = Empty
            .Params = Empty
        End With
    Next I
    
    Xs.Show Me
    DropField = True
End Function

Public Function ChangeType(Field As Variant, RecordType As VbVarType) As Boolean
'Changes the data type of a field.

    Dim Field2 As Long
    Dim NewData As Variant
    Dim I As Long
    
    If HasError Then Exit Function
    If Not FieldCheck(Field) Then Exit Function
    Field2 = FieldCol(Field)
    If DataCount = 0 Then
        If Not DimArray(NewData, GranSize, RecordType) Then Exit Function
    Else
        If Not DimArray(NewData, DataCount, RecordType) Then Exit Function
    End If
    
    With Fld(Field2)
        RecStart
        Do While RecNext
            I = Pointer(RecordPos)
            On Error Resume Next
            If .VarType = vbString Then
                If Len(.Data(I)) = 0 Then
                    NewData(I) = Empty
                Else
                    NewData(I) = .Data(I)
                End If
            Else
                NewData(I) = .Data(I)
            End If
            If Err = 0 Then
                On Error GoTo 0
            Else
                On Error GoTo 0
                Msg "Bad data " & CStr(.Data(I)) & " on field " & FieldName(Field2)
                Exit Function
            End If
        Loop
        .Data = NewData
        .VarType = RecordType
    End With
    
    Xs.Show Me
    ChangeType = True
End Function

Public Function GetHeader(Index As Variant) As Variant

    If HasError Then Exit Function
    If FieldCheck(Index) Then GetHeader = FieldName(Index)
End Function

Private Function XData_GetHeader(Index As Variant) As Variant
'Returns single header (Index is 1-based)

    If HasError Then Exit Function
    If FieldCheck(Index) Then XData_GetHeader = FieldName(Index)
End Function

Public Function GetType(Index As Variant) As Variant
'Returns VarType of field (Index is 1-based)

    If HasError Then Exit Function
    If FieldCheck(Index) Then GetType = Fld(FieldCol(Index)).VarType
End Function

Private Function XData_GetType(Index As Variant) As Variant

    If HasError Then Exit Function
    If FieldCheck(Index) Then XData_GetType = Fld(FieldCol(Index)).VarType
End Function

Public Function Column(Field As Variant) As Integer

    If HasError Then Exit Function
    If FieldCheck(Field) Then Column = FieldCol(Field)
End Function

Private Function XData_Column(Field As Variant) As Integer
'Returns the column number of the given field.

    If HasError Then Exit Function
    If FieldCheck(Field) Then XData_Column = FieldCol(Field)
End Function

Public Function Columns() As Integer
'Returns the number of columns.

    If HasError Then Exit Function
    Columns = UBound(Fld)
End Function

Private Function XData_Columns() As Integer
'Returns the number of columns.

    If HasError Then Exit Function
    XData_Columns = UBound(Fld)
End Function

Public Function Exists(Field As Variant) As Boolean

    Exists = XData_Exists(Field)
End Function

Private Function XData_Exists(Field As Variant) As Boolean
'Checks if a field exists.

    If HasError Then Exit Function
    Select Case VarType(Field)
    Case vbString
        XData_Exists = FieldList.Exists(Field)
    Case vbInteger, vbLong
        XData_Exists = Field >= 1 And Field <= UBound(Fld)
    End Select
End Function

Public Function GetColumnData(Field) As Boolean
'Creates in GroupData a column as a 2-d array suitable for placing on a worksheet.
'Uses a variant array in order to format data by type.
'String values are preceded by '.
    
    Dim I As Long, J As Long
    
    If HasError Then Exit Function
    If DataRowCount <= 0 Then
        GroupData = Xs.ExtractArray
        GetColumnData = True
        Exit Function
    End If
    If Not FieldCheck(Field) Then Exit Function
    With Fld(FieldCol(Field))
        ReDim GroupData(1 To DataRowCount, 1 To 1) As Variant
        For I = 1 To DataRowCount
            If .VarType = vbString Then
                GroupData(I, 1) = "'" & .Data(Pointer(I))
            Else
                GroupData(I, 1) = .Data(Pointer(I))
            End If
        Next I
    End With
    GetColumnData = True
End Function

Public Function GetArrayData(Field) As Boolean
'FOR INTERNAL USE ONLY - not recommended for end users to use directly.
'Copies Data array to Groupdata.
'If Field=0 copies Pointer to GroupData.
    
    If HasError Then Exit Function
    If CStr(Field) = "0" Then
        GroupData = Pointer
    ElseIf Not FieldCheck(Field) Then
        Exit Function
    Else
        GroupData = Fld(FieldCol(Field)).Data
    End If
    GetArrayData = True
End Function

Public Function Rename(OldName As String, NewName As String) As Boolean
'Renames a field.

    Dim I As Long

    If HasError Then Exit Function
    If Not FieldCheck(OldName) Then Exit Function
    If Exists(NewName) Then
        Msg NewName & " already exists"
        Exit Function
    End If
    
    I = FieldCol(OldName)
    Fld(I).Name = NewName
    FieldList.Key(OldName) = NewName
    
    Xs.Show Me
    Rename = True
End Function

Public Function CreateFrom(Xd As XData) As Boolean
'Create a new ArrData object from an ArrData or RsetData object.

    Dim Headers As Variant
    Dim DataTypes As Variant
    Dim Ad As ArrData
    Dim Rd As Object
    Dim I As Long

    If HasError Then Exit Function
    If Xd.HasError Then
        HasError = True
        Exit Function
    End If
    Select Case Xd.ObjType
    Case "ArrData"
        Set Ad = Xd
        ReDim Headers(0 To Ad.Columns - 1) As String
        ReDim DataTypes(0 To Ad.Columns - 1) As VbVarType
        For I = 0 To Ad.Columns - 1
            Headers(I) = Ad.GetHeader(I + 1)
            DataTypes(I) = Ad.GetType(I + 1)
        Next I
        Xs.NoShow = True
        If Not Create(Headers, DataTypes) Then Exit Function
        
        If Ad.DataRowCount = 0 Then
            CreateFrom = True
            Exit Function
        End If
        Xs.NoShow = True
        Ad.RecClose
        If Not Ad.GetArrayData(0) Then Exit Function
        Pointer = Ad.GroupData
        For I = 1 To Ad.Columns
            If Not Ad.GetArrayData(I) Then Exit Function
            Fld(I).Data = Ad.GroupData
        Next I
        DataRowCount = Ad.DataRowCount
        DataCount = UBound(Fld(1).Data)
        
    Case "RsetData"
        ReDim Headers(0 To Xd.Columns - 1) As String
        ReDim DataTypes(0 To Xd.Columns - 1) As VbVarType
        For I = 0 To Xd.Columns - 1
            Headers(I) = Xd.GetHeader(I + 1)
            DataTypes(I) = Xd.GetType(I + 1)
            If DataTypes(I) < 0 Then Exit Function
        Next I
        Xs.NoShow = True
        If Not Create(Headers, DataTypes) Then Exit Function
        Xs.NoShow = True
        If Not CopyFrom(Xd, 0, 0) Then Exit Function
        
    Case "TableData"
        Set Rd = Xs.GetRd
        Rd.CreateFrom Xd
        Xs.NoShow = True
        CreateFrom Rd
        If HasError Then Exit Function
    End Select
    RecStart
    
    Xs.Show Me
    CreateFrom = True
End Function

Function CopyFrom(Xd As XData, ByVal Start As Variant, _
    ByVal Recs As Variant, ParamArray Fields() As Variant) As Boolean
'Copy TableData, Rsetdata or ArrData object and append.
'If Start and Recs are both 0, assumes start at the first record and finises at the last record.
'If Start is <1, assumes starts at the first record.
'If Recs is negative, assumes finishes at the last record.
'If no list of fields supplied then all fields are copied if present on Xd2.
'If Fields = * then all fields are copied and must be present on Xd2.

    Dim Fields2 As Variant 'Fields as supplied
    Dim Fieldsx As Variant
    
    If HasError Then Exit Function
    RemoveIndex
    Fieldsx = Fields
    Fields2 = Xs.ExtractArray(Fieldsx)
    CopyFrom = Xs.Copy(Me, Xd, Start, Recs, Fields2)
    Xs.Show Me
End Function

Public Function CopyWhere(Xd As XData, Criteria As Variant, ParamArray Fields()) As Boolean
'Copy from TableData, RsetData or ArrData object and append.
'If no list of fields supplied (empty array) then all fields are copied if present on Xd.
'If Fields = * then all fields are copied and must be present on Xd.
'Criteria can be the name of a callback procedure or an array of Field/Value
'where Value can text be preceded by any of =, <, <=, >, >=, <>.
'Value of Empty ignores test for this field.
'Not as first item in criteria reverses comparison.
'All comparison fields must be on this Arrdata object.

    Dim Fields2 As Variant 'Fields as supplied
    Dim Fieldsx As Variant
    Dim Fields3 As Variant 'Fields columns
    Dim Fields4 As Variant 'Fields columns on Xd
    Dim Result As Boolean
    Dim Criteria2 As Variant
    Dim WhereFields() As Long
    Dim WhereValues() As Variant
    Dim WhereTypes() As TestType
    Dim HasNot As Boolean
    Dim I As Long
    
    If HasError Then Exit Function
    If Xd.HasError Then
        HasError = True
        Exit Function
    End If
    RemoveIndex
    Fieldsx = Fields
    Fields2 = Xs.ExtractArray(Fieldsx)
    If UBound(Fields2) < LBound(Fields2) Then
        Xs.MatchFields Me, Xd, Xs.ExtractArray, Fields3, , Fields4
    ElseIf Fields2(LBound(Fields2)) = "*" Then
        Xs.MatchFields Me, Nothing, Xs.ExtractArray, Fields3
        If Not Xs.CheckFields(Me, Xd, Fields3, Fields4) Then Exit Function
    Else
        If Not Xs.ProcessFields(Me, Fields2, Fields3) Then Exit Function
        If Not Xs.CheckFields(Me, Xd, Fields2, Fields4) Then Exit Function
    End If
    If UBound(Fields3) < LBound(Fields3) Then
        Msg "No fields on data"
        Exit Function
    End If
    
    Criteria2 = Xs.ExtractArray(Criteria)
    If Not ConvertCriteria(Criteria2, WhereFields, WhereValues, WhereTypes, HasNot) _
        Then Exit Function
    
    Xd.RecStart
    Do While Xd.RecNext
        RecAdd
        For I = 1 To UBound(Fields3)
            XData_RecValue2(Fields3(I)) = Xd.RecValue2(Fields4(I))
            If HasError Then Exit Function
            If Xd.HasError Then
                HasError = True
                Exit Function
            End If
        Next I
        
        If Not CompareCriteria(Criteria2, WhereFields, WhereValues, WhereTypes, HasNot, Result) _
            Then Exit Function
        If Not Result Then
            DataRowCount = DataRowCount - 1
            DataCount = DataCount - 1
            RecordPos = DataRowCount
        End If
    Loop
    Xs.NoShow = True
    RecClose
    
    Xs.Show Me
    CopyWhere = True
End Function

Public Function CopyCsv2(Filename As Variant, WithHeaders As Boolean, WithQuotes As Boolean, _
    SepChar As String, Terminator As Variant, ParamArray Fields()) As Boolean
'Reads data from a csv file.
'Also accepts an array of records in FileName.
'WithHeaders determines whether header is present.
'WithQuotes determines whether to look for text quotes and allow for commas in csv data
'in which case it converts "" to ".
'If Fields = * then all fields are copied and must be present on csv.
'If Fields is specified, skips records at the start of the file till headers are found.
'Stops if Terminator is encountered (changed to CopyCsv2 as incompatible with previous version)
'Terminator can be an array of possible values.
'Removes UTF8 byte order mark at start of file.

    Dim Fields2 As Variant 'Fields as supplied
    Dim Fieldsx As Variant
    Dim Fields3 As Variant 'Fields columns
    Dim Fields4 As Variant 'Fields columns on csv
    Dim Terminator2 As Variant
    Dim FileName2 As Variant
    Dim FileNum As Long
    Dim Record As String
    Dim SplitRec As Variant
    Dim VarTypes() As VbVarType
    Dim Ad As New ArrData
    Dim Value As Variant
    Dim IsEOF As Boolean
    Dim IsBOF As Boolean
    Dim RecNo As Long
    Dim I As Long, J As Long, K As Long
    
    If HasError Then Exit Function
    IsBOF = True
    
    If IsArray(Filename) Then
        RecNo = LBound(Filename) - 1
    ElseIf Len(Filename) > 0 Then
        FileName2 = Filename
    Else
        FileName2 = Application.GetOpenFilename("csv files (*.csv),*.csv", , "Select csv file to open")
        If FileName2 = False Then Exit Function
    End If
    
    RemoveIndex
    If Len(FileName2) > 0 Then
        FileNum = FreeFile()
        Open FileName2 For Input As #FileNum
    End If
    
    Terminator2 = Xs.ExtractArray(Terminator)
    Fieldsx = Fields
    Fields2 = Xs.ExtractArray(Fieldsx)
    If WithHeaders Then
        Do
            GoSub CheckEOF
            If IsEOF Then
                Msg "Csv header not found"
                GoTo Finish
            End If
            GoSub DoSplit
            ReDim VarTypes(UBound(SplitRec))
            For I = 0 To UBound(SplitRec)
                VarTypes(I) = vbInteger
                If Len(SplitRec(I)) = 0 Then
                    SplitRec(I) = "_Dummy" & I
                ElseIf Application.WorksheetFunction.Match(SplitRec(I), SplitRec, 0) <> I + 1 Then
                    SplitRec(I) = "_Dummy" & I
                End If
            Next I
            Ad.Xs.NoShow = True
            If Not Ad.Create(SplitRec, VarTypes) Then
                HasError = True
                GoTo Finish
            End If
            If UBound(Fields2) < LBound(Fields2) Then
                Xs.MatchFields Me, Ad, Xs.ExtractArray, Fields3, , Fields4
                Exit Do
            ElseIf Fields2(LBound(Fields2)) = "*" Then
                Xs.MatchFields Me, Nothing, Xs.ExtractArray, Fields3
                If Xs.Exists(Me, Ad, Fields3) Then
                    If Not Xs.CheckFields(Me, Ad, Fields3, Fields4) Then GoTo Finish
                    Exit Do
                End If
            Else
                If Not Xs.ProcessFields(Me, Fields2, Fields3) Then GoTo Finish
                If Xs.Exists(Me, Ad, Fields2) Then
                    If Not Xs.CheckFields(Me, Ad, Fields2, Fields4) Then GoTo Finish
                    Exit Do
                End If
            End If
        Loop
    Else
        If UBound(Fields2) < LBound(Fields2) Then
            Xs.MatchFields Me, Nothing, Xs.ExtractArray, Fields3
        ElseIf Fields2(LBound(Fields2)) = "*" Then
            Xs.MatchFields Me, Nothing, Xs.ExtractArray, Fields3
        Else
            If Not Xs.ProcessFields(Me, Fields2, Fields3) Then GoTo Finish
        End If
        Fields4 = Fields3
    End If
    If UBound(Fields3) < LBound(Fields3) Then
        Msg "No fields on data"
        GoTo Finish
    End If
    
    Do
        GoSub CheckEOF
        If IsEOF Then Exit Do
        GoSub DoSplit
        If UBound(SplitRec) > LBound(SplitRec) Or Len(SplitRec(LBound(SplitRec))) > 0 Then
            If UBound(Terminator2) >= LBound(Terminator2) Then
                For I = LBound(Terminator2) To UBound(Terminator2)
                    If Left(Record, Len(Terminator2(I))) = Terminator2(I) Then
                        Exit Do
                    End If
                Next I
            End If
            If UBound(SplitRec) < UBound(Fields3) - 1 Then
                Msg "Missing fields on csv file"
                GoTo Finish
            End If
            RecAdd
            For I = 1 To UBound(Fields3)
                J = Fields3(I)
                K = Fields4(I) - 1
                Value = SplitRec(K)
                If Not ConvertType(Value, Fld(J).VarType) Then GoTo Finish
                Fld(J).Data(Pointer(RecordPos)) = Value
            Next I
        End If
    Loop
    
    Xs.Show Me
    CopyCsv2 = True
    
Finish:
    If FileNum > 0 Then
        Close #FileNum
    End If
    Exit Function
    
CheckEOF:
    If FileNum > 0 Then
        IsEOF = EOF(FileNum)
    Else
        IsEOF = (RecNo >= UBound(Filename))
    End If
    Return
    
DoSplit:
    If FileNum > 0 Then
        Line Input #FileNum, Record
        If IsBOF Then
            If Left(Record, 3) = "﻿" Then
                Record = Mid(Record, 4)
            End If
            IsBOF = False
        End If
    Else
        RecNo = RecNo + 1
        Record = Filename(RecNo)
    End If
    If Len(Record) = 0 Then
        SplitRec = Array(vbNullString)
    ElseIf WithQuotes Then
        SplitRec = Split2(Record, SepChar)
    Else
        SplitRec = Split(Record, SepChar)
    End If
    Return
End Function

Public Function Pivot(PivotField, DataField) As Boolean
'Pivots by adding columns for the values of PivotField,
'and placing in those columns the values of DataField.
'Now adds fields in the order they are in the data.
'Must first have been grouped, so there is only one value of Datafield
'for each set of values of other fields.
'20/12/23 if pivot field is a date, formats new fields as yyyy-mm-dd

    Dim Fields2 As Variant 'PivotCol and DataCol as supplied
    Dim Fields3 As Variant 'PivotCol and DataCol columns
    Dim Fields6 As Variant 'Fields to group by columns
    Dim PivotCol As Long 'PivotField column
    Dim DataCol As Long 'DataField column
    Dim PivotName As String 'PivotField Name
    Dim DataName As String 'DataField Name
    Dim NewField As String
    Dim DataType As VbVarType
    Dim PivotType As VbVarType
    Dim Value As Variant
    Dim Items As Variant
    Dim Icol As Collection
    Dim Ival As Variant
    Dim Ival1 As Long
    Dim ItemCol As Long
    Dim I As Long, J As Long, K As Long

    If HasError Then Exit Function
    If DataRowCount <= 0 Then
        Pivot = True
        Exit Function
    End If
    RemoveIndex
    
    Fields2 = Array(PivotField, DataField)
    If Not Xs.ProcessFields(Me, Fields2, Fields3) Then Exit Function
    Xs.MatchFields Me, Nothing, Fields2, Fields6
    If UBound(Fields6) < LBound(Fields6) Then
        Msg "No fields to group by"
        Exit Function
    End If
    PivotCol = Fields3(1)
    DataCol = Fields3(2)
    PivotName = FieldName(PivotCol)
    DataName = FieldName(DataCol)
    DataType = Fld(DataCol).VarType
    PivotType = Fld(PivotCol).VarType
    
    For I = 1 To DataRowCount
        If PivotType = vbDate Then
            NewField = Format(Fld(PivotCol).Data(Pointer(I)), "yyyy-mm-dd")
        Else
            NewField = Fld(PivotCol).Data(Pointer(I))
        End If
        If Not Exists(NewField) Then
            Xs.NoShow = True
            If Not AddField(NewField, DataType) Then GoTo Finish
        End If
    Next I
    
    If Not CreateIndex(Fields6) Then Exit Function
    Items = RowIndex.Items
    For I = LBound(Items) To UBound(Items)
        Set Icol = Items(I)
        J = 0
        Ival1 = Icol.Item(1)
        For Each Ival In Icol
            J = J + 1
            If PivotType = vbDate Then
                NewField = Format(Fld(PivotCol).Data(Pointer(Ival)), "yyyy-mm-dd")
            Else
                NewField = Fld(PivotCol).Data(Pointer(Ival))
            End If
            Value = Fld(DataCol).Data(Pointer(Ival))
            ItemCol = FieldCol(NewField)
            If Not CheckNull(Fld(ItemCol).Data(Pointer(Ival1))) Then
                Msg "Duplicate value for " & NewField & _
                    " on row " & Ival
                GoTo Finish
            End If
            Fld(ItemCol).Data(Pointer(Ival1)) = Value
            If J > 1 Then Pointer(Ival) = 0
        Next Ival
    Next I
    
    Xs.NoShow = True
    If Not DropField(PivotName, DataName) Then GoTo Finish
    CompressPointers
    
    Xs.Show Me
    Pivot = True

Finish:
    RemoveIndex
End Function

Public Function UnPivot(Xd As XData, PivotField, DataField, ParamArray JoinFields()) As Boolean
'Unpivots a Xdata object.
'Places field name in PivotField and field value in DataField
'Transfers JoinFields on each row.

    Dim PivotCol As Long 'PivotField column
    Dim DataCol As Long 'DataField column
    Dim Fields2 As Variant 'JoinFields as supplied
    Dim Fieldsx As Variant
    Dim Fields3 As Variant 'JoinFields columns
    Dim Fields4 As Variant 'JoinFields names
    Dim Fields5 As Variant 'JoinFields columns in Xd
    Dim Fields6 As Variant 'Remaining fields in Xd names
    Dim Fields7 As Variant 'Remaining fields in Xd columns in Xd
    Dim Value As Variant
    Dim I As Long, J As Long, K As Long
    
    If HasError Then Exit Function
    If Xd.HasError Then
        HasError = True
        Exit Function
    End If
    If Xd.DataRowCount = 0 Then
        UnPivot = True
        Exit Function
    End If
    RemoveIndex
    
    If Not FieldCheck(PivotField) Then Exit Function
    PivotCol = FieldCol(PivotField)
    If Not FieldCheck(DataField) Then Exit Function
    DataCol = FieldCol(DataField)
    
    Fieldsx = JoinFields
    Fields2 = Xs.ExtractArray(Fieldsx)
    If UBound(Fields2) < LBound(Fields2) Then
        Msg "No join fields"
        Exit Function
    End If
    If Not Xs.ProcessFields(Me, Fields2, Fields3, Fields4) Then Exit Function
    If Not Xs.CheckFields(Me, Xd, Fields2, Fields5) Then Exit Function
    If Not Xs.CheckDuplicates(Me, Array(PivotField, DataField), Fields2) Then Exit Function
    
    ReDim Fields6(1 To Xd.Columns)
    ReDim Fields7(1 To Xd.Columns)
    For I = 1 To Xd.Columns
        For J = 1 To UBound(Fields4)
            If Xd.GetHeader(I) = Fields4(J) Then Exit For
        Next J
        If J > UBound(Fields4) Then
            K = K + 1
            Fields6(K) = Xd.GetHeader(I)
            Fields7(K) = I
        End If
    Next I
    If K = 0 Then
        Msg "No fields to unpivot"
        Exit Function
    End If
    ReDim Preserve Fields6(1 To K)
    ReDim Preserve Fields7(1 To K)
    
    Xd.RecStart
    Do While Xd.RecNext
        For I = 1 To UBound(Fields6)
            Value = Xd.RecValue2(Fields7(I))
            If Not CheckNull(Value) Then
                RecAdd
                For J = 1 To UBound(Fields3)
                    XData_RecValue2(Fields3(J)) = Xd.RecValue2(Fields5(J))
                Next J
                XData_RecValue2(PivotCol) = Fields6(I)
                XData_RecValue2(DataCol) = Value
                If HasError Then Exit Function
            End If
        Next I
    Loop
    RecStart
    
    Xs.Show Me
    UnPivot = True
End Function

Public Function Filter(ParamArray Criteria()) As Boolean
'Filter by criteria.
'Criteria can be the name of a callback procedure or an array of Field/Value
'where Value can text be preceded by any of =, <, <=, >, >=, <>.
'Value of Empty ignores test for this field.
'Not as first item in criteria reverses comparison.

    Dim Result As Boolean
    Dim Criteria2 As Variant
    Dim Criteriax As Variant
    Dim WhereFields() As Long
    Dim WhereValues() As Variant
    Dim WhereTypes() As TestType
    Dim HasNot As Boolean
    Dim I As Long, J As Long
    
    If HasError Then Exit Function
    If DataRowCount <= 0 Then
        Filter = True
        Exit Function
    End If
    
    RemoveIndex
    Criteriax = Criteria
    Criteria2 = Xs.ExtractArray(Criteriax)
    If Not ConvertCriteria(Criteria2, WhereFields, WhereValues, WhereTypes, HasNot) _
        Then Exit Function
    
    RecStart
    Do While RecNext
        If Not CompareCriteria(Criteria2, WhereFields, WhereValues, WhereTypes, HasNot, Result) _
            Then Exit Function
        If Not Result Then
            Pointer(RecordPos) = 0
        End If
    Loop
    
    CompressPointers
    Xs.Show Me
    Filter = True
End Function

Public Function Calc(Field As Variant, ByVal CalcProc As Variant, Optional Scratch As Range, _
    Optional Params As Variant) As Boolean
'Calculates a field using callback or formula.
'Formula is an array of formula elements where field names are separate elements or a macro name
'(also accepts formula starting with = where field names are delimited by `).
'\ before a field name indicates use value from previous record.
'If calculation result starts with # and ends with !, it is treated as an error and the contents are the message.
'If Scratch is provided, it is used for calculation, otherwise formula values are evaluated individually.

    Dim Result As Variant
    Dim ResultSet As Variant
    Dim CalcFields As Variant 'Fields in CalcProc
    Dim CalcOffset As Variant 'True to use previous value
    Dim Params2 As Variant
    Dim Value As Variant
    Dim FormText As String
    Dim I As Long
    
    If HasError Then Exit Function
    Set Xs = New XShared 'To initialise if necessary
    If DataRowCount <= 0 Then
        Calc = True
        Exit Function
    End If
    
    If Not FieldCheck(Field) Then Exit Function
    If Fld(FieldCol(Field)).Indexed Then RemoveIndex
    
    If Not IsArray(CalcProc) Then
        If Left(CalcProc, 1) = "=" Then
            CalcProc = Split(Mid(CalcProc, 2), "`")
        End If
    End If
    
    If IsArray(CalcProc) Then
        ReDim CalcFields(LBound(CalcProc) To UBound(CalcProc)) As Long
        ReDim CalcOffset(LBound(CalcProc) To UBound(CalcProc)) As Boolean
        For I = LBound(CalcProc) To UBound(CalcProc)
            If Left(CalcProc(I), 1) = "\" Then
                CalcProc(I) = Mid(CalcProc(I), 2)
                CalcOffset(I) = True
            End If
            If Exists(CalcProc(I)) And (Len(CalcProc(I)) > 0 Or (I > LBound(CalcProc) And I < UBound(CalcProc))) Then
                CalcFields(I) = FieldCol(CalcProc(I))
            End If
        Next I
    End If
    Params2 = Xs.ExtractArray(Params)
    
    I = FieldCol(Field)
    
    If IsArray(CalcProc) And Not Scratch Is Nothing Then
        RecStart
        ReDim GroupData(1 To DataRowCount, 1 To 1) As Variant
        Do While RecNext
            GroupData(RecordPos, 1) = "=" & GetFormText(CalcProc, CalcFields, CalcOffset)
        Loop
        With Scratch.Resize(DataRowCount)
            On Error Resume Next
            .Formula = GroupData
            If Err = 0 Then
                On Error GoTo 0
                ResultSet = .Value
                If Not IsArray(ResultSet) Then
                    Value = ResultSet
                    ReDim ResultSet(1 To 1, 1 To 1) As Variant
                    ResultSet(1, 1) = Value
                End If
                .Clear
            Else
                On Error GoTo 0
                Msg "Error in formula carrying out calculation" & vbLf & _
                    "Please see VBA Immediate Pane for more information"
                Debug.Print "First formula =" & GroupData(1, 1)
                .Clear
                GoTo Finish
            End If
        End With
        RecStart
        Do While RecNext
            Result = ResultSet(RecordPos, 1)
            FormText = GroupData(RecordPos, 1)
            If Not CheckResult(Result, FormText) Then GoTo Finish
            XData_RecValue2(I) = Result
            If HasError Then GoTo Finish
        Loop
    Else
        RecStart
        Do While RecNext
            If Not DoCalculate(CalcProc, Result, CalcFields, CalcOffset, Xs, Params2) Then Exit Function
            XData_RecValue2(I) = Result
            If HasError Then Exit Function
        Loop
    End If
    
    Xs.Show Me
    Calc = True
    
Finish:
    GroupData = Empty
End Function

Public Function Formula(Field As Variant, ByVal CalcProc As Variant, _
    Optional Params As Variant) As Boolean
'Saves a callback or formula for future calculation.
'Formula is an array of formula elements where field names are separate elements or a macro name
'(also accepts formula starting with = where field names are delimited by `).
'\ before a field name indicates use value from previous record.
'If calculation result starts with # and ends with !, it is treated as an error and the contents are the message.

    Dim I As Long
        
    If HasError Then Exit Function
    If Not FieldCheck(Field) Then Exit Function
    With Fld(FieldCol(Field))
    
        If Not IsArray(CalcProc) Then
            If Left(CalcProc, 1) = "=" Then
                CalcProc = Split(Mid(CalcProc, 2), "`")
            End If
        End If
        
        If IsArray(CalcProc) Then
            ReDim .CalcFields(LBound(CalcProc) To UBound(CalcProc)) As Long
            ReDim .CalcOffset(LBound(CalcProc) To UBound(CalcProc)) As Boolean
            For I = LBound(CalcProc) To UBound(CalcProc)
                If Left(CalcProc(I), 1) = "\" Then
                    CalcProc(I) = Mid(CalcProc(I), 2)
                    .CalcOffset(I) = True
                End If
                If Exists(CalcProc(I)) Then
                    .CalcFields(I) = FieldCol(CalcProc(I))
                End If
            Next I
        End If
        .CalcProc = CalcProc
        .Params = Xs.ExtractArray(Params)
    End With
    Formula = True
End Function

Public Function CalcAll(Optional Scratch As Range) As Boolean
'Calculates by rows using previously saved formulae.
'Formula values are evaluated individually.
'Formulae are then removed.
'If Scratch is provided, it is used for calculation.
'If calculation result starts with # and ends with !, it is treated as an error and the contents are the message.

    Dim Result As Variant
    Dim Value As Variant
    Dim Field2() As Long 'Fields with formulae
    Dim MyXs() As New XShared
    Dim FormText As String
    Dim I As Long, J As Long
    
    If HasError Then Exit Function
    ReDim MyXs(1 To UBound(Fld))
    If DataRowCount <= 0 Then
        CalcAll = True
        Exit Function
    End If
    
    ReDim Field2(1 To UBound(Fld))
    For I = 1 To UBound(Fld)
        If Not IsEmpty(Fld(I).CalcProc) Then
            J = J + 1
            Field2(J) = I
            If Fld(I).Indexed Then RemoveIndex
        End If
    Next I
    If J = 0 Then
        CalcAll = True
        Exit Function
    Else
        ReDim Preserve Field2(1 To J)
    End If
        
    RecStart
    Do While RecNext
        For J = 1 To UBound(Field2)
            I = Field2(J)
            With Fld(I)
                If IsArray(.CalcProc) And Not Scratch Is Nothing Then
                    FormText = GetFormText(.CalcProc, .CalcFields, .CalcOffset)
                    On Error Resume Next
                    Scratch.Cells(1).Formula = "=" & FormText
                    If Err = 0 Then
                        On Error GoTo 0
                        Result = Scratch.Cells(1).Value
                    Else
                        On Error GoTo 0
                        Msg "Error in formula carrying out calculation" & vbLf & _
                            "Please see VBA Immediate Pane for more information"
                        Debug.Print "Record Number " & RecordPos
                        Debug.Print "Formula =" & FormText
                        GoTo Finish
                    End If
                    If Not CheckResult(Result, FormText) Then GoTo Finish
                Else
                    If Not DoCalculate(.CalcProc, Result, .CalcFields, .CalcOffset, MyXs(I), .Params) Then Exit Function
                End If
                XData_RecValue2(I) = Result
                If HasError Then Exit Function
            End With
        Next J
    Loop
    
    If Not Scratch Is Nothing Then Scratch.Clear
        
    For I = 1 To UBound(Fld)
        With Fld(I)
            .CalcFields = Empty
            .CalcOffset = Empty
            .CalcProc = Empty
        End With
    Next I
    
    Xs.Show Me
    CalcAll = True
    
Finish:
    GroupData = Empty
End Function

Public Function GetOffset(ByVal Field1 As Variant, ByVal Field2 As Variant, _
    Offset As Long, Optional ByVal Default As Variant = Empty) As Boolean
'Places in Field1 the value in Field2 offset by Offset, or Default if out of range.

    Dim Result As Variant
    Dim I As Long
    
    If HasError Then Exit Function
    If Not FieldCheck(Field1) Then Exit Function
    If Not FieldCheck(Field2) Then Exit Function
    Field1 = FieldCol(Field1)
    Field2 = FieldCol(Field2)
    If Not ConvertType(Default, Fld(Field1).VarType) Then Exit Function
    
    RecStart
    Do While RecNext
        I = RecordPos + Offset
        Select Case I
        Case Is < 1
            Result = Default
        Case Is > DataRowCount
            Result = Default
        Case Else
            Result = Fld(Field2).Data(Pointer(I))
        End Select
        XData_RecValue2(Field1) = Result
        If HasError Then Exit Function
    Loop
    
    Xs.Show Me
    GetOffset = True
    End Function

Public Function RecStart() As Boolean

    RecStart = XData_RecStart
End Function

Private Function XData_RecStart() As Boolean
'Moves to BOF position while reading data in record mode and returns True, or False if no records.
'Also can be used to initialise record adding if records had previously been read.
    
    If HasError Then Exit Function
    Xs.NoShow = True
    RecClose
    RecordPos = 0
    If DataRowCount > 0 Then XData_RecStart = True
End Function

Public Function RecNext() As Boolean

    RecNext = XData_RecNext
End Function

Private Function XData_RecNext() As Boolean
'Moves to next position while reading data in record mode and returns True, or False if at EOF.

    If HasError Then Exit Function
    RecordPos = RecordPos + 1
    If RecordPos > DataRowCount Then
        XData_RecNext = False
    Else
        XData_RecNext = True
    End If
End Function

Public Property Get RecPos() As Long

    If HasError Then Exit Property
    RecPos = RecordPos
End Property

Private Property Get XData_RecPos() As Long
'Returns record position while reading data.

    If HasError Then Exit Property
    XData_RecPos = RecordPos
End Property

Public Property Let RecPos(Pos As Long)

    XData_RecPos = Pos
End Property

Private Property Let XData_RecPos(Pos As Long)
'Sets record position while reading data.

    If HasError Then Exit Property
    Select Case Pos
    Case Is > DataRowCount
        Msg "Attempting to set record position beyond end of table"
    Case Is <= 0
        RecStart
    Case Else
        RecordPos = Pos
    End Select
End Property

Public Function RecAdd() As Boolean

    RecAdd = XData_RecAdd
End Function

Private Function XData_RecAdd() As Boolean
'Adds a record and positions at that record.

    Dim I As Long
    
    If HasError Then Exit Function
    RemoveIndex
    DataRowCount = DataRowCount + 1
    DataCount = DataCount + 1
    If DataCount > UBound(Fld(1).Data) Then
        For I = 1 To UBound(Fld)
            ReDim Preserve Fld(I).Data(1 To UBound(Fld(I).Data) + GranSize)
        Next I
    End If
    If DataRowCount > UBound(Pointer) Then
        ReDim Preserve Pointer(1 To UBound(Pointer) + GranSize)
    End If
    Pointer(DataRowCount) = DataCount
    RecordPos = DataRowCount
    XData_RecAdd = True
End Function

Public Function RecDelete() As Boolean

    RecDelete = XData_RecDelete
End Function

Private Function XData_RecDelete() As Boolean
'Deletes the current record.
'Leaves RecPos poritioned at previous record (or 0 if none)

    Dim I As Long

    If HasError Then Exit Function
    If RecordPos <= 0 Or RecordPos > DataRowCount Then
        Msg "No record to delete"
        Exit Function
    End If
    RemoveIndex
    For I = RecordPos To DataRowCount - 1
        Pointer(I) = Pointer(I + 1)
    Next I
    DataRowCount = DataRowCount - 1
    If DataRowCount > 0 Then
        ReDim Preserve Pointer(1 To DataRowCount)
    Else
        ReDim Pointer(1 To GranSize)
    End If
    RecordPos = RecordPos - 1
    XData_RecDelete = True
End Function

Public Function RecClose(Optional Cancel As Boolean = False) As Boolean
'Cancel is not implemented.

    RecClose = XData_RecClose(Cancel)
End Function

Private Function XData_RecClose(Optional Cancel As Boolean = False) As Boolean
'Resize arrays when all data has been added.
'Cancel is not implemented.

    Dim I As Long
    
    If HasError Then Exit Function
    For I = 1 To UBound(Fld)
        If DataCount > 0 Then
            ReDim Preserve Fld(I).Data(1 To DataCount)
        Else
            DimData I, GranSize
        End If
    Next I
    If DataRowCount > 0 Then
        ReDim Preserve Pointer(1 To DataRowCount)
    Else
        ReDim Pointer(1 To GranSize)
    End If
    
    Xs.Show Me
    XData_RecClose = True
End Function

Public Property Get RecValue(Field) As Variant
Attribute RecValue.VB_UserMemId = 0
'Returns data from current record while reading data.

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        RecValue = Fld(FieldCol(Field)).Data(Pointer(RecordPos))
    End If
End Property

Public Property Get XData_RecValue(Field) As Variant
'Returns data from current record while reading data.

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        XData_RecValue = Fld(FieldCol(Field)).Data(Pointer(RecordPos))
    End If
End Property

Private Property Get XData_RecValue2(Field As Variant) As Variant
'Returns data from current record while reading data.
'Assumes Field is a valid column number.

    If HasError Then Exit Property
    XData_RecValue2 = Fld(Field).Data(Pointer(RecordPos))
End Property

Public Property Let RecValue(Field, FieldValue As Variant)
'Sets value in current record while writing data.

    Dim Field2 As Long

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Field2 = FieldCol(Field)
        If Fld(Field2).Indexed Then RemoveIndex
        XData_RecValue2(Field2) = FieldValue
    End If
End Property

Public Property Let XData_RecValue(Field, FieldValue As Variant)
'Sets value in current record while writing data.

    Dim Field2 As Long

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Field2 = FieldCol(Field)
        If Fld(Field2).Indexed Then RemoveIndex
        XData_RecValue2(Field2) = FieldValue
    End If
End Property

Private Property Let XData_RecValue2(Field As Variant, FieldValue As Variant)
'Sets value in current record while writing data.
'Assumes Field is a valid column number.

    If HasError Then Exit Property
    With Fld(Field)
        On Error Resume Next
        If VarType(FieldValue) = vbString Then
            If Len(FieldValue) = 0 Then
                .Data(Pointer(RecordPos)) = Empty
            Else
                .Data(Pointer(RecordPos)) = FieldValue
            End If
        Else
            .Data(Pointer(RecordPos)) = FieldValue
        End If
        If Err = 0 Then
            On Error GoTo 0
        Else
            On Error GoTo 0
            Msg "Bad data " & CStr(FieldValue) & " on field " & FieldName(Field)
        End If
    End With
End Property

Public Function CreateIndex(ParamArray Fields() As Variant) As Boolean

    Dim Fields2 As Variant
    Dim Fieldsx As Variant
    
    Fieldsx = Fields
    Fields2 = Xs.ExtractArray(Fieldsx)
    CreateIndex = CreateIndex2(Fields2)
End Function

Private Function XData_CreateIndex(ParamArray Fields() As Variant) As Boolean

    Dim Fields2 As Variant
    Dim Fieldsx As Variant
    
    Fieldsx = Fields
    Fields2 = Xs.ExtractArray(Fieldsx)
    XData_CreateIndex = CreateIndex2(Fields2)
End Function

Private Function CreateIndex2(Fields2 As Variant) As Boolean
'Creates an index as a dictionary object.
'Keys are the specified fields separated by SepChar
'Items are a collection of record numbers.
'Indexes empty records (unlike TableData).

    Dim Fields3 As Variant 'index field numbers
    Dim Fields4 As Variant 'index field names
    Dim Ikey As String
    Dim Icol As Collection
    Dim ErrMsg As String
    Dim I As Long, J As Long
    
    If HasError Then Exit Function
    RemoveIndex
    Set RowIndex = New Dictionary
    RowIndex.CompareMode = TextCompare
    If DataRowCount <= 0 Then
        CreateIndex2 = True
        Exit Function
    End If
    
    If UBound(Fields2) < LBound(Fields2) Then
        Msg "No fields"
        Exit Function
    End If
    If Not Xs.ProcessFields(Me, Fields2, Fields3, Fields4) Then Exit Function
    For I = 1 To UBound(Fields3)
        Fld(Fields3(I)).Indexed = True
    Next I
    
    For I = 1 To DataRowCount
        Ikey = vbNullString
        For J = 1 To UBound(Fields3)
            If J > 1 Then
                Ikey = Ikey & (ChrW(SepChar) & CStr(Fld(Fields3(J)).Data(Pointer(I))))
            Else
                Ikey = CStr(Fld(Fields3(J)).Data(Pointer(I)))
            End If
        Next J
        If RowIndex.Exists(Ikey) Then
            Select Case IndexType
            Case IndexMultiple
                Set Icol = RowIndex.Item(Ikey)
                Icol.Add I
                Set RowIndex.Item(Ikey) = Icol
            Case IndexUnique
                ErrMsg = "Duplicate value on record " & I
                If Len(Ikey) > 0 Then
                    For J = 1 To UBound(Fields3)
                        ErrMsg = ErrMsg & vbLf & "field " & Fields4(J) & " value " & Split(Ikey, ChrW(SepChar))(J - 1)
                    Next J
                End If
                Msg ErrMsg
                Exit Function
            Case Else
                Msg "Index type not recognised " & IndexType
                Exit Function
            End Select
        Else
            Set Icol = New Collection
            Icol.Add I
            RowIndex.Add Ikey, Icol
        End If
    Next I
    
    CreateIndex2 = True
End Function

Public Function RemoveIndex() As Boolean

    RemoveIndex = XData_RemoveIndex()
End Function

Private Function XData_RemoveIndex() As Boolean

    Dim I As Long
    
    If HasError Then Exit Function
    If RowIndex Is Nothing Then Exit Function
    Set RowIndex = Nothing
    For I = 1 To UBound(Fld)
        Fld(I).Indexed = False
    Next I
    XData_RemoveIndex = True
End Function

Public Function Lookup(ParamArray Values() As Variant) As Variant

    Dim Values2 As Variant
    Dim Valuesx As Variant
    
    Valuesx = Values
    Values2 = Xs.ExtractArray(Valuesx)
    Lookup = Lookup2(Values2)
End Function

Private Function XData_Lookup(ParamArray Values() As Variant) As Variant

    Dim Values2 As Variant
    Dim Valuesx As Variant
    
    Valuesx = Values
    Values2 = Xs.ExtractArray(Valuesx)
    XData_Lookup = Lookup2(Values2)
End Function

Private Function Lookup2(Values2 As Variant) As Variant
'Returns a zero-based array of row numbers from index.
'Returns an empty array if there is no index or the key is not found.

    Dim Ikey As String
    Dim Icol As Collection
    Dim Ival As Variant
    Dim Result As Variant
    Dim I As Long
    
    Lookup2 = Xs.ExtractArray
    If HasError Then Exit Function
    If RowIndex Is Nothing Then Exit Function
    For I = LBound(Values2) To UBound(Values2)
        If I > LBound(Values2) Then
            Ikey = Ikey & (ChrW(SepChar) & CStr(Values2(I)))
        Else
            Ikey = CStr(Values2(I))
        End If
    Next I
    If Not RowIndex.Exists(Ikey) Then Exit Function
    Set Icol = RowIndex.Item(Ikey)
    ReDim Result(0 To Icol.Count - 1) As Long
    I = -1
    For Each Ival In Icol
        I = I + 1
        Result(I) = Ival
    Next Ival
    Lookup2 = Result
End Function

Public Function GetIndex() As Variant
'Returns entire index as a 2-dimensional zero-based array, entries x fields (0 based).
'Returns an empty array if no index or no data.

    Dim Keys As Variant
    Dim KeyCount As Long
    Dim Values As String
    Dim SplitValues As Variant
    Dim Result As Variant
    Dim I As Long, J As Long

    GetIndex = Xs.ExtractArray
    If HasError Then Exit Function
    If RowIndex Is Nothing Then Exit Function
    If RowIndex.Count = 0 Then Exit Function
    Keys = RowIndex.Keys
    For I = 1 To UBound(Fld)
        If Fld(I).Indexed Then KeyCount = KeyCount + 1
    Next I
    ReDim Result(UBound(Keys), KeyCount - 1)
    For I = 0 To UBound(Keys)
        Values = Keys(I)
        If Len(Values) = 0 Then
            SplitValues = Array(vbNullString)
        Else
            SplitValues = Split(Values, ChrW(SepChar))
        End If
        For J = 0 To KeyCount - 1
            Result(I, J) = SplitValues(J)
        Next J
    Next I
    GetIndex = Result
End Function

Public Function LeftJoin(Xd As XData, JoinFields As Variant, ParamArray CopyFields()) As Boolean
'Left join the FIRST matching row on TableData or ArrData object on equal JoinFields.
'Brings in CopyFields (or all matching fields if not specified).

    Dim Fields2 As Variant 'JoinFields as supplied
    Dim Fields3 As Variant 'JoinFields columns
    Dim Fields4 As Variant 'JoinFields names
    Dim Fields5 As Variant 'CopyFields as supplied
    Dim Fieldsx As Variant
    Dim Fields6 As Variant 'CopyFields columns
    Dim Fields7 As Variant 'CopyFields names
    Dim Fields8 As Variant 'CopyFields columns in Xd
    Dim IFields As Variant
    Dim Lookup As Variant
    Dim I As Long, J As Long, K As Long
    
    If HasError Then Exit Function
    If Xd.HasError Then
        HasError = True
        Exit Function
    End If
    If DataRowCount <= 0 Then
        LeftJoin = True
        Exit Function
    End If

    Fields2 = Xs.ExtractArray(JoinFields)
    If UBound(Fields2) < LBound(Fields2) Then
        Msg "No join fields"
        Exit Function
    End If
    If Not Xs.ProcessFields(Me, Fields2, Fields3, Fields4) Then Exit Function
    If Not Xs.CheckFields(Me, Xd, Fields2) Then Exit Function
    
    Fieldsx = CopyFields
    Fields5 = Xs.ExtractArray(Fieldsx)
    If UBound(Fields5) < LBound(Fields5) Then
        Xs.MatchFields Me, Xd, Fields2, Fields6, Fields7, Fields8
    Else
        If Not Xs.ProcessFields(Me, Fields5, Fields6, Fields7) Then Exit Function
        If Not Xs.CheckFields(Me, Xd, Fields5, Fields8) Then Exit Function
        If Not Xs.CheckDuplicates(Me, Fields2, Fields5) Then Exit Function
    End If
    
    For I = 1 To UBound(Fields6)
        If Fld(Fields6(I)).Indexed Then RemoveIndex
        If Not DimData(CLng(Fields6(I)), DataCount) Then Exit Function
    Next I
    
    If Not Xd.CreateIndex(Fields4) Then Exit Function
    ReDim IFields(1 To UBound(Fields3))
    
    RecStart
    Do While RecNext
        For J = 1 To UBound(Fields3)
            IFields(J) = CStr(Fld(Fields3(J)).Data(Pointer(RecordPos)))
        Next J
        Lookup = Xd.Lookup(IFields)
        If UBound(Lookup) >= 0 Then
            Xd.RecPos = CLng(Lookup(0))
            For J = 1 To UBound(Fields6)
                XData_RecValue2(Fields6(J)) = Xd.RecValue2(Fields8(J))
                If HasError Then Exit Function
            Next J
        End If
    Loop
    
    Xs.Show Me
    LeftJoin = True
    
Finish:
    Xd.RemoveIndex
End Function

Public Function Match(Ad As ArrData, MatchMethod As MatchType, MatchField As Variant, CopyField As Variant, _
    ByVal Default As Variant, ParamArray JoinFields() As Variant) As Boolean
'Matches JoinFields, then searches all the matching records for MatchField.
'Brings back to CopyField either the matching value or the next highest or lowest
'(depending on MatchMethod) or Default if none.
'Ad must be sorted in ascending order on JoinFields and MatchField.

    Dim Xd As XData
    Dim Fields2 As Variant 'JoinFields as supplied
    Dim Fieldsx As Variant
    Dim Fields3 As Variant 'JoinFields columns
    Dim Fields4 As Variant 'JoinFields names
    Dim Fields5 As Variant 'MatchField/CopyField as supplied
    Dim Fields6 As Variant 'MatchField/CopyField columns
    Dim Field8 As Variant 'MatchField column in Ad
    Dim IFields As Variant
    Dim Lookup As Variant
    Dim Value As Variant
    Dim Lo As Long
    Dim Hi As Long
    Dim Md As Long
    Dim I As Long, J As Long, K As Long
    
    If HasError Then Exit Function
    If Ad.HasError Then
        HasError = True
        Exit Function
    End If
    If DataRowCount <= 0 Then
        Match = True
        Exit Function
    End If
    
    Set Xd = Ad
    Fieldsx = JoinFields
    Fields2 = Xs.ExtractArray(Fieldsx)
    If UBound(Fields2) < LBound(Fields2) Then
        If Ad.DataRowCount = 0 Then
            Lookup = Xs.ExtractArray
        Else
            ReDim Lookup(0 To Ad.DataRowCount - 1) As Long
            For I = 0 To Ad.DataRowCount - 1
                Lookup(I) = I + 1
            Next I
        End If
    End If
    If Not Xs.ProcessFields(Me, Fields2, Fields3, Fields4) Then Exit Function
    If Not Xs.CheckFields(Me, Xd, Fields2) Then Exit Function

    Fields5 = Array(MatchField, CopyField)
    If Not Xs.ProcessFields(Me, Fields5, Fields6) Then Exit Function
    If Not Xs.CheckFields(Me, Xd, Array(Fields5(0)), Field8) Then Exit Function
    If Not Xs.CheckDuplicates(Me, Fields2, Fields5) Then Exit Function
    Field8 = Field8(1)
    
    If Not ConvertType(Default, Fld(Fields6(2)).VarType) Then Exit Function
    
    If Fld(Fields6(2)).Indexed Then RemoveIndex
    If Not DimData(CLng(Fields6(2)), DataCount) Then Exit Function
    
    If UBound(Fields2) >= LBound(Fields2) Then
        If Not Ad.CreateIndex(Fields4) Then Exit Function
        ReDim IFields(1 To UBound(Fields3))
    End If
    
    RecStart
    Do While RecNext
        If UBound(Fields2) >= LBound(Fields2) Then
            For J = 1 To UBound(Fields3)
                IFields(J) = CStr(Fld(Fields3(J)).Data(Pointer(RecordPos)))
            Next J
            Lookup = Ad.Lookup(IFields)
        End If
        If UBound(Lookup) < 0 Then
            XData_RecValue2(Fields6(2)) = Default
            If HasError Then Exit Function
            
        Else
            Value = Fld(Fields6(1)).Data(Pointer(RecordPos))
            If Not ConvertType(Value, Ad.GetType(Field8)) Then Exit Function
            
            Lo = 0
            Hi = UBound(Lookup)
            Do While Lo <= Hi
                Md = Lo + Int((Hi - Lo) / 2)
                Xd.RecPos = CLng(Lookup(Md))
                Select Case Xd.RecValue2(Field8)
                Case Value
                    I = Md
                    Exit Do
                Case Is < Value
                    Lo = Md + 1
                Case Else
                    Hi = Md - 1
                End Select
            Loop
            If Lo > Hi Then
                Select Case MatchMethod
                Case MatchMinGE
                    I = Lo
                Case MatchMaxLE
                    I = Hi
                Case MatchEqual
                    I = -1
                Case Else
                    Msg "Match method " & MatchMethod & " not recognised"
                    Exit Function
                End Select
            End If
        
            If I < 0 Or I > UBound(Lookup) Then
                XData_RecValue2(Fields6(2)) = Default
            Else
                Xd.RecPos = CLng(Lookup(I))
                XData_RecValue2(Fields6(2)) = Xd.RecValue2(Field8)
            End If
            If HasError Then Exit Function
        End If
    Loop
    
    Xs.Show Me
    Match = True
    
Finish:
    Ad.RemoveIndex
End Function

Public Function CrossProduct(Xd As XData, ParamArray CopyFields()) As Boolean
'Create a cross product of the existing table with TableData or ArrData object.
'Brings in CopyFields (or all matching fields if not specified).

    Dim Fields5 As Variant 'CopyFields as supplied
    Dim Fieldsx As Variant
    Dim Fields6 As Variant 'CopyFields columns
    Dim Fields7 As Variant 'CopyFields names
    Dim Fields8 As Variant 'CopyFields columns in Xd
    Dim OldRowCount As Long
    Dim I As Long, J As Long, K As Long
    
    If HasError Then Exit Function
    If Xd.HasError Then
        HasError = True
        Exit Function
    End If
    If DataRowCount <= 0 Then
        CrossProduct = True
        Exit Function
    End If
    
    RemoveIndex
    Fieldsx = CopyFields
    Fields5 = Xs.ExtractArray(Fieldsx)
    If UBound(Fields5) < LBound(Fields5) Then
        Xs.MatchFields Me, Xd, Xs.ExtractArray, Fields6, Fields7, Fields8
    Else
        If Not Xs.ProcessFields(Me, Fields5, Fields6, Fields7) Then Exit Function
        If Not Xs.CheckFields(Me, Xd, Fields5, Fields8) Then Exit Function
    End If
    
    For I = LBound(Fields6) To UBound(Fields6)
        If Not DimData(CLng(Fields6(I)), DataCount) Then Exit Function
    Next I
    
    OldRowCount = DataRowCount
    For I = 2 To Xd.DataRowCount
        For J = 1 To OldRowCount
            RecAdd
            For K = 1 To UBound(Fld)
                With Fld(K)
                    .Data(Pointer(RecordPos)) = .Data(Pointer(J))
                End With
            Next K
        Next J
    Next I
    
    RecStart
    Xd.RecStart
    Do While Xd.RecNext
        For J = 1 To OldRowCount
            RecNext
            For K = 1 To UBound(Fields6)
                XData_RecValue2(Fields6(K)) = Xd.RecValue2(Fields8(K))
                If HasError Then Exit Function
            Next K
        Next J
    Loop
    Xs.NoShow = True
    RecClose
    
    Xs.Show Me
    CrossProduct = True
End Function

Public Function Group(Optional GroupFields, Optional GroupFunctions) As Boolean
'Group GroupFields by GroupFunctions (which are callback Functions2).
'Group by all other fields (or to a single record if none).
'If group function is an array function, changes field to string.

    Dim Fields2 As Variant 'GroupFields as supplied
    Dim Fields3 As Variant 'GroupFields columns
    Dim Fields6 As Variant 'Fields to group by columns
    Dim Functions2 As Variant
    Dim Functions3 As Variant
    Dim Items As Variant
    Dim Icol As Collection
    Dim Ival As Variant
    Dim Ival1 As Long
    Dim Result As Variant
    Dim I As Long, J As Long, K As Long
    
    If HasError Then Exit Function
    If DataRowCount <= 0 Then
        Group = True
        Exit Function
    End If
    RemoveIndex
    
    Fields2 = Xs.ExtractArray(GroupFields)
    If Not Xs.ProcessFields(Me, Fields2, Fields3) Then Exit Function
    Xs.MatchFields Me, Nothing, Fields2, Fields6
    
    Functions2 = Xs.ExtractArray(GroupFunctions)
    If Not ProcessFunctions(Functions2, Functions3) Then Exit Function
    If UBound(Functions3) <> UBound(Fields3) Then
        Msg "GroupFields and GroupFunctions must match"
        Exit Function
    End If
    
    For I = 1 To UBound(Functions3)
        Select Case Functions3(I)
        Case "NumArray", "TextArray", "ArrArray", "UniqueList"
            If Fld(Fields3(I)).VarType <> vbString Then
                Fld(Fields3(I)).PrevVarType = Fld(Fields3(I)).VarType
                Xs.NoShow = True
                If Not ChangeType(Fields3(I), vbString) Then Exit Function
            End If
        End Select
    Next I
    
    If UBound(Fields6) < LBound(Fields6) Then
        Set Icol = New Collection
        For I = 1 To DataRowCount
            Icol.Add I
        Next I
        Items = Array(Icol)
    Else
        If Not CreateIndex(Fields6) Then Exit Function
        Items = RowIndex.Items
    End If
    For I = LBound(Items) To UBound(Items)
        Set Icol = Items(I)
        Ival1 = Icol.Item(1)
        For J = 1 To UBound(Fields3)
            If Not DimArray(GroupData, Icol.Count, Fld(Fields3(J)).VarType) _
                Then GoTo Finish
            K = 0
            For Each Ival In Icol
                K = K + 1
                GroupData(K) = Fld(Fields3(J)).Data(Pointer(Ival))
            Next Ival
            If Not Callback(CStr(Functions3(J)), Result, Xs) Then GoTo Finish
            Fld(Fields3(J)).Data(Pointer(Ival1)) = Result
        Next J
        K = 0
        For Each Ival In Icol
            K = K + 1
            If K > 1 Then Pointer(Ival) = 0
        Next Ival
    Next I
    
    CompressPointers
    Xs.Show Me
    Group = True

Finish:
    RemoveIndex
End Function

Public Function UnGroup(ParamArray GroupFields()) As Boolean
'Ungroups where there is an array in one or more fields.
'Other fields are copied from the record being ungrouped.
'All group fields must contain arrays with the same number of elements.
'Ungroups a two-dimensional array to one-dimensional arrays.

    Dim Fields2 As Variant 'GroupFields as supplied
    Dim Fieldsx As Variant
    Dim Fields3 As Variant 'GroupFields columns
    Dim Fields6 As Variant 'Fields to group by columns
    Dim OldRowCount As Long
    Dim GroupValue As String
    Dim SplitValue As Variant
    Dim IsArray As Boolean
    Dim ArrMax As Long
    Dim I As Long, J As Long, K As Long, L As Long
    
    If HasError Then Exit Function
    If DataRowCount <= 0 Then
        UnGroup = True
        Exit Function
    End If
    RemoveIndex
    
    Fieldsx = GroupFields
    Fields2 = Xs.ExtractArray(Fieldsx)
    If Not Xs.ProcessFields(Me, Fields2, Fields3) Then Exit Function
    Xs.MatchFields Me, Nothing, Fields2, Fields6
    If UBound(Fields2) < LBound(Fields2) Then
        Msg "No group fields specified"
        Exit Function
    End If
    
    OldRowCount = DataRowCount
    For I = 1 To OldRowCount
        For J = 1 To UBound(Fields3)
            GroupValue = Fld(Fields3(J)).Data(Pointer(I))
            If Left(GroupValue, 1) = "{" And Right(GroupValue, 1) = "}" Then
                GroupValue = Mid(GroupValue, 2, Len(GroupValue) - 2)
            ElseIf Len(GroupValue) = 0 Then
                GroupValue = """"""
            Else
                Msg FieldName(Fields3(J)) & " is not an array"
                Exit Function
            End If
            
            If InStr(GroupValue, ";") > 0 Then
                IsArray = True
                SplitValue = Split(GroupValue, ";")
            ElseIf Right(GroupValue, 1) = " " Then
                IsArray = True
                SplitValue = Array(Left(GroupValue, Len(GroupValue) - 1))
            Else
                IsArray = False
                SplitValue = Split(GroupValue, ",")
            End If
            
            If J = 1 Then
                ArrMax = UBound(SplitValue)
                For K = 0 To ArrMax
                    RecAdd
                    For L = 1 To UBound(Fields6)
                        XData_RecValue2(Fields6(L)) = Fld(Fields6(L)).Data(Pointer(I))
                    Next L
                    If HasError Then Exit Function
                Next K
            ElseIf ArrMax <> UBound(SplitValue) Then
                Msg "Array sizes do not match"
                Exit Function
            End If
            
            For K = 0 To ArrMax
                RecordPos = DataRowCount - ArrMax + K
                If IsArray Then
                    XData_RecValue2(Fields3(J)) = "{" & SplitValue(K) & "}"
                ElseIf Left(SplitValue(K), 1) = """" And Right(SplitValue(K), 1) = """" Then
                    SplitValue(K) = Replace(SplitValue(K), RepChar4, """")
                    SplitValue(K) = Replace(SplitValue(K), RepChar3, ";")
                    SplitValue(K) = Replace(SplitValue(K), RepChar2, ",")
                    SplitValue(K) = Replace(SplitValue(K), RepChar1, "\")
                    XData_RecValue2(Fields3(J)) = Mid(SplitValue(K), 2, Len(SplitValue(K)) - 2)
                Else
                    XData_RecValue2(Fields3(J)) = SplitValue(K)
                End If
                If HasError Then Exit Function
            Next K
        Next J
        Pointer(I) = 0
    Next I
    
    CompressPointers
    Xs.NoShow = True
    RecClose
    
    Xs.Show Me
    UnGroup = True
End Function

Public Function RestoreTypes() As Boolean
'Restore VarTypes changed by Group.

    Dim I As Long
    
    If HasError Then Exit Function
    For I = 1 To UBound(Fld)
        If Fld(I).PrevVarType <> 0 Then
            If Not ChangeType(I, Fld(I).PrevVarType) Then Exit Function
            Fld(I).PrevVarType = 0
        End If
    Next I
    RestoreTypes = True
End Function

Public Function Total(Field, GroupFunction As String) As Variant
'Calculates the total of Field using GroupFunction

    Dim Result As Variant
    Dim I As Long

    If HasError Then Exit Function
    If DataRowCount = 0 Then Exit Function
    If Not FieldCheck(Field) Then Exit Function
    With Fld(FieldCol(Field))
        If Not DimArray(GroupData, DataRowCount, .VarType) Then Exit Function
        For I = 1 To DataRowCount
            GroupData(I) = .Data(Pointer(I))
        Next I
    End With
    If Not Callback(GroupFunction, Result, Xs) Then Exit Function
    Total = Result
End Function

Public Function Sort(Optional Key1, Optional Order1 As XlSortOrder = xlAscending, _
    Optional Key2, Optional Order2 As XlSortOrder = xlAscending, _
    Optional Key3, Optional Order3 As XlSortOrder = xlAscending, _
    Optional Key4, Optional Order4 As XlSortOrder = xlAscending, _
    Optional Key5, Optional Order5 As XlSortOrder = xlAscending, _
    Optional Key6, Optional Order6 As XlSortOrder = xlAscending) _
    As Boolean

    Dim SortFields() As Long 'SortFields columns
    Dim SortOrders() As Integer 'SortOrder corresponding
    Dim SortField As Variant
    Dim SortOrder As XlSortOrder
    Dim I As Long, J As Long
    
    If HasError Then Exit Function
    If DataRowCount <= 0 Then
        Sort = True
        Exit Function
    End If
    
    RemoveIndex
    J = -1
    ReDim SortFields(0 To 5)
    ReDim SortOrders(0 To 5)
    For I = 0 To 5
        SortField = Array(Key1, Key2, Key3, Key4, Key5, Key6)(I)
        SortOrder = Array(Order1, Order2, Order3, Order4, Order5, Order6)(I)
        If Not IsMissing(SortField) Then
            If Len(CStr(SortField)) > 0 Then
                J = J + 1
                If Not FieldCheck(SortField) Then Exit Function
                SortFields(J) = FieldCol(SortField)
                Select Case SortOrder
                Case xlAscending
                    SortOrders(J) = 1
                Case xlDescending
                    SortOrders(J) = -1
                Case Else
                    Msg "Invalid sortorder " & SortOrder
                    Exit Function
                End Select
            End If
        End If
    Next I
    If J < 0 Then
        Sort = True
        Exit Function
    End If
    ReDim Preserve SortFields(0 To J)
    ReDim Preserve SortOrders(0 To J)
    QuickSort SortFields, SortOrders, 1, DataRowCount
    
    Xs.Show Me
    Sort = True
End Function

Public Function Subtotal(GroupBy As Variant, SubFunction As XlConsolidationFunction, _
    SummaryBelowData As XlSummaryRow, ParamArray TotalList()) As Boolean
'Creates subtotals.
'In this version, Subfunction is ignored - always totals.
'SummaryBelowData is ignored, always below data.

    Dim Fields1 As Variant 'Fields to group by as supplied
    Dim Fields2 As Variant 'Fields to group by as column numbers
    Dim Fields3 As Variant 'Totals fields as supplied
    Dim Fields3a As Variant 'Totals fields as supplied
    Dim Fields4 As Variant 'Totals fields column numbers
    Dim Pointer2() As Long
    Dim StartPoint As Boolean
    Dim BreakPoint As Boolean
    Dim GTrec As Long
    Dim I As Long, J As Long
    
    If HasError Then Exit Function
    RemoveIndex
    Fields1 = Xs.ExtractArray(GroupBy)
    If Not Xs.ProcessFields(Me, Fields1, Fields2) Then Exit Function
    Fields3 = TotalList
    Fields3a = Xs.ExtractArray(Fields3)
    If Not Xs.ProcessFields(Me, Fields3a, Fields4) Then Exit Function
    
    'Add grand total row.
    RecAdd
    GTrec = DataRowCount
    
    'Look for startpoints set on the first record of a group.
    'The last record in the table is now the grand total row.
    ReDim Pointer2(1 To UBound(Pointer))
    RecStart
    Do While RecNext
        StartPoint = False
        If RecordPos = DataRowCount Then
        ElseIf RecordPos = 1 Then
            StartPoint = True
        ElseIf BreakPoint Then
            StartPoint = True
        End If
        
        'look for breakpoints set on the last record of a group.
        BreakPoint = False
        If RecordPos = DataRowCount Then
        ElseIf RecordPos = DataRowCount - 1 Then
            BreakPoint = True
        Else
            For I = 1 To UBound(Fields2)
                If Fld(Fields2(I)).Data(Pointer(RecordPos)) <> Fld(Fields2(I)).Data(Pointer(RecordPos + 1)) Then Exit For
            Next I
            If I <= UBound(Fields2) Then
                BreakPoint = True
            End If
        End If
        
        'Make space.
        If StartPoint Then
            DataCount = DataCount + 1
            If DataCount > UBound(Fld(1).Data) Then
                For I = 1 To UBound(Fld)
                    ReDim Preserve Fld(I).Data(1 To UBound(Fld(I).Data) + GranSize)
                Next I
            End If
        End If
        
        'Create subtotals.
        If RecordPos < DataRowCount Then
            J = J + 1
            If J > UBound(Pointer2) Then
                ReDim Preserve Pointer2(1 To UBound(Pointer2) + GranSize)
            End If
            Pointer2(J) = Pointer(RecordPos)
            For I = 1 To UBound(Fields4)
                On Error Resume Next
                With Fld(Fields4(I))
                    .Data(DataCount) = .Data(DataCount) + .Data(Pointer(RecordPos))
                    .Data(Pointer(GTrec)) = .Data(Pointer(GTrec)) + .Data(Pointer(RecordPos))
                End With
                If Err = 0 Then
                    On Error GoTo 0
                Else
                    On Error GoTo 0
                    Msg "Cannot create subtotal for " & Fld(Fields4(I)).Name
                    Exit Function
                End If
            Next I
        End If
            
        If BreakPoint Then
            J = J + 1
            If J > UBound(Pointer2) Then
                ReDim Preserve Pointer2(1 To UBound(Pointer2) + GranSize)
            End If
            Pointer2(J) = DataCount
            Fld(Fields2(1)).Data(DataCount) = "Total"
'            For I = 2 To UBound(Fields2)
'                Fld(Fields2(I)).Data(DataCount) = Fld(Fields2(I)).Data(Pointer(RecordPos))
'            Next I
        End If
        
        If RecordPos = DataRowCount Then
            J = J + 1
            If J > UBound(Pointer2) Then
                ReDim Preserve Pointer2(1 To UBound(Pointer2) + GranSize)
            End If
            Pointer2(J) = Pointer(GTrec)
            Fld(Fields2(1)).Data(Pointer(GTrec)) = "Grand Total"
        End If
    Loop
    
    DataRowCount = J
    Pointer = Pointer2
    Xs.NoShow = True
    RecClose
    Subtotal = True
End Function

Public Function Export(Name As String, ParamArray Field()) As Boolean
'Export all or named fields to external workbook.

    Dim Ws As Worksheet
    Dim Wb2 As Workbook
    Dim Field2 As Variant
    Dim Fieldx As Variant
    Dim Field3 As Variant 'Field columns
    Dim I As Long
    
    If HasError Then Exit Function
    Set Ws = Xs.CreateSource(Me, Name)
    If Ws Is Nothing Then Exit Function
    DoEvents
    Fieldx = Field
    Field2 = Xs.ExtractArray(Fieldx)
    If UBound(Field2) >= LBound(Field2) Then
        If Not Xs.ProcessFields(Me, Field2, Field3) Then Exit Function
    Else
        Xs.MatchFields Me, Nothing, Xs.ExtractArray, Field3
    End If
    
    For I = 1 To UBound(Field3)
        Ws.Cells(1, I).Value = Fld(Field3(I)).Name
        If DataRowCount > 0 Then
            If Not GetColumnData(Field3(I)) Then Exit Function
            Ws.Cells(2, I).Resize(DataRowCount).Value = GroupData
        End If
    Next I
    GroupData = Empty
    If DataRowCount > 0 Then _
        Ws.Cells(2, 1).Resize(DataRowCount).EntireRow.RowHeight = Ws.Rows(1).RowHeight
    Export = True
End Function

Public Function SaveXML(Filename As Variant) As Boolean
'Saves as a XML file in PersistXML format.
'Do not save more than once to the same destination.

    Dim Rd As Object

    If HasError Then Exit Function
    Set Rd = Xs.GetRd
    Rd.CreateFrom Me
    Rd.SaveXML Filename
    Rd.RsetClose
    If Rd.HasError Then Exit Function
    SaveXML = True
End Function

Public Function OpenXML(Filename As Variant) As Boolean
'Opens a XML file in PersistXML format.
'Recordset is opened forward only.

    Dim Rd As Object

    If HasError Then Exit Function
    Set Rd = Xs.GetRd
    Rd.OpenXML Filename
    If Rd.HasError Then
        HasError = True
        Exit Function
    End If
    Xs.NoShow = True
    CreateFrom Rd
    If HasError Then Exit Function
    Set Rd = Xs.GetRd(Rd)
    Rd.RsetClose
    
    Xs.Show Me
    OpenXML = True
End Function

Public Function SaveBin(Filename As Variant) As Boolean
'Saves as a binary file.
'Does not support extended character set.

    Dim FileNum As Integer
    Dim Fld2 As FieldType2
    Dim BinData As BinDataType
    Dim I As Long
    
    'Delete contents of Filename before writing.
    If HasError Then Exit Function
    FileNum = FreeFile
    Open Filename For Output Access Write Lock Read Write As #FileNum
    Close #FileNum
    
    FileNum = FreeFile
    Open Filename For Binary Access Write Lock Read Write As #FileNum
    With BinData
        .ThisObjName = ThisObjName
        .DataRowCount = DataRowCount
        .DataCount = DataCount
        .Pointer = Pointer
        .FldCount = UBound(Fld)
    End With
    Put FileNum, , BinData
    
    For I = 1 To UBound(Fld)
        Fld2.Name = Fld(I).Name
        Fld2.VarType = Fld(I).VarType
        Fld2.Data = Fld(I).Data
        Put FileNum, , Fld2
    Next I
    Close #FileNum
    SaveBin = True
End Function

Public Function OpenBin(Filename As Variant) As Boolean
'Opens from a binary file.

    Dim FileNum As Integer
    Dim FldCount As Long
    Dim PtrCount As Long
    Dim Fld2 As FieldType2
    Dim BinData As BinDataType
    Dim I As Long
    
    'Check file exists before reading.
    If HasError Then Exit Function
    FileNum = FreeFile
    On Error Resume Next
    Open Filename For Input Access Read Lock Write As #FileNum
    If Err = 0 Then
        On Error GoTo 0
        Close #FileNum
    Else
        Msg Err.Description
        On Error GoTo 0
        Exit Function
    End If
    
    IndexType = IndexMultiple
    RecordPos = 0
    GroupData = Empty
    Set RowIndex = Nothing
    Set FieldList = New Dictionary
    FieldList.CompareMode = TextCompare
    
    FileNum = FreeFile
    Open Filename For Binary Access Read Lock Write As #FileNum
    Get FileNum, , BinData
    With BinData
        ThisObjName = .ThisObjName
        DataRowCount = .DataRowCount
        DataCount = .DataCount
        Pointer = .Pointer
        ReDim Fld(1 To .FldCount)
    End With
    
    For I = 1 To UBound(Fld)
        Get FileNum, , Fld2
        Fld(I).Name = Fld2.Name
        FieldList.Add Fld2.Name, I
        Fld(I).VarType = Fld2.VarType
        Fld(I).Data = Fld2.Data
    Next I
    Close FileNum
    
    Xs.Show Me
    OpenBin = True
End Function

Public Function ExtractXML(Filename As Variant) As Boolean
'Extracts from a XML file created using SaveAllXML.
'Returns xml DomDocument in Filename.

    Dim XMLDoc0 As Object
    Dim XMLDoc1 As Object
    Dim NodeList0 As Object
    Dim NodeList1 As Object
    Dim Element0 As Object
    Dim Element1 As Object

    If HasError Then Exit Function
    
    If IsObject(Filename) Then
        Set XMLDoc0 = Filename
    Else
        Set XMLDoc0 = CreateObject("MSXML2.DOMDocument.6.0")
        On Error Resume Next
        XMLDoc0.Load Filename
        If Err = 0 Then
            On Error GoTo 0
        Else
            Msg Err.Description
            On Error GoTo 0
            Exit Function
        End If
        If XMLDoc0.parseError.ErrorCode <> 0 Then
            Msg "Unable to open input file" & vbLf & XMLDoc0.parseError.reason
            Exit Function
        End If
        Set Filename = XMLDoc0
    End If
    
    Set NodeList0 = XMLDoc0.getElementsByTagName("AllTables")
    If NodeList0.Length <> 1 Then
        Msg "No tables on xml file"
        Exit Function
    End If
    
    Set Element0 = NodeList0.Item(0)
    Set NodeList1 = Element0.getElementsByTagName(ThisObjName)
    If NodeList1.Length <> 1 Then
        Msg "Table " & ThisObjName & " not on xml file"
        Exit Function
    End If
    
    Set Element1 = NodeList1.Item(0)
    Set XMLDoc1 = CreateObject("MSXML2.DOMDocument.6.0")
    XMLDoc1.appendChild Element1
    Xs.NoShow = True
    OpenXML XMLDoc1
    If HasError Then Exit Function
    
    Xs.Show Me
    ExtractXML = True
End Function

Public Function RefreshPivot(Ws As Worksheet, PTName As String) As Boolean
'Create pivot table, or refresh if already exists.

    Dim Xs As New XShared

    If HasError Then Exit Function
    RefreshPivot = Xs.RefreshPivot(Me, Ws, PTName)
End Function

Private Function ProcessFunctions(Functions As Variant, Functions2 As Variant) As Boolean
'Checks and converts an array of functions.
'Returns Functions2 as a 1-based array.
    
    Dim I As Long, J As Long
    
    If UBound(Functions) < LBound(Functions) Then
        Functions2 = Xs.ExtractArray
        ProcessFunctions = True
        Exit Function
    End If
    
    ReDim Functions2(1 To UBound(Functions) - LBound(Functions) + 1) As String
    For I = LBound(Functions) To UBound(Functions)
        J = J + 1
        Functions2(J) = Functions(I)
    Next I
    ProcessFunctions = True
End Function

Private Sub QuickSort(SortFields() As Long, SortOrders() As Integer, lngFirst As Long, lngLast As Long)
                            
    ' QuickSort algorithm used to sort array.
    
    Dim lngLow      As Long
    Dim lngHigh     As Long
    Dim lngMiddle   As Long
    Dim lngTempVal  As Long
    Dim lngTestVal  As Long

    
    If lngFirst < lngLast Then
        lngMiddle = (lngFirst + lngLast) / 2
        lngTestVal = Pointer(lngMiddle)
        lngLow = lngFirst
        lngHigh = lngLast
        Do
            Do While Comp(SortFields, SortOrders, Pointer(lngLow), lngTestVal) < 0
                lngLow = lngLow + 1
            Loop
            Do While Comp(SortFields, SortOrders, Pointer(lngHigh), lngTestVal) > 0
                lngHigh = lngHigh - 1
            Loop
            If (lngLow <= lngHigh) Then
                lngTempVal = Pointer(lngLow)
                Pointer(lngLow) = Pointer(lngHigh)
                Pointer(lngHigh) = lngTempVal
                lngLow = lngLow + 1
                lngHigh = lngHigh - 1
            End If
        Loop While (lngLow <= lngHigh)
        If lngFirst < lngHigh Then QuickSort SortFields, SortOrders, lngFirst, lngHigh
        If lngLow < lngLast Then QuickSort SortFields, SortOrders, lngLow, lngLast
    End If

End Sub

Private Function Comp(SortFields() As Long, SortOrders() As Integer, _
    Index1 As Long, Index2 As Long) As Integer
'Compares record at Index1 with record at Index2
'Returns -1, 0, 1 if <, =, or >

    Dim I As Long
    
    For I = 0 To UBound(SortFields)
        With Fld(SortFields(I))
            Select Case .Data(Index1)
            Case Is < .Data(Index2)
                Comp = -1 * SortOrders(I)
                Exit Function
            Case Is > .Data(Index2)
                Comp = 1 * SortOrders(I)
                Exit Function
            End Select
        End With
    Next I
End Function

Private Function DimData(Field As Long, DataSize As Long) As Boolean
'Dimension data area.

    With Fld(Field)
        DimData = DimArray(.Data, DataSize, .VarType)
    End With
End Function

Private Function DimArray(DataArray As Variant, DataSize As Long, DataType As VbVarType) As Boolean
'Dimension a variant as a 1-dimensional array of specified type

    Select Case DataType
    Case vbInteger
        ReDim DataArray(1 To DataSize) As Integer
    Case vbLong
        ReDim DataArray(1 To DataSize) As Long
    Case vbSingle
        ReDim DataArray(1 To DataSize) As Single
    Case vbDouble
        ReDim DataArray(1 To DataSize) As Double
    Case vbCurrency
        ReDim DataArray(1 To DataSize) As Currency
    Case vbDate
        ReDim DataArray(1 To DataSize) As Date
    Case vbString
        ReDim DataArray(1 To DataSize) As String
    Case vbBoolean
        ReDim DataArray(1 To DataSize) As Boolean
    Case Else
        Msg "Data type not recognised " & DataType
        Exit Function
    End Select
    DimArray = True
End Function

Private Function ConvertCriteria(Criteria2 As Variant, WhereFields() As Long, _
    WhereValues() As Variant, WhereTypes() As TestType, HasNot As Boolean) As Boolean
'Extract and convert criteria, separating type from value and converting values by data type.
'Returns Criteria2, WhereFields, WhereValues,  WhereTypes and HasNot if first item is "Not".

    Dim I As Long, J As Long, K As Long

    If UBound(Criteria2) - LBound(Criteria2) > 0 Then
        If Criteria2(LBound(Criteria2)) = "Not" Then
            HasNot = True
            For I = LBound(Criteria2) + 1 To UBound(Criteria2)
                Criteria2(I - 1) = Criteria2(I)
            Next I
            ReDim Preserve Criteria2(LBound(Criteria2) To UBound(Criteria2) - 1)
        End If
        If (UBound(Criteria2) - LBound(Criteria2) + 1) Mod 2 <> 0 Then
            Msg "Criteria must have the same number of fields and values"
            Exit Function
        End If
        ReDim WhereFields(1 To (UBound(Criteria2) - LBound(Criteria2) + 1) / 2)
        ReDim WhereValues(1 To (UBound(Criteria2) - LBound(Criteria2) + 1) / 2)
        ReDim WhereTypes(1 To (UBound(Criteria2) - LBound(Criteria2) + 1) / 2)
        For I = LBound(Criteria2) To UBound(Criteria2) - 1 Step 2
            J = J + 1
            If Not FieldCheck(Criteria2(I)) Then Exit Function
            WhereFields(J) = FieldCol(Criteria2(I))
            SplitTest Criteria2(I + 1), WhereValues(J), WhereTypes(J)
            If WhereTypes(J) = TypIn Then
                For K = LBound(WhereValues(J)) To UBound(WhereValues(J))
                    If Not ConvertType(WhereValues(J)(K), Fld(WhereFields(J)).VarType) Then Exit Function
                Next K
            Else
                If Not ConvertType(WhereValues(J), Fld(WhereFields(J)).VarType) Then Exit Function
            End If
        Next I
    Else
        Criteria2 = Criteria2(LBound(Criteria2))
    End If
    ConvertCriteria = True
End Function

Private Function CompareCriteria(Criteria2 As Variant, WhereFields() As Long, _
    WhereValues() As Variant, WhereTypes() As TestType, HasNot As Boolean, Result As Boolean) As Boolean
'Check current record against criteria and return Result.
    
    Dim I As Long

    Result = True
    If IsArray(Criteria2) Then
        For I = 1 To UBound(WhereFields)
            If Not RunTest(Fld(WhereFields(I)).Data(Pointer(RecordPos)), _
                WhereValues(I), WhereTypes(I)) Then Result = False
        Next I
        Result = Result Xor HasNot
    ElseIf Len(Criteria2) > 0 Then
        If Not Callback(CStr(Criteria2), Result, Xs) Then Exit Function
    End If
    CompareCriteria = True
End Function

Private Function DoCalculate(CalcProc As Variant, Result As Variant, Field2 As Variant, _
    Offset As Variant, MyXs As XShared, ByVal Params As Variant) As Boolean
'Calculates result using callback or formula.
'Formula is an array of formula elements where field names are separate elements.

    Dim FormText As String
    
    If IsArray(CalcProc) Then
        FormText = GetFormText(CalcProc, Field2, Offset)
        On Error Resume Next
        Result = Application.Evaluate(FormText)
        If Err = 0 Then
            On Error GoTo 0
            If Not CheckResult(Result, FormText) Then Exit Function
        Else
            On Error GoTo 0
            Msg "Error in carrying out calculation" & vbLf & _
                "Please see VBA Immediate Pane for more information"
            Debug.Print "Record Number " & RecordPos
            Debug.Print "Formula =" & FormText
            Exit Function
        End If
    ElseIf Len(CalcProc) > 0 Then
        If Not Callback(CStr(CalcProc), Result, MyXs, Params) Then Exit Function
    End If
    DoCalculate = True
End Function

Private Function GetFormText(ByVal CalcProc As Variant, Field2 As Variant, Offset As Variant) As String
'Returns text for formula.

    Dim I As Long
    Dim Pos As Long
    
    For I = LBound(CalcProc) To UBound(CalcProc)
        If Field2(I) > 0 Then
            If Offset(I) Then
                If RecordPos <= 1 Then
                    Pos = 0
                    Select Case Fld(Field2(I)).VarType
                    Case vbString
                        CalcProc(I) = """"""
                    Case vbBoolean
                        CalcProc(I) = "False"
                    Case Else
                        CalcProc(I) = 0
                    End Select
                Else
                    Pos = RecordPos - 1
                End If
            Else
                Pos = RecordPos
            End If
            
            If Pos > 0 Then
                Select Case Fld(Field2(I)).VarType
                Case vbString
                    CalcProc(I) = CStr(Fld(Field2(I)).Data(Pointer(Pos)))
                    If Left(CalcProc(I), 2) = "{""" And Right(CalcProc(I), 2) = """}" Then
                        CalcProc(I) = Replace(CalcProc(I), RepChar4, """""")
                        CalcProc(I) = Replace(CalcProc(I), RepChar3, ";")
                        CalcProc(I) = Replace(CalcProc(I), RepChar2, ",")
                        CalcProc(I) = Replace(CalcProc(I), RepChar1, "\")
                    ElseIf Left(CalcProc(I), 1) = "{" And Right(CalcProc(I), 1) = "}" Then
                            'No quotes needed for array
                    Else
                        CalcProc(I) = SplitText(CStr(CalcProc(I)))
                    End If
                Case vbDate
                    CalcProc(I) = CStr(CDbl(Fld(Field2(I)).Data(Pointer(Pos))))
                Case Else
                    CalcProc(I) = CStr(Fld(Field2(I)).Data(Pointer(Pos)))
                End Select
            End If
        End If
    Next I
    GetFormText = Join(CalcProc, vbNullString)
End Function

Private Function Callback(ProcName As String, Result As Variant, MyXs As XShared, Optional Params As Variant) As Boolean
'Runs callback procedure and returns Result.
'If callback procedure is not in XShared tries XUserFunctions

    Dim ThisObj As ArrData
        
    Set ThisObj = Me
    On Error Resume Next
    If IsArray(Params) Then
        Result = CallByName(Xs, ProcName, VbMethod, ThisObj, Params)
    Else
        Result = CallByName(Xs, ProcName, VbMethod, ThisObj)
    End If
    If Err = 0 Then
        On Error GoTo 0
        If HasError Then Exit Function
    ElseIf Err = 438 Then 'Object doesn't support this property or method
        Err.Clear
        If IsArray(Params) Then
            Result = CallByName(Xs.GetUf, ProcName, VbMethod, ThisObj, Params)
        Else
            Result = CallByName(Xs.GetUf, ProcName, VbMethod, ThisObj)
        End If
        If Err = 0 Then
            On Error GoTo 0
            If HasError Then Exit Function
        Else
            On Error GoTo 0
            Msg "Error calling " & ProcName
            Exit Function
        End If
    Else
        On Error GoTo 0
        Msg "Error calling " & ProcName
        Exit Function
    End If
    Callback = True
End Function

Private Sub SplitTest(TestInfo As Variant, SplitValue As Variant, SplitType As TestType)
'Splits TestInfo and returns value in SplitValue and type in SplitType.
    
    Dim I As Long

    If IsArray(TestInfo) Then
        ReDim SplitValue(LBound(TestInfo) To UBound(TestInfo)) As Variant
        For I = LBound(TestInfo) To UBound(TestInfo)
            SplitValue(I) = TestInfo(I)
        Next I
        SplitType = TypIn
        Exit Sub
    ElseIf IsEmpty(TestInfo) Then
        SplitValue = Empty
        SplitType = TypIgnore
        Exit Sub
    End If
    
    Select Case Left(CStr(TestInfo), 2)
    Case ">="
        SplitValue = Mid(CStr(TestInfo), 3)
        SplitType = TypGE
        Exit Sub
    Case "<="
        SplitValue = Mid(CStr(TestInfo), 3)
        SplitType = TypLE
        Exit Sub
    Case "<>"
        SplitValue = Mid(CStr(TestInfo), 3)
        SplitType = TypNE
        Exit Sub
    End Select
    
    Select Case Left(CStr(TestInfo), 1)
    Case ">"
        SplitValue = Mid(CStr(TestInfo), 2)
        SplitType = TypGreater
        Exit Sub
    Case "<"
        SplitValue = Mid(CStr(TestInfo), 2)
        SplitType = TypLess
        Exit Sub
    Case "="
        SplitValue = Mid(CStr(TestInfo), 2)
        SplitType = TypEquals
        Exit Sub
    End Select
    SplitValue = TestInfo
    SplitType = TypEquals
End Sub

Private Function ConvertType(DataValue As Variant, DataType As VbVarType) As Boolean
'Converts DataValue to specified data type. returns True if successful.
'Converts empty string as if it were Empty.

    If Xs.ConvertType(DataValue, DataType, HasError) Then
        ConvertType = True
    Else
        HasError = True
    End If
End Function

Private Function CheckNull(Value As Variant) As Boolean
'Returns true if Value is empty, an empty string or 0 or False

    Select Case VarType(Value)
    Case vbEmpty, vbNull
        CheckNull = True
    Case vbString
        If Len(Value) = 0 Then CheckNull = True
    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate
        If Value = 0 Then CheckNull = True
    Case vbBoolean
        If Not Value Then CheckNull = True
    End Select
End Function

Private Function RunTest(TestValue As Variant, SplitValue As Variant, SplitType As TestType) As Boolean
'Runs comparison - True if equal.

    Dim I As Long

    Select Case SplitType
    Case TypIgnore
        RunTest = True
    Case TypEquals
        If TestValue = SplitValue Then RunTest = True
    Case TypLess
        If TestValue < SplitValue Then RunTest = True
    Case TypLE
        If TestValue <= SplitValue Then RunTest = True
    Case TypGreater
        If TestValue > SplitValue Then RunTest = True
    Case TypGE
        If TestValue >= SplitValue Then RunTest = True
    Case TypNE
        If TestValue <> SplitValue Then RunTest = True
    Case TypIn
        For I = LBound(SplitValue) To UBound(SplitValue)
            If TestValue = SplitValue(I) Then
                RunTest = True
                Exit For
            End If
        Next I
    End Select
End Function

Private Function ConvertError(Value As Variant) As String

    If Not IsError(Value) Then Exit Function
    Select Case Value
    Case CVErr(xlErrDiv0)
        ConvertError = "Error #DIV/0!"
    Case CVErr(xlErrNA)
        ConvertError = "Error #N/A"
    Case CVErr(xlErrName)
        ConvertError = "Error #NAME?"
    Case CVErr(xlErrNull)
        ConvertError = "Error #NULL!"
    Case CVErr(xlErrNum)
        ConvertError = "Error #NUM!"
    Case CVErr(xlErrRef)
        ConvertError = "Error #REF!"
    Case CVErr(xlErrValue)
        ConvertError = "Error #VALUE"
    Case Else
        ConvertError = CStr(Value)
    End Select
End Function

Private Sub CompressPointers()
'Remove zero pointers and move everything down.

    Dim I As Long, J As Long

    J = 0
    For I = 1 To DataRowCount
        If Pointer(I) > 0 Then
            J = J + 1
            Pointer(J) = Pointer(I)
        End If
    Next I
    DataRowCount = J
    If J > 0 Then
        ReDim Preserve Pointer(1 To J)
    Else
        ReDim Pointer(1 To GranSize)
    End If
End Sub

Private Function SplitText(TextVal As String) As String
'Splits a text into quote-delimited chunks up to 127 chars.
'Replaces " with "" - 127 ensures this will never go over 255.

    Dim ChunkCount As Long
    Dim Chunks() As String
    Dim I As Long
    Const ChunkSize As Long = 127
    
    ChunkCount = (Len(TextVal) - 1) \ ChunkSize + 1
    ReDim Chunks(1 To ChunkCount)
    For I = 1 To ChunkCount
        Chunks(I) = Mid(TextVal, (I - 1) * ChunkSize + 1, ChunkSize)
        Chunks(I) = Replace(Chunks(I), """", """""")
    Next I
    SplitText = """" & Join(Chunks, """&""") & """"
End Function

Private Function CheckResult(Result As Variant, FormText As String) As Boolean
'Checks calculation result for errors.

    If IsError(Result) Then
        Msg ConvertError(Result) & " returned when carrying out calculation" & vbLf & _
            "Please see VBA Immediate Pane for more information"
        Debug.Print "Record Number " & RecordPos
        Debug.Print "Formula =" & FormText
        Exit Function
    End If
    If VarType(Result) = vbString Then
        If Left(Result, 1) = "#" Then
            If Right(Result, 1) = "!" Then
                Msg Mid(Result, 2, Len(Result) - 2)
                Exit Function
            End If
        End If
    End If
    CheckResult = True
End Function

Private Function FieldCheck(Field As Variant) As Boolean
'Check if a field exists. May be name or number.

    If XData_Exists(Field) Then
        FieldCheck = True
    ElseIf Not HasError Then
        MsgBox "Cannot find field " & Field & " on data"
        HasError = True
    End If
End Function

Private Function FieldCol(Field As Variant) As Long
'Returns the column number of Field which may be column name or number.

    If VarType(Field) = vbString Then
        FieldCol = FieldList(Field)
    Else
        FieldCol = Field
    End If
End Function

Private Function FieldName(Field As Variant) As String
'Returns the column name of Field which may be column name or number.

    If VarType(Field) = vbString Then
        FieldName = Field
    Else
        FieldName = Fld(Field).Name
    End If
End Function

Private Sub Msg(ErrMsg As String)
'Displays error message unless HasError is already true, then sets HasError.

    If Not HasError Then MsgBox ErrMsg
    HasError = True
End Sub

Private Function Split2(ByVal TextList As String, SepChar As String) As Variant
    'Unpack a comma delimited list into an array.
    'Modified from X_Split to produce a 0 based array and not to trim values.

    Dim TextVal As String 'One value unpacked from TextList.
    Dim LastField As Boolean
    Dim QuoteFound As Boolean
    Dim I As Integer
    Dim ListArray() As String
    Const Granule As Integer = 50
    
    If Len(TextList) = 0 Then
        Split2 = Xs.ExtractArray
        Exit Function
    End If
        
    ReDim ListArray(0 To Granule - 1)
    LastField = True
    I = -1
    Do
        X_Unpack TextList, TextVal, LastField, QuoteFound, SepChar
        TextVal = Replace(TextVal, """""", """")
        If LastField Then Exit Do
        I = I + 1
        If I > UBound(ListArray) Then ReDim Preserve ListArray(0 To UBound(ListArray) + Granule)
        ListArray(I) = TextVal
    Loop
    ReDim Preserve ListArray(0 To I)
    Split2 = ListArray
End Function

Private Static Sub X_Unpack(RecordText As String, StringValue As String, _
    LastField As Boolean, QuoteFound As Boolean, SepChar As String)
    'Unpack next field of RecordText into StringValue.
    'LastField should be initialised as true and is set to true
    'when unpacking beyond the last field.
    'Strips out text delimiting quotes.
    'Handles comma and quotes within text delimited by quotes,
    'and recognises that an even number of quotes before comma or
    'end of line is not a terminator.
    'Returns an empty string for two adjacent commas or one
    'comma at the end of RecordText or empty RecordText.
    'QuoteFound is TRUE if the field returned is delimited by quotes.
    'Assumes the user does not modify LastField between calls.
    'If the last field starts with quotes but does not end with quotes,
    'it is not returned but both LastField and QuoteFound are set to true;
    'in this case there may be a return character in the middle of the field,
    'and RecordText can be extended and LastField cleared before next call.
    'Modified to allow SepChar as a variable
    
    Dim FirstPos As Long 'Position of start of current field
    Dim SecondPos As Long 'Position of end of current field
    Dim SearchPos As Long 'Position from which to look for terminator
    Dim CountQuotes As Long 'Number of quotes counted before terminator
    Dim RecLen As Long 'Length of RecordText
    
    If LastField Then
        FirstPos = 1
        LastField = False
    ElseIf FirstPos = 0 Then
        QuoteFound = False
        LastField = True
        Exit Sub
    End If
    If Mid$(RecordText, FirstPos, 1) = """" Then
        QuoteFound = True
        SearchPos = FirstPos + 1
        Do
            SecondPos = InStr(SearchPos, RecordText, SepChar)
            If SecondPos = 0 Then
                RecLen = Len(RecordText)
                For CountQuotes = 0 To RecLen - FirstPos - 1
                    If Mid$(RecordText, RecLen - CountQuotes, 1) <> """" Then Exit For
                Next CountQuotes
                If CountQuotes Mod 2 = 1 Then
                    StringValue = Mid$(RecordText, FirstPos + 1, RecLen - FirstPos - 1)
                    FirstPos = 0
                Else
                    LastField = True
                End If
                Exit Sub
            Else
                For CountQuotes = 0 To SecondPos - FirstPos - 2
                    If Mid$(RecordText, SecondPos - CountQuotes - 1, 1) <> """" Then Exit For
                Next CountQuotes
                If CountQuotes Mod 2 = 1 Then
                    StringValue = Mid$(RecordText, FirstPos + 1, SecondPos - FirstPos - 2)
                    FirstPos = SecondPos + 1
                    Exit Sub
                Else
                    SearchPos = SecondPos + 1
                End If
            End If
        Loop
    Else
        QuoteFound = False
        SecondPos = InStr(FirstPos, RecordText, SepChar)
        If SecondPos = 0 Then
            StringValue = Mid$(RecordText, FirstPos)
            FirstPos = 0
            Exit Sub
        Else
            StringValue = Mid$(RecordText, FirstPos, SecondPos - FirstPos)
            FirstPos = SecondPos + 1
        End If
    End If
End Sub


