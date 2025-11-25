VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Implements XData
#Const ValueOverwrite = True 'If true, setting Value overwrites existing value in cell.
#Const LiveMode = True 'If true, some procedures/functions become private.

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

Private Type FieldType
    Name As String
    VarType As VbVarType
    IsText As Boolean
    Formula As String
    ArrFormula As String
    NumFormat As String
End Type

Public HasError As Boolean
Public DataRowCount As Long
Public DataRange As Range
Public SheetName As String
Public HdrRange As Range
Dim Xs As New XShared
Dim FieldList As Dictionary
Dim Fld() As FieldType '1-based
Dim RowIndex As Dictionary
Dim IndexFields As Variant
Dim RowPending As Long
Dim KeyPending As String
Dim BufPending As Boolean
Dim ListObj As ListObject
Dim Buf As Variant
Dim Arr As Variant
Dim ArrPos As Long
Dim ArrCount As Long
Dim Dirty As Boolean
Dim RecordPos As Long
Dim ShowTotals As Boolean
Dim TotalsRowCount As Long
Const SepChar As Long = &H25C6
Const SepChar2 As Long = &H25C8

Enum AlwaysFillOptions
    AlwaysFillWithRow = 1
    AlwaysFillWithoutRow = -1
    DoNotAlwaysFill = 0
End Enum

Enum ClearOptions
    ClearExceptFirstRow = 1
    ClearAllRows = -1
    DoNotClear = 0
End Enum

#If LiveMode Then
Const ArrSize As Long = 100
#Else
Const ArrSize As Long = 3
#End If

Function ObjType() As String
    ObjType = "TableData"
End Function

Private Function XData_ObjType() As String
    XData_ObjType = "TableData"
End Function

Private Function XData_GetType(Index As Variant) As Variant

    XData_GetType = -1
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

Function Initialise(ByVal TableRange As Range, Optional AllowNull As Boolean = False, Optional ExtendAcross As Boolean = False, _
    Optional ExtendDown As Boolean = False, Optional IgnoreDupHeaders As Boolean = False) As Boolean
'Initialises for an existing range.
'Unless AllowNull is true, gives an error message if no data.
'If ExtendAcross or ExtendDown is true, extends to the last filled cells below/right of TableRange
'(if a table, extends to the last row of the table if nothing below the table but does not change size of table).
'Picks up formulae if a table or at least one data row
'Picks up numberformats from row below header in all cases.
'Returns True if no errors.
    
    Dim ExtraRow As Boolean
    Dim I As Long, J As Long, K As Long, M As Long
    
    If HasError Then Exit Function
    If ExtendAcross Then
        With TableRange
            I = .Worksheet.Cells(.Row, .Worksheet.Columns.Count).End(xlToLeft).Column
            If I < .Column Then I = .Column
            Set TableRange = .Resize(, I - .Column + 1)
        End With
    End If
    If ExtendDown Then
        With TableRange
            M = .Row
            For J = .Column To .Column + .Columns.Count - 1
                K = .Worksheet.Cells(.Worksheet.Rows.Count, J).End(xlUp).Row
                If M < K Then M = K
            Next J
            If M < .Row Then M = .Row
            Set TableRange = .Resize(M - .Row + 1)
        End With
    End If
    
    Set ListObj = Nothing
    On Error Resume Next
    Set ListObj = TableRange.ListObject
    On Error GoTo 0
    
    'Remove totals row from data.
    'Remove unused first row of empty table from data.
    If Not ListObj Is Nothing Then
        If TableRange.Row <> ListObj.HeaderRowRange.Row Then
            MsgBox "Table header is not first row on " & TableRange.Worksheet.Name
            HasError = True
            Exit Function
        ElseIf ListObj.ShowTotals Then
            ShowTotals = True
            TotalsRowCount = 1
            If ListObj.TotalsRowRange.Row = TableRange.Row + TableRange.Rows.Count - 1 Then
                Set TableRange = TableRange.Resize(TableRange.Rows.Count - 1)
            End If
        Else
            ShowTotals = False
            TotalsRowCount = 0
        End If
        If ListObj.ListRows.Count = 0 And TableRange.Rows.Count = 2 Then
            Set TableRange = TableRange.Resize(1)
        End If
    Else
        ShowTotals = False
        TotalsRowCount = 0
    End If
    
    DataRowCount = TableRange.Rows.Count - 1
    Set DataRange = Nothing
    Set HdrRange = TableRange.Resize(1)
    ReDim Buf(1 To 1, 1 To HdrRange.Columns.Count) As Variant
    SheetName = TableRange.Worksheet.Name
    Set FieldList = New Dictionary
    FieldList.CompareMode = TextCompare
    RemoveIndex
    
    ReDim Fld(1 To TableRange.Columns.Count)
    If Not ListObj Is Nothing Then
        If ListObj.ListRows.Count = 0 Then
            ListObj.ListRows.Add
            ExtraRow = True
        End If
    End If
    For I = 1 To TableRange.Columns.Count
        With TableRange.Cells(1, I)
            If IsError(.Value) Then
                MsgBox "Error in header on " & TableRange.Worksheet.Name
                HasError = True
                Exit Function
            End If
            Fld(I).Name = Trim(CStr(.Value))
            If Len(Trim(.Value)) > 0 Then
                If FieldList.Exists(Trim(.Value)) Then
                    If Not IgnoreDupHeaders Then
                        MsgBox "Duplicate header " & Trim(.Value) & " on " & TableRange.Worksheet.Name
                        HasError = True
                        Exit Function
                    End If
                Else
                    FieldList.Add Trim(.Value), I
                End If
            End If
        End With
        With TableRange.Cells(1, I).Offset(1, 0)
            If DataRowCount >= 1 Or ExtraRow Then
                If .HasFormula Then
                    If .HasArray Then
                        Fld(I).ArrFormula = .FormulaArray
                    Else
                        Fld(I).Formula = .Formula
                    End If
                End If
            End If
            Fld(I).NumFormat = .NumberFormatLocal
            Fld(I).VarType = VarType(.Value)
            If Fld(I).VarType = vbDouble Then
                If .Value = Int(.Value) Then
                    Fld(I).VarType = vbLong
                End If
            End If
        End With
    Next I
    If ExtraRow Then ListObj.ListRows(1).Delete
    
    If DataRowCount >= 1 Then
        Set DataRange = TableRange.Resize(DataRowCount).Offset(1)
    ElseIf Not AllowNull Then
        MsgBox "No data on " & TableRange.Worksheet.Name
        HasError = True
        Exit Function
    End If
    Arr = Empty
    ArrPos = 0
    RecordPos = 0
    ArrCount = 0
    Dirty = False
    Initialise = True
End Function

Function Create(TableStart As Range, ParamArray Headers()) As Boolean
'Initialises a new range and creates headers.
'Parameter is a variant containing an array (for use with Array function), a list or a single value.

    Dim Headers2 As Variant
    Dim Headersx As Variant
    Dim I As Long, J As Long

    If HasError Then Exit Function
    Headersx = Headers
    Headers2 = Xs.ExtractArray(Headersx)
    DataRowCount = 0
    
    Set ListObj = Nothing
    On Error Resume Next
    Set ListObj = TableStart.ListObject
    On Error GoTo 0
    If Not ListObj Is Nothing Then
        If ListObj.ShowTotals Then
            ShowTotals = True
            TotalsRowCount = 1
        Else
            ShowTotals = False
            TotalsRowCount = 0
        End If
    Else
        ShowTotals = False
        TotalsRowCount = 0
    End If
    
    Set DataRange = Nothing
    Set HdrRange = TableStart.Resize(1, UBound(Headers2) - LBound(Headers2) + 1)
    ReDim Buf(1 To 1, 1 To HdrRange.Columns.Count) As Variant
    SheetName = TableStart.Worksheet.Name
    Set FieldList = New Dictionary
    FieldList.CompareMode = TextCompare
    RemoveIndex
    ReDim Fld(1 To UBound(Headers2) - LBound(Headers2) + 1)
    For I = LBound(Headers2) To UBound(Headers2)
        J = J + 1
        HdrRange.Cells(J).Value = Headers2(I)
        Fld(J).Name = Headers2(I)
        If Len(CStr(Headers2(I))) > 0 Then
            If FieldList.Exists(Headers2(I)) Then
                MsgBox "Duplicate header " & CStr(Headers2(I)) & " on " & TableStart.Worksheet.Name
                HasError = True
                Exit Function
            Else
                FieldList.Add Headers2(I), J
            End If
        End If
    Next I
    Arr = Empty
    ArrPos = 0
    RecordPos = 0
    ArrCount = 0
    Dirty = False
    Create = True
End Function

Function CreateFrom(Xd As XData, Optional ByVal TableStart As Range, _
    Optional VariableWidth As Boolean = False) As Boolean
'If not already initialised, intialises a new range taking the headers from a recordset.
'If already initialised, clears first (except formats  - user must clear if required).
'Creates from RsetData or ArrData object.
'If already initialised and VariableWidth is False, the columns on Xd must match the corresponding
'columns starting from the first - subsequent table columns/formulae are protected.
'If VariableWidth is True, resizes table to width of output and clears surplus headings and table columns.

    Dim Headers() As String
    Dim Result As Boolean
    Dim Wb2 As Workbook
    Dim Ws2 As Worksheet
    Dim Rd As Object
    Dim Ad As Object
    Dim I As Long
    
    If HasError Then Exit Function
    If Xd.HasError Then
        HasError = True
        Exit Function
    End If
    If Not HdrRange Is Nothing Then
        If ListObj Is Nothing Then
            Clear True, , True, True, VariableWidth
        Else
            Clear True, , True, True
            If VariableWidth Then
                For I = 1 To HdrRange.Columns.Count
                    HdrRange(I) = "Temp_" & I
                Next I
            End If
        End If
    End If
        
    If Not HdrRange Is Nothing And Not VariableWidth Then
        If Columns < Xd.Columns Then
            MsgBox "Too few columns"
            HasError = True
            Exit Function
        End If
        For I = 1 To Xd.Columns
            If Fld(I).Name <> Xd.GetHeader(I) Then
                MsgBox "Field " & Fld(I).Name & " does not match source data"
                HasError = True
                Exit Function
            End If
        Next I
        ClearBuf
        RemoveIndex
        Arr = Empty
        ArrPos = 0
        RecordPos = 0
        ArrCount = 0
        Dirty = False
    
    Else
        If TableStart Is Nothing Then Set TableStart = HdrRange.Cells(1, 1)
        ReDim Headers(Xd.Columns - 1)
        For I = 0 To Xd.Columns - 1
            Headers(I) = Xd.GetHeader(I + 1)
        Next I
        If Not Create(TableStart, Headers) Then Exit Function
    End If
    
    If Xd.DataRowCount > 0 Then
        If Not ListObj Is Nothing Then
            ListObj.ShowTotals = False
            TotalsRowCount = 0
        End If
        Select Case Xd.ObjType
        Case "RsetData"
            Set Rd = Xs.GetRd(Xd)
            Rd.Rsett.MoveFirst
            Set Wb2 = ActiveWorkbook
            Set Ws2 = ActiveSheet
            HdrRange.Worksheet.Parent.Activate
            HdrRange.Worksheet.Activate
            HdrRange.Cells(1).Offset(1, 0).CopyFromRecordset Rd.Rsett
            Wb2.Activate
            Ws2.Activate
            
        Case "ArrData"
            Set Ad = Xs.GetAd(Xd)
            For I = 1 To Ad.Columns
                If Not Ad.GetColumnData(I) Then Exit Function
                HdrRange.Cells(1).Offset(1, I - 1).Resize(Ad.DataRowCount) = Ad.GroupData
            Next I
            Ad.GroupData = Empty
        End Select
        
        DataRowCount = Xd.DataRowCount
        Set DataRange = HdrRange.Offset(1).Resize(DataRowCount)
        If Not ListObj Is Nothing Then
            I = ListObj.HeaderRowRange.Columns.Count
            If VariableWidth Then
                ListObj.Resize HdrRange.Resize(DataRowCount + 1)
                If I > UBound(Fld) Then
                    HdrRange.Offset(0, UBound(Fld)).Resize(DataRowCount + 1, I - UBound(Fld)).ClearContents
                End If
            Else
                ListObj.Resize ListObj.HeaderRowRange.Resize(DataRowCount + 1)
            End If
        End If
    End If
    If ShowTotals Then
        ListObj.ShowTotals = True
        TotalsRowCount = 1
    End If
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
    Fieldsx = Fields
    Fields2 = Xs.ExtractArray(Fieldsx)
    CopyFrom = Xs.Copy(Me, Xd, Start, Recs, Fields2)
End Function

Function Clear(Optional Contents As Boolean = True, Optional Formats As ClearOptions = DoNotClear, _
    Optional Borders As Boolean = False, Optional Filters As Boolean = False, _
    Optional Headers As Boolean = False, Optional DataValidation As ClearOptions = DoNotClear, _
    Optional Comments As Boolean = False) As Boolean
'Clears the data.
    
    Dim I As Long

    If HasError Then Exit Function
    If Filters Then
        If ListObj Is Nothing Then
            HdrRange.Worksheet.AutoFilterMode = False
        Else
            ListObj.ShowAutoFilter = False
        End If
    End If
    
    If Comments Then
        If Headers Then HdrRange.ClearComments
        If Not DataRange Is Nothing Then DataRange.ClearComments
    End If
    
    Select Case Formats
    Case ClearAllRows
        If Headers Then HdrRange.ClearFormats
        If Not DataRange Is Nothing Then DataRange.ClearFormats
    Case ClearExceptFirstRow
        If Headers Then HdrRange.ClearFormats
        If DataRowCount > 1 Then DataRange.Offset(1).Resize(DataRowCount - 1).ClearFormats
    End Select
    
    If Borders Then
        If Headers Then
            With HdrRange.Borders
                 .ColorIndex = xlColorIndexNone
                .LineStyle = xlLineStyleNone
            End With
        End If
        If Not DataRange Is Nothing Then
            With DataRange.Borders
                .ColorIndex = xlColorIndexNone
                .LineStyle = xlLineStyleNone
            End With
        End If
    End If
    
    If Not DataRange Is Nothing Then
        Select Case DataValidation
        Case ClearAllRows
            With DataRange.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, _
                    Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
        Case ClearExceptFirstRow
            If Contents And Not ListObj Is Nothing Then
                'On a table, validation is cleared automatically except first row if all rows are removed.
                'Clearing manually seems to clear all validation.
            ElseIf DataRowCount > 1 Then
                With DataRange.Offset(1).Resize(DataRowCount - 1).Validation
                    .Delete
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = True
                End With
            End If
        End Select
        
        If Contents Then
            DataRange.ClearContents
            Set DataRange = Nothing
            If Not ListObj Is Nothing Then
                If Headers Then
                    If ShowTotals Then
                        For I = 2 To ListObj.ListColumns.Count
                            ListObj.ListColumns(I).TotalsCalculation = 0 'xlTotalsCalculationNone
                        Next I
                    End If
                    ListObj.Resize ListObj.HeaderRowRange.Resize(2, 1)
                Else
                    ListObj.Resize ListObj.HeaderRowRange.Resize(2)
                End If
                If DataRowCount > 1 Then _
                    ListObj.HeaderRowRange.Offset(2 + TotalsRowCount).Resize(DataRowCount - 1).ClearContents
                For I = ListObj.ListRows.Count To 1 Step -1
                    ListObj.ListRows(I).Delete
                Next I
            End If
            DataRowCount = 0
            If Not RowIndex Is Nothing Then
                RowIndex.RemoveAll
                RowPending = 0
                BufPending = False
                KeyPending = vbNullString
            End If
            Arr = Empty
            ArrPos = 0
            RecordPos = 0
            ArrCount = 0
            Dirty = False
        End If
    End If
    
    If Headers Then HdrRange.ClearContents
    Clear = True
End Function

Function FormulaFill(Optional AlwaysFill As AlwaysFillOptions = DoNotAlwaysFill) As Boolean
'Fills formulae with those saved during Initialise.
'If AlwaysFill = AlwaysFillWithRow, always creates one row of formulae even when no data records
'(adds a row).
'If AlwaysFill = AlwaysFillWithoutRow (True), always creates one row of formulae even when no data records
'(does not a row).
'If AlwaysFill = DoNotAlwaysFill (False), does nothing when no data records

    Dim ExtraRow As Boolean
    Dim I As Long

    If HasError Then Exit Function
    If DataRowCount = 0 Then
        Select Case AlwaysFill
        Case AlwaysFillWithRow
            AddRow
        Case AlwaysFillWithoutRow
            AddRow
            ExtraRow = True
        Case Else
            FormulaFill = True
            Exit Function
        End Select
    End If
    
    For I = 1 To UBound(Fld)
        With Fld(I)
            If Len(.Formula) > 0 Then
                DataRange.Resize(, 1).Offset(, I - 1).Formula = .Formula
            ElseIf Len(.ArrFormula) > 0 Then
                If Len(.ArrFormula) > 255 Then
                    MsgBox "Array formula too long on " & SheetName
                    HasError = True
                    Exit Function
                End If
                If DataRowCount > 1 Then _
                    DataRange.Resize(1, 1).Offset(1, I - 1) = 1
                DataRange.Resize(1, 1).Offset(, I - 1).FormulaArray = .ArrFormula
                If DataRowCount > 1 Then _
                    DataRange.Resize(1, 1).Offset(, I - 1).AutoFill DataRange.Resize(, 1).Offset(, I - 1), xlFillValues
            End If
        End With
    Next I
    
    If ExtraRow Then
        DataRowCount = 0
        Set DataRange = Nothing
    End If
    FormulaFill = True
End Function

Function NumFormatFill() As Boolean
'Fills numberformats with those saved during initialisation.
'Always fills at least one row.

    Dim I As Long

    If HasError Then Exit Function
    If DataRowCount = 0 Then
        For I = 1 To UBound(Fld)
            HdrRange.Cells(1).Offset(1, I - 1).NumberFormatLocal = Fld(I).NumFormat
        Next I
    Else
        For I = 1 To UBound(Fld)
            DataRange.Resize(, 1).Offset(, I - 1).NumberFormatLocal = Fld(I).NumFormat
        Next I
    End If
    NumFormatFill = True
End Function

Function BorderFill() As Boolean
'Fills vertical borders from those in the header row.

    Dim I As Long

    If HasError Then Exit Function
    If DataRowCount = 0 Then
        BorderFill = True
        Exit Function
    End If
    
    For I = 1 To UBound(Fld)
        With DataRange.Resize(, 1).Offset(, I - 1)
            With .Borders(xlEdgeLeft)
                .Weight = HdrRange(I).Borders(xlEdgeLeft).Weight
                .Color = HdrRange(I).Borders(xlEdgeLeft).Color
                .TintAndShade = HdrRange(I).Borders(xlEdgeLeft).TintAndShade
                .LineStyle = HdrRange(I).Borders(xlEdgeLeft).LineStyle
            End With
            With .Borders(xlEdgeRight)
                .Weight = HdrRange(I).Borders(xlEdgeRight).Weight
                .Color = HdrRange(I).Borders(xlEdgeRight).Color
                .TintAndShade = HdrRange(I).Borders(xlEdgeRight).TintAndShade
                .LineStyle = HdrRange(I).Borders(xlEdgeRight).LineStyle
            End With
        End With
    Next I
    BorderFill = True
End Function

Function FormatFill(Optional Source As Range, Optional Formats As Boolean = True, _
    Optional Validation As Boolean = False, Optional FormatConditions As Boolean = False) As Boolean
'Fills formats and/or validation from those in the first row of data or Source if provided.
'In the case of a table, if there is a source, always fills one row.

    Dim Wb As Workbook
    Dim Ws As Worksheet
    Dim ActCell As Range
    Dim SUpdating As Boolean
    Dim TargetRows As Long
    Dim CondRange As Range
    Dim I As Long
    
    If HasError Then Exit Function
    If Source Is Nothing Then
        TargetRows = DataRowCount - 1
    ElseIf ListObj Is Nothing Then
        TargetRows = DataRowCount
    ElseIf DataRowCount > 1 Then
        TargetRows = DataRowCount
    Else
        TargetRows = 1
    End If
    If TargetRows < 1 Then
        FormatFill = True
        Exit Function
    End If
    SUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Set Wb = ActiveWorkbook
    Set Ws = ActiveSheet
    HdrRange.Worksheet.Activate
    Set ActCell = ActiveCell
    
    With HdrRange.Offset(IIf(Source Is Nothing, 2, 1)).Resize(TargetRows)
        If Formats Then
            .ClearFormats
        End If
        If Source Is Nothing Then
            DataRange.Resize(1).Copy
        Else
            Source.Copy
        End If
        If Formats Then
            .PasteSpecial xlPasteFormats
        End If
        If Validation Then
            .PasteSpecial xlPasteValidation
        End If
        If FormatConditions Then
            If Source Is Nothing Then
                Set CondRange = DataRange.Resize(1)
            Else
                Set CondRange = Source
            End If
            For I = 1 To CondRange.FormatConditions.Count
                CondRange.FormatConditions.Item(I).ModifyAppliesToRange _
                    Union(CondRange.FormatConditions.Item(I).AppliesTo, .Cells)
            Next I
        End If
    End With
    Application.CutCopyMode = False
    ActCell.Select
    Wb.Activate
    Ws.Activate
    Application.ScreenUpdating = SUpdating
    FormatFill = True
End Function

Property Get GetFormula(Field) As String
'Returns formula if any. Places {} round it if an array formula.

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        With Fld(FieldCol(Field))
            If Len(.Formula) > 0 Then
                GetFormula = .Formula
            ElseIf Len(.ArrFormula) > 0 Then
                GetFormula = "{" & .ArrFormula & "}"
            End If
        End With
    End If
End Property

Property Let GetFormula(Field, Formula As String)
'Sets formula. {} indicates an array formula.

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        With Fld(FieldCol(Field))
            If Left(Formula, 1) = "{" And Right(Formula, 1) = "}" Then
                .ArrFormula = Mid(Formula, 2, Len(Formula) - 2)
            Else
                .Formula = Formula
            End If
        End With
    End If
End Property

'****** Index procedures

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
'Creates an index from the fields specified.
'Does not index all blank or empty string keys.
    
    Dim Values As Variant
    Dim JoinFormula As String
    Dim Addr As String
    Dim NullKey As String
    Dim IndexCells As Collection
    Dim Wb As Workbook
    Dim Ws As Worksheet
    Dim I As Long, J As Integer
    
    If HasError Then Exit Function
    ReDim IndexFields(UBound(Fields2) - LBound(Fields2)) As Long
    RowPending = 0
    BufPending = False
    KeyPending = vbNullString
    
    Set Wb = ActiveWorkbook
    Set Ws = ActiveSheet
    If Not Wb Is HdrRange.Worksheet.Parent Then
        HdrRange.Worksheet.Parent.Activate
    End If
    For I = LBound(Fields2) To UBound(Fields2)
        If Not FieldCheck(Fields2(I)) Then
            Set RowIndex = Nothing
            IndexFields = Empty
            GoTo Finish
        End If
        IndexFields(J) = FieldCol(Fields2(I))
        If DataRowCount > 0 Then
            If I > LBound(Fields2) Then JoinFormula = JoinFormula & "&""" & ChrW(SepChar) & """&"
            Addr = DataRange.Columns(FieldCol(Fields2(I))).Address(True, True)
            JoinFormula = JoinFormula & Addr
        End If
        J = J + 1
    Next I
    
    Set RowIndex = New Dictionary
    RowIndex.CompareMode = TextCompare
    Select Case DataRowCount
    Case 0
        CreateIndex2 = True
        GoTo Finish
    Case 1
        ReDim Values(1 To 1, 1 To 1)
        Values(1, 1) = HdrRange.Worksheet.Evaluate(JoinFormula)
    Case Else
        Values = HdrRange.Worksheet.Evaluate(JoinFormula)
    End Select
    
    If IsError(Values) Then
        MsgBox "Cannot create index"
        HasError = True
        RemoveIndex
        GoTo Finish
    End If
    
    NullKey = String(UBound(IndexFields), ChrW(SepChar))
    For I = 1 To UBound(Values)
        If CStr(Values(I, 1)) <> NullKey Then
            If RowIndex.Exists(CStr(Values(I, 1))) Then
                RowIndex(CStr(Values(I, 1))).Add DataRange.Cells(I, IndexFields(0))
            Else
                Set IndexCells = New Collection
                IndexCells.Add DataRange.Cells(I, IndexFields(0))
                RowIndex.Add CStr(Values(I, 1)), IndexCells
            End If
        End If
    Next I
    CreateIndex2 = True
    
Finish:
    If Not Wb Is ActiveWorkbook Then
        Wb.Activate
        Ws.Activate
    End If
End Function

Public Function RemoveIndex() As Boolean

    RemoveIndex = XData_RemoveIndex()
End Function

Private Function XData_RemoveIndex() As Boolean
'Removes index.

    If HasError Then Exit Function
    Set RowIndex = Nothing
    IndexFields = Empty
    RowPending = 0
    BufPending = False
    KeyPending = vbNullString
    XData_RemoveIndex = True
End Function

Function GetIndex() As Variant
'Returns entire index as a 2-dimensional zero-based array of type long, entries x fields

    Dim Keys As Variant
    Dim Values As String
    Dim SplitValues As Variant
    Dim Result As Variant
    Dim I As Long, J As Long

    GetIndex = Xs.ExtractArray
    If HasError Then Exit Function
    If RowIndex Is Nothing Then Exit Function
    FlushIndex
    If RowIndex.Count = 0 Then Exit Function
    Keys = RowIndex.Keys
    ReDim Result(UBound(Keys), UBound(IndexFields))
    For I = 0 To UBound(Keys)
        Values = Keys(I)
        SplitValues = Split(Values, ChrW(SepChar))
        For J = 0 To UBound(IndexFields)
            Result(I, J) = SplitValues(J)
        Next J
    Next I
    GetIndex = Result
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
'Checks for invalid entries and removes from index if found (removed).

    Dim Key As String
    Dim IndexCells As Collection
    Dim Lookup3() As Long
    Dim Row As Long
    Dim I As Long, J As Long
    
    Lookup2 = Xs.ExtractArray
    If HasError Then Exit Function
    If RowIndex Is Nothing Then Exit Function
    FlushIndex
    For I = LBound(Values2) To UBound(Values2)
        If I > LBound(Values2) Then
            Key = Key & (ChrW(SepChar) & CStr(Values2(I)))
        Else
            Key = CStr(Values2(I))
        End If
    Next I
    If Not RowIndex.Exists(Key) Then Exit Function
    Set IndexCells = RowIndex(Key)
    ReDim Lookup3(IndexCells.Count - 1)
    For I = 1 To IndexCells.Count
        Lookup3(I - 1) = IndexCells(I).Row - HdrRange.Row
    Next I
    Lookup2 = Lookup3
    
'    I = 1
'    Do While I <= IndexCells.Count
'        Row = IndexCells(I).Row - DataRange.Row + 1
'        If GetKey(Row) = Key Then
'            lookup3(J) = Row
'            I = I + 1
'            J = J + 1
'        Else
'            IndexCells.Remove I
'        End If
'    Loop
'    If IndexCells.Count = 0 Then RowIndex.Remove Key
'    If J > 0 Then
'        ReDim Preserve lookup3(J - 1)
'        Lookup = lookup3
'    End If
End Function

Private Sub FlushIndex()
'If any row has pending index changes, update index with these.

    Dim NewKey As String
    
    If RowIndex Is Nothing Then Exit Sub
    If RowPending = 0 Then Exit Sub
    NewKey = GetKey(RowPending, BufPending)
    If NewKey <> KeyPending Then
        DelIndex RowPending, KeyPending
        InsIndex RowPending, NewKey
    End If
    RowPending = 0
    BufPending = False
    KeyPending = vbNullString
End Sub

#If LiveMode Then
Private Function DelIndex(Row As Long, Key As String) As Boolean
#Else
Function DelIndex(Row As Long, Key As String) As Boolean
#End If
'Deletes the index entry for a row.
'Returns True if anything was found/deleted.

    Dim IndexCells As Collection
    Dim LoBound As Long
    Dim HiBound As Long
    Dim MidPoint As Long
    
    If RowIndex Is Nothing Then Exit Function
    If Not RowIndex.Exists(Key) Then Exit Function
    Set IndexCells = RowIndex(Key)
    LoBound = 1
    HiBound = IndexCells.Count
    Do While HiBound >= LoBound
        MidPoint = LoBound + (HiBound - LoBound) \ 2
        Select Case IndexCells(MidPoint).Row - HdrRange.Row
        Case Row
            IndexCells.Remove MidPoint
            If IndexCells.Count = 0 Then RowIndex.Remove Key
            DelIndex = True
            Exit Function
        Case Is > Row
            HiBound = MidPoint - 1
        Case Else
            LoBound = MidPoint + 1
        End Select
    Loop
End Function

#If LiveMode Then
Private Function InsIndex(Row As Long, Key As String) As Boolean
#Else
Function InsIndex(Row As Long, Key As String) As Boolean
#End If
'Inserts the index entry for a row.
'Returns True if anything was inserted.

    Dim IndexCells As Collection
    Dim LoBound As Long
    Dim HiBound As Long
    Dim MidPoint As Long
    
    If RowIndex Is Nothing Then Exit Function
    If Key = String(UBound(IndexFields), ChrW(SepChar)) Then Exit Function
    If RowIndex.Exists(Key) Then
        Set IndexCells = RowIndex(Key)
        LoBound = 1
        HiBound = IndexCells.Count
        Do While HiBound >= LoBound
            MidPoint = LoBound + (HiBound - LoBound) \ 2
            Select Case IndexCells(MidPoint).Row - HdrRange.Row
            Case Row
                Exit Function
            Case Is > Row
                HiBound = MidPoint - 1
            Case Else
                LoBound = MidPoint + 1
            End Select
        Loop
        If HiBound > 0 Then
            IndexCells.Add HdrRange.Cells(1, IndexFields(0)).Offset(Row), After:=HiBound
        Else
            IndexCells.Add HdrRange.Cells(1, IndexFields(0)).Offset(Row), Before:=LoBound
        End If
    Else
        Set IndexCells = New Collection
        IndexCells.Add HdrRange.Cells(1, IndexFields(0)).Offset(Row)
        RowIndex.Add Key, IndexCells
    End If
    InsIndex = True
End Function

#If Not LiveMode Then
Sub PrintIndex()
'For testing

    Dim Keys As Variant
    Dim IndexCells As Collection
    Dim Values As String
    Dim I As Long, J As Long
    
    If RowIndex Is Nothing Then Exit Sub
    FlushIndex
    Keys = RowIndex.Keys
    For I = 0 To UBound(Keys)
        Set IndexCells = RowIndex(Keys(I))
        Values = Keys(I)
        For J = 1 To IndexCells.Count
            Values = Values & " " & IndexCells(J).Row - HdrRange.Row
        Next J
        Debug.Print Values
    Next I
    Debug.Print
End Sub
#End If

#If LiveMode Then
Private Function GetKey(Row As Long, UseBuf As Boolean) As String
#Else
Function GetKey(Row As Long, UseBuf As Boolean) As String
#End If
'Returns the key for a given row

    Dim Key As String
    Dim I As Long
    
    If RowIndex Is Nothing Then Exit Function
    If UseBuf Then
        For I = 0 To UBound(IndexFields)
            If I > 0 Then Key = Key & ChrW(SepChar)
            Key = Key & CStr(Arr(RecordPos - ArrPos + 1, IndexFields(I)))
        Next I
    Else
        For I = 0 To UBound(IndexFields)
            If I > 0 Then Key = Key & ChrW(SepChar)
            Key = Key & CStr(DataRange.Cells(Row, IndexFields(I)).Value)
        Next I
    End If
    GetKey = Key
End Function

'******Header procedures

Public Function CheckHeaders(ParamArray Headers()) As Boolean
'Checks if headers exist.
'Parameter is a variant containing an array (for use with Array function), a list or a single value.
'Returns True if check is successful.

    Dim Headers2 As Variant
    Dim Headersx As Variant
    Dim I As Integer
    
    If HasError Then Exit Function
    Headersx = Headers
    Headers2 = Xs.ExtractArray(Headersx)
    For I = LBound(Headers2) To UBound(Headers2)
        If Len(CStr(Headers2(I))) > 0 Then
            If Not FieldCheck(CStr(Headers2(I))) Then Exit Function
        End If
    Next I
    CheckHeaders = True
End Function

Public Property Get GetHeader(Index As Variant) As Variant

    If HasError Then Exit Property
    If FieldCheck(Index) Then GetHeader = FieldName(Index)
End Property

Public Property Let GetHeader(Index As Variant, HdrValue As Variant)
'Sets a header value.

    Dim HdrCol As Long
    Dim HdrVal As String

    If HasError Then Exit Property
    If FieldCheck(Index) Then
        HdrCol = FieldCol(Index)
        HdrVal = Trim(CStr(HdrValue))
        If Fld(HdrCol).Name = HdrVal Then Exit Property
        If FieldList.Exists(HdrVal) Then
            MsgBox "Duplicate header " & HdrVal
            HasError = True
            Exit Property
        End If
        If Len(Fld(HdrCol).Name) > 0 Then
            If FieldList.Exists(Fld(HdrCol).Name) Then
                FieldList.Remove Fld(HdrCol).Name
            End If
        End If
        Fld(HdrCol).Name = HdrValue
        If Len(HdrValue) > 0 Then FieldList.Add HdrValue, HdrCol
    End If
End Property

Private Function XData_GetHeader(Index As Variant) As Variant
'Returns single header (Index is 1-based)
    
    If HasError Then Exit Function
    If FieldCheck(Index) Then XData_GetHeader = FieldName(Index)
End Function

Property Get GetHeaders() As Variant
'Returns non-blank headers as a comma-separated string

    Dim Hdr As Variant
    
    If HasError Then Exit Property
    For Each Hdr In FieldList.Keys
        GetHeaders = GetHeaders & """" & Hdr & ""","
    Next Hdr
    If Len(GetHeaders) > 0 Then GetHeaders = Left(GetHeaders, Len(GetHeaders) - 1)
End Property

Property Let GetHeaders(Headers As Variant)
'Sets headers from an array of strings

    Dim I As Long
    
    If Not IsArray(Headers) Then
        MsgBox "Not an array"
        HasError = True
        Exit Property
    ElseIf UBound(Headers) - LBound(Headers) <> UBound(Fld) - 1 Then
        MsgBox "Array wrong size"
        HasError = True
        Exit Property
    End If
    
    Set FieldList = New Dictionary
    FieldList.CompareMode = TextCompare
    For I = 1 To UBound(Fld)
        Fld(I).Name = vbNullString
        GetHeader(I) = Headers(LBound(Headers) + I - 1)
        If HasError Then Exit Property
    Next I
End Property

Function GetTypes() As String
'Returns data types as a comma-separated string based on first row

    Dim TypeStr As String
    Dim I As Long
    
    If HasError Then Exit Function
    For I = 1 To UBound(Fld)
        Select Case Fld(I).VarType
        Case vbInteger, vbLong
            TypeStr = "vbLong"
        Case vbSingle, vbDouble
            TypeStr = "vbDouble"
        Case vbCurrency
            TypeStr = "vbCurrency"
        Case vbDate
            TypeStr = "vbDate"
        Case vbString
            TypeStr = "vbString"
        Case vbBoolean
            TypeStr = "vbBoolean"
        Case Else
            TypeStr = "unknown"
        End Select
        GetTypes = GetTypes & TypeStr & ","
        Next I
    If Len(GetTypes) > 0 Then GetTypes = Left(GetTypes, Len(GetTypes) - 1)

End Function

Function GetColHeader(Field) As Range
'Gets the column header of a field as a range.

    If HasError Then Exit Function
    If FieldCheck(Field) Then
        Set GetColHeader = HdrRange.Cells(1, FieldCol(Field))
    End If
End Function

'******Column procedures

Function AddColumn(Field As String) As Boolean
'Adds a column at the end (assumes space is available, does not shift).
'Returns True if successful
    
    If HasError Then Exit Function
    If FieldList.Exists(Field) Then
        MsgBox "Duplicate field " & Field & " on " & SheetName
        HasError = True
        Exit Function
    End If
    HdrRange.Resize(1, 1).Offset(, HdrRange.Columns.Count).Value = Field
    Set HdrRange = HdrRange.Resize(1, HdrRange.Columns.Count + 1)
    ReDim Buf(1 To 1, 1 To HdrRange.Columns.Count) As Variant
    If Not DataRange Is Nothing Then _
        Set DataRange = DataRange.Resize(, HdrRange.Columns.Count)
    FieldList.Add Field, HdrRange.Columns.Count
    ReDim Preserve Fld(1 To HdrRange.Columns.Count)
    Fld(HdrRange.Columns.Count).Name = Field
    AddColumn = True
End Function

Function InsertColumn(Field As String, Optional Pos As Variant, Optional EntireColumn As Boolean = False) As Boolean
'Inserts a column at position Pos with name Field and shifts anything on the right to the right.
'If Pos is missing, inserts at the end.
'Pos may be a field name or number.
'Returns True if successful.

    Dim FieldPos As Long
    Dim I As Integer
    
    If HasError Then Exit Function
    If FieldList.Exists(Field) Then
        MsgBox "Duplicate field " & Field & " on " & SheetName
        HasError = True
        Exit Function
    End If
    If IsMissing(Pos) Then Pos = HdrRange.Columns.Count + 1
    If Not IsNumeric(Pos) Then
        If Not FieldCheck(Pos) Then Exit Function
    ElseIf Pos <> HdrRange.Columns.Count + 1 Then
        If Not FieldCheck(Pos) Then Exit Function
    End If
    FieldPos = FieldCol(Pos)
    If EntireColumn Then
        HdrRange.Resize(1, 1).Offset(, FieldPos - 1).EntireColumn.Insert xlShiftToRight
    ElseIf Not ListObj Is Nothing Then
        ListObj.ListColumns.Add FieldPos
    Else
        HdrRange.Resize(DataRowCount + TotalsRowCount + 1, 1).Offset(, FieldPos - 1).Insert xlShiftToRight
    End If
    Select Case FieldPos
    Case 1
        Set HdrRange = HdrRange.Offset(0, -1).Resize(1, HdrRange.Columns.Count + 1)
    Case Is > HdrRange.Columns.Count
        Set HdrRange = HdrRange.Resize(1, HdrRange.Columns.Count + 1)
    End Select
    ReDim Buf(1 To 1, 1 To HdrRange.Columns.Count) As Variant
    If DataRowCount > 0 Then _
        Set DataRange = HdrRange.Offset(1, 0).Resize(DataRowCount)
    HdrRange.Resize(1, 1).Offset(, FieldPos - 1).Value = Field
    For I = 0 To FieldList.Count - 1
        If FieldList.Items(I) >= FieldPos Then
            FieldList(FieldList.Keys(I)) = FieldList.Items(I) + 1
        End If
    Next I
    FieldList.Add Field, FieldPos
    ReDim Preserve Fld(1 To HdrRange.Columns.Count)
    For I = HdrRange.Columns.Count - 1 To FieldPos Step -1
        Fld(I + 1) = Fld(I)
    Next I
    With Fld(FieldPos)
        .Name = Field
        .Formula = vbNullString
        .ArrFormula = vbNullString
        .IsText = False
    End With
    
    If Not RowIndex Is Nothing Then
        For I = 0 To UBound(IndexFields)
            If IndexFields(I) >= FieldPos Then IndexFields(I) = IndexFields(I) + 1
        Next I
    End If
    InsertColumn = True
End Function

Function DeleteColumn(Field, Optional EntireColumn As Boolean = False) As Boolean
'Deletes a column and shifts anything on the right to the left.
'Returns True if successful.

    Dim I As Integer

    If HasError Then Exit Function
    If Not FieldCheck(Field) Then Exit Function
    If HdrRange.Columns.Count <= 1 Then
        MsgBox "Cannot delete last column" & " on " & SheetName
        HasError = True
        Exit Function
    End If
    If Not RowIndex Is Nothing Then
        For I = 0 To UBound(IndexFields)
            Select Case IndexFields(I)
            Case FieldCol(Field)
                RemoveIndex
                Exit For
            Case Is > FieldCol(Field)
                IndexFields(I) = IndexFields(I) - 1
            End Select
        Next I
    End If
    If EntireColumn Then
        HdrRange(FieldCol(Field)).EntireColumn.Delete xlShiftToLeft
    ElseIf Not ListObj Is Nothing Then
        ListObj.ListColumns(FieldCol(Field)).Delete
    Else
        HdrRange(FieldCol(Field)).Resize(DataRowCount + TotalsRowCount + 1).Delete xlShiftToLeft
    End If
    ReDim Buf(1 To 1, 1 To HdrRange.Columns.Count) As Variant
    For I = FieldCol(Field) To HdrRange.Columns.Count
        Fld(I) = Fld(I + 1)
    Next I
    ReDim Preserve Fld(1 To HdrRange.Columns.Count)
    For I = 0 To FieldList.Count - 1
        If FieldList.Items(I) > FieldCol(Field) Then
            FieldList(FieldList.Keys(I)) = FieldList.Items(I) - 1
        End If
    Next I
    FieldList.Remove FieldName(Field)
    DeleteColumn = True
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
    Columns = HdrRange.Columns.Count
End Function

Private Function XData_Columns() As Integer
'Returns the number of columns.

    If HasError Then Exit Function
    XData_Columns = HdrRange.Columns.Count
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
    Case vbInteger, vbLong, vbDouble
        XData_Exists = Field >= 1 And Field <= UBound(Fld)
    End Select
End Function

Function GetColumn(Field) As Range
'Gets the column data of a field as a range.

    If HasError Then Exit Function
    If FieldCheck(Field) And Not DataRange Is Nothing Then
        Set GetColumn = DataRange.Columns(FieldCol(Field))
    End If
End Function

'****** Row procedures

Function AddRow(Optional NumRows As Long = 1) As Long
'Logically adds a row or rows at the end and returns the number of the last row.
'Does not physically change the worksheet, except to add table rows if applicable.
'If there is an error returns -1.

    Dim I As Long

    If HasError Then
        AddRow = -1
        Exit Function
    End If
    If NumRows <= 0 Then Exit Function
    DataRowCount = DataRowCount + NumRows
    AddRow = DataRowCount
    If Not ListObj Is Nothing Then
        If NumRows = 1 And DataRowCount = 1 Then
            ListObj.ListRows.Add
        Else
            If ListObj.ListRows.Count = 0 Then ListObj.ListRows.Add '(doesn't work otherwise)
            ListObj.Resize ListObj.HeaderRowRange.Resize(DataRowCount + TotalsRowCount + 1)
        End If
    End If
    Set DataRange = HdrRange.Offset(1, 0).Resize(DataRowCount)
End Function

Function DropRow(Optional NumRows As Long = 1) As Long
'Drops the last NumRows rows.
'If a table, resizes it but otherwise does not change the worksheet.
'Must leave at least one data row if a table with data in first row
'Returns the number of rows (or -1 if an error)
'Note table complexity is because resizing a table to one data row which is empty
'may leave an empty table with no data rows or DataBodyRange.

    Dim HasData As Boolean
    Dim I As Long

    If HasError Then
        DropRow = -1
        Exit Function
    End If
    If NumRows <= 0 Then Exit Function
    If DataRowCount < NumRows Then NumRows = DataRowCount
    DataRowCount = DataRowCount - NumRows
    DropRow = DataRowCount
    If DataRowCount = 0 Then
        Set DataRange = Nothing
    Else
        Set DataRange = HdrRange.Offset(1, 0).Resize(DataRowCount)
    End If
    If Not ListObj Is Nothing Then
        With ListObj
            Select Case DataRowCount
            Case 0
                GoSub CheckData
                If HasData Then
                    MsgBox "Cannot remove row below table header while leaving data in place"
                    HasError = True
                    DropRow = -1
                Else
                    .Resize .HeaderRowRange.Resize(2)
                    .HeaderRowRange.Offset(1)(1) = 0
                    .ListRows(1).Delete
                End If
            Case 1
                GoSub CheckData
                .Resize ListObj.HeaderRowRange.Resize(2)
                If Not HasData Then
                    .HeaderRowRange.Offset(1)(1) = 0
                    .HeaderRowRange.Offset(1)(1).ClearContents
                End If
            Case Else
                .Resize ListObj.HeaderRowRange.Resize(DataRowCount + 1)
            End Select
        End With
    End If
    Exit Function
    
CheckData:
        For I = 1 To ListObj.ListColumns.Count
            If Not IsEmpty(ListObj.HeaderRowRange.Offset(1)(I)) Then
                HasData = True
                Exit For
            End If
        Next I
        Return
End Function

Function InsertRows(Optional NumRows As Long = 1, Optional ByVal Pos As Long = -1, Optional EntireRow As Boolean = False) As Long
'Physically inserts a row or rows and returns the number of the last row (-1 if an error).
'Pos specifies the relative position of the new rows.

    Dim I As Long

    If HasError Then
        InsertRows = -1
        Exit Function
    End If
    If Pos <= 0 Then Pos = DataRowCount + 1
    If NumRows <= 0 Then
        InsertRows = DataRowCount
        Exit Function
    ElseIf Pos > DataRowCount + 1 Then
        MsgBox "Attempting to insert beyond last row of data"
        HasError = True
        InsertRows = -1
        Exit Function
    End If
    FlushIndex
    If EntireRow Then
        HdrRange.EntireRow.Offset(Pos, 0).Resize(NumRows).Insert xlShiftDown
    ElseIf Not ListObj Is Nothing Then
        ListObj.HeaderRowRange.Offset(Pos, 0).Resize(NumRows).Insert xlShiftDown
    Else
        HdrRange.Offset(Pos, 0).Resize(NumRows).Insert xlShiftDown
    End If
    If Pos = DataRowCount + 1 Then
        If Not ListObj Is Nothing Then
            If ListObj.ListRows.Count < DataRowCount + NumRows Then ListObj.ListRows.Add
            If NumRows > 1 Then
                ListObj.Resize ListObj.HeaderRowRange.Resize(DataRowCount + NumRows + TotalsRowCount + 1)
            End If
        End If
    End If
    DataRowCount = DataRowCount + NumRows
    Set DataRange = HdrRange.Offset(1, 0).Resize(DataRowCount)
    InsertRows = DataRowCount
End Function

Function DeleteRows(Optional NumRows As Long = 1, Optional ByVal Pos As Long = -1, Optional EntireRow As Boolean = False) As Long
'Deletes rows and returns number of rows (-1 if an error).

    Dim I As Long

    If HasError Then
        DeleteRows = -1
        Exit Function
    End If
    If Pos <= 0 Then Pos = DataRowCount - NumRows + 1
    If NumRows <= 0 Then
        DeleteRows = DataRowCount
        Exit Function
    End If
    If DataRowCount = 0 Then
        MsgBox "No rows to delete"
        HasError = True
        DeleteRows = -1
        Exit Function
    ElseIf Pos + NumRows > DataRowCount + 1 Then
        MsgBox "Attempting to delete beyond last row of data"
        HasError = True
        DeleteRows = -1
        Exit Function
    End If
    If Not RowIndex Is Nothing Then
        FlushIndex
        For I = 1 To NumRows
            DelIndex Pos + I - 1, GetKey(Pos + I - 1, False)
        Next I
    End If
    If EntireRow Then
        HdrRange.EntireRow.Offset(Pos).Resize(NumRows).Delete xlShiftUp
    ElseIf Not ListObj Is Nothing Then
        ListObj.HeaderRowRange.Offset(Pos).Resize(NumRows).Delete xlShiftUp
    Else
        HdrRange.Offset(Pos).Resize(NumRows).Delete xlShiftUp
    End If
    DataRowCount = DataRowCount - NumRows
    If DataRowCount <= 0 Then
        Set DataRange = Nothing
    Else
        Set DataRange = DataRange.Resize(DataRowCount)
    End If
    DeleteRows = DataRowCount
End Function

Function GetRow(Optional ByVal Row As Long = -1) As Range
'Returns row as a range.
    
    If HasError Then Exit Function
    If DataRowCount = 0 Then Exit Function
    If Row < 0 Then Row = DataRowCount
    Set GetRow = DataRange.Rows(Row)
End Function

'****** Cell procedures

Function GetCell(Field, Optional ByVal Row As Long = -1) As Range
'Returns cell as a range.

    If HasError Then Exit Function
    If FieldCheck(Field) And DataRowCount > 0 Then
        If Row < 0 Then Row = DataRowCount
        Set GetCell = DataRange.Cells(Row, FieldCol(Field))
    End If
End Function

Property Get Value(Field, ByVal Row As Long, Optional Columns As Integer = 1) As Variant
Attribute Value.VB_UserMemId = 0
'Gets value of cell, or range of cells Columns wide.
'Default property
'Attribute Value.VB_UserMemId = 0

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Value = DataRange.Cells(Row, FieldCol(Field)).Resize(, Columns)
    End If
End Property

Property Let Value(Field, Optional ByVal Row As Long = -1, Optional Columns As Integer = 1, ByVal FieldValue As Variant)
'Sets value of cell, or range of cells Columns wide.
'(if cell is empty depending on compiler constant)
'If row is omitted, sets last.
'Note, if setting multiple cells, only the first is checked as empty.
'Places ' character before strings if necessary

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        If Row < 0 Then Row = DataRowCount
        If Not RowIndex Is Nothing Then
            If Row <> RowPending Then
                FlushIndex
                RowPending = Row
                BufPending = False
                KeyPending = GetKey(Row, False)
            End If
        End If
        
        If VarType(FieldValue) = vbString Then
            If Left(FieldValue, 1) = "=" Or Left(FieldValue, 1) = "'" Or IsNumeric(FieldValue) Or IsDate(FieldValue) Then
                FieldValue = "'" & FieldValue
            End If
        End If
        
#If ValueOverwrite Then
        DataRange.Cells(Row, FieldCol(Field)).Resize(, Columns) = FieldValue
#Else
        If IsEmpty(DataRange.Cells(Row, FieldCol(Field))) Then
            DataRange.Cells(Row, FieldCol(Field)).Resize(, Columns) = FieldValue
        End If
#End If

    End If
End Property

Function Text(Field, ByVal Row As Long) As Variant
'Gets text of cell.

    If HasError Then Exit Function
    If FieldCheck(Field) Then
        Text = DataRange.Cells(Row, FieldCol(Field)).Text
    End If
End Function

'****** Buffer procedures

Public Property Get Buffer(Field) As Variant
'Gets value in buffer for specified field.

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Buffer = Buf(1, FieldCol(Field))
    End If
End Property

Public Property Let Buffer(Field, FieldValue As Variant)
'Sets value in buffer for specified field.

    If HasError Then Exit Property
    If FieldCheck(Field) Then
         Buf(1, FieldCol(Field)) = FieldValue
    End If
End Property

Public Function ReadBuf(Row As Long) As Boolean
'Reads requested row into buffer.

    If HasError Then Exit Function
    If HdrRange.Columns.Count = 1 Then
        ReDim Buf(1 To 1, 1 To 1)
        Buf(1, 1) = DataRange.Rows(Row).Value
    Else
        Buf = DataRange.Rows(Row).Value
    End If
    ReadBuf = True
End Function

Public Function WriteBuf(Optional ByVal Row As Long = -1, Optional Clear As Boolean = True) As Boolean
'Writes requested row from buffer.
'Places ' character before strings if necessary

    Dim OldKey As String
    Dim NewKey As String
    Dim Buf2 As Variant
    Dim I As Long
    
    If HasError Then Exit Function
    If Row < 0 Then Row = DataRowCount
    FlushIndex
    OldKey = GetKey(Row, False)
    Buf2 = Buf
    For I = 1 To UBound(Buf2, 2)
        If VarType(Buf2(1, I)) = vbString Then
            If Left(Buf2(1, I), 1) = "=" Or Left(Buf2(1, I), 1) = "'" Or IsNumeric(Buf2(1, I)) Or IsDate(Buf2(1, I)) Then
                Buf2(1, I) = "'" & Buf2(1, I)
            End If
        End If
    Next I
    DataRange.Rows(Row).Value = Buf2
    NewKey = GetKey(Row, False)
    If NewKey <> OldKey Then
        DelIndex Row, OldKey
        InsIndex Row, NewKey
    End If
    If Clear Then ReDim Buf(1 To 1, 1 To HdrRange.Columns.Count) As Variant
    WriteBuf = True
End Function

Public Function CopyBuf(Td As TableData) As Boolean
'Copies buffer from buffer in Td

    Dim I As Long
    
    If HasError Then Exit Function
    For I = 1 To UBound(Fld)
        Buffer(I) = Td.Buffer(FieldName(I))
        If Td.HasError Then
            HasError = True
            Exit Function
        End If
    Next I
    CopyBuf = True
End Function

Public Function ClearBuf() As Boolean
'Clears buffer.

    If HasError Then Exit Function
    ReDim Buf(1 To 1, 1 To HdrRange.Columns.Count) As Variant
    ClearBuf = True
End Function

'******Buffered input/output procedures

Public Function RecEOF() As Boolean
'True if at EOF while reading data in record mode.

    If HasError Then Exit Function
    RecEOF = (RecordPos > DataRowCount)
End Function

Public Function RecStart() As Boolean

    RecStart = XData_RecStart
End Function

Private Function XData_RecStart() As Boolean
'Moves to BOF position while reading data in record mode and returns True, or False if no records.
'Also can be used to initialise record adding if records had previously been read.
    
    If HasError Then Exit Function
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
    FlushIndex
    RecordPos = RecordPos + 1
    If RecordPos <= DataRowCount Then
        If RecordPos > (ArrPos + ArrSize - 1) Or ArrPos = 0 Then
            RecSave
            RecLoad
        End If
        XData_RecNext = True
    End If
End Function

Public Property Get RecPos() As Long

    If HasError Then Exit Property
    RecPos = RecordPos
End Property

Private Property Get XData_RecPos() As Long
'Returns record position while reading data in record mode

    If HasError Then Exit Property
    XData_RecPos = RecordPos
End Property

Public Property Let RecPos(Pos As Long)

    XData_RecPos = Pos
End Property

Private Property Let XData_RecPos(Pos As Long)
'Sets record position while reading data in record mode

    If HasError Then Exit Property
    If Pos > DataRowCount Then
        MsgBox "Attempting to set record position beyond end of table"
        HasError = True
    ElseIf Pos <= 0 Then
        RecStart
    ElseIf Pos > (ArrPos + ArrSize - 1) Or Pos < ArrPos Or ArrPos = 0 Then
        FlushIndex
        RecSave
        RecordPos = Pos
        RecLoad
    Else
        FlushIndex
        RecordPos = Pos
    End If
End Property

Public Function RecAdd() As Boolean

    RecAdd = XData_RecAdd
End Function

Private Function XData_RecAdd() As Boolean
'Adds a record while writing data in record mode.
'Call before putting values in the record.
    
    If HasError Then Exit Function
    FlushIndex
    If ArrPos = 0 Then
        If Not ListObj Is Nothing Then
            ListObj.ShowTotals = False
            TotalsRowCount = 0
        End If
    End If
    
    If ArrCount >= ArrSize Or ArrPos = 0 Then
        RecSave
        RecordPos = DataRowCount + 1
        RecLoad
    End If
    ArrCount = ArrCount + 1
    RecordPos = ArrCount + ArrPos - 1

    If DataRowCount = 0 Then
        AddRow 'otherwise initial row damages index.
    End If
    XData_RecAdd = True
End Function

Private Function XData_RecDelete() As Boolean

    If HasError Then Exit Function
    XData_RecClose
    DeleteRows 1, RecordPos
    If HasError Then Exit Function
    XData_RecDelete = True
End Function

Public Function RecClose(Optional Cancel As Boolean = False) As Boolean
'Cancel is not implemented.

    RecClose = XData_RecClose(Cancel)
End Function

Private Function XData_RecClose(Optional Cancel As Boolean = False) As Boolean
'Flush any remaining data while writing data in record mode.
'Cancel is not implemented.

    If HasError Then Exit Function
    FlushIndex
    RecSave

    If ShowTotals Then
        ListObj.ShowTotals = True
        TotalsRowCount = 1
    End If
    XData_RecClose = True
End Function

Property Get RecValue(Field) As Variant
'Returns data from current record while reading data in record mode.

    Dim I As Long
    Dim Fld As Long

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Fld = FieldCol(Field)
        I = RecordPos - ArrPos + 1
        RecValue = Arr(I, Fld)
    End If
End Property

Property Get XData_RecValue(Field) As Variant
'Returns data from current record while reading data in record mode.

    Dim I As Long
    Dim Fld As Long

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Fld = FieldCol(Field)
        I = RecordPos - ArrPos + 1
        XData_RecValue = Arr(I, Fld)
    End If
End Property

Private Property Get XData_RecValue2(Field As Variant) As Variant
'Returns data from current record while reading data in record mode.
'Assumes Field is a valid column number.
'Removes ' from text value which would otherwise be a formula, number or date.

    Dim I As Long

    If HasError Then Exit Property
    I = RecordPos - ArrPos + 1
    XData_RecValue2 = Arr(I, Field)
End Property

Property Let RecValue(Field, FieldValue As Variant)
'Sets value in current record while reading data in record mode.

    Dim Field2 As Long

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Field2 = FieldCol(Field)
        XData_RecValue2(Field2) = FieldValue
    End If
End Property

Property Let XData_RecValue(Field, FieldValue As Variant)
'Sets value in current record while reading data in record mode.

    Dim Field2 As Long

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Field2 = FieldCol(Field)
        XData_RecValue2(Field2) = FieldValue
    End If
End Property

Private Property Let XData_RecValue2(Field As Variant, FieldValue As Variant)
'Sets value in current record while reading data in record mode.
'Assumes Field is a valid column number.
'Adds ' to text value which would otherwise be a formula, number or date.

    Dim I As Long

    If HasError Then Exit Property
    If Not RowIndex Is Nothing Then
        If RecordPos <> RowPending Then
            RowPending = RecordPos
            BufPending = True
            KeyPending = GetKey(RecordPos, True)
        End If
    End If
    
    I = RecordPos - ArrPos + 1
    Arr(I, Field) = FieldValue
    Dirty = True
End Property

Private Sub RecSave()
'Write data from Buf if necessary.
'ArrCount = the number of records in Buf.
'LastPos = position of DataRowCount in Buf.
'Puts ' in front of strings which would otherwise be incorrectly written to worksheet.

    Dim Arr2 As Variant
    Dim LastPos As Long
    Dim I As Long, J As Long
    
    If ArrPos > 0 Then
        LastPos = DataRowCount - ArrPos + 1
        AddRow ArrCount - LastPos
        If Dirty Then
            If ArrCount >= ArrSize Then
                For I = 1 To ArrSize
                    For J = 1 To HdrRange.Columns.Count
                        If VarType(Arr(I, J)) = vbString Then
                            If Left(Arr(I, J), 1) = "=" Or Left(Arr(I, J), 1) = "'" Or IsNumeric(Arr(I, J)) Or IsDate(Arr(I, J)) Then
                                Arr(I, J) = "'" & Arr(I, J)
                            End If
                        End If
                    Next J
                Next I
                HdrRange.Offset(ArrPos).Resize(ArrSize).Value = Arr
            Else
                ReDim Arr2(1 To ArrCount, 1 To HdrRange.Columns.Count)
                For I = 1 To ArrCount
                    For J = 1 To HdrRange.Columns.Count
                        If VarType(Arr(I, J)) = vbString Then
                            If Left(Arr(I, J), 1) = "=" Or Left(Arr(I, J), 1) = "'" Or IsNumeric(Arr(I, J)) Or IsDate(Arr(I, J)) Then
                                Arr2(I, J) = "'" & Arr(I, J)
                            Else
                                Arr2(I, J) = Arr(I, J)
                            End If
                        Else
                            Arr2(I, J) = Arr(I, J)
                        End If
                    Next J
                Next I
                HdrRange.Offset(ArrPos).Resize(ArrCount).Value = Arr2
            End If
        End If
    End If
    ArrPos = 0
    ArrCount = 0
    Arr = Empty
    Dirty = False
End Sub

Private Sub RecLoad()
'Load data into Buf, clearing beyond data if necessary.
'Sets ArrCount to the number of records in Buf
'LastPos = position of DataRowCount in Buf.

    Dim LastPos As Long
    Dim I As Long, J As Long
    
    ArrPos = ((RecordPos - 1) \ ArrSize) * ArrSize + 1
    LastPos = DataRowCount - ArrPos + 1
    If LastPos > 0 Then
        Arr = HdrRange.Offset(ArrPos).Resize(ArrSize).Value
        For I = LastPos + 1 To ArrSize
            For J = 1 To HdrRange.Columns.Count
                Arr(I, J) = Empty
            Next J
        Next I
    Else
        ReDim Arr(1 To ArrSize, 1 To HdrRange.Columns.Count)
    End If
    ArrCount = DataRowCount - ArrPos + 1
    If ArrCount > ArrSize Then ArrCount = ArrSize
End Sub

'****** Miscellaneous
Public Function GetUnique(Optional Field, Optional FirstRow As Long = 1, Optional LastRow As Long = 0, _
    Optional Criteria As Variant, Optional SortOutput As Boolean = False) As Variant
'Gets the unique values in a column (as a zero-based array of strings).
'Field can be a single field or an array of fields (name or number).
'Result array second dimension will have the same number of entries as Field
'(or may be a 1-demensional array if one field).
'Criteria will be an array of condition elements which join to form a condition formula,
'where data fields are represented by their field name,
'Or a formula with field names delimited by ` characters
'(formula will be evaluated as an array formula)
'may also be a criteria range with a single row of values.
'If none, returns an array with an upper bound of -1.
'If Field is missing, returns row numbers of rows which match Criteria (do not sort).

    Dim Vals As New Dictionary
    Dim Field2 As Variant
    Dim Criteria2 As Variant
    Dim PartRange As Range
    Dim JoinFormula As String
    Dim Addr As String
    Dim Values As Variant
    Dim Keys As Variant
    Dim SplitKey As Variant
    Dim Results As Variant
    Dim I As Long, J As Long
    
    If HasError Then
        GetUnique = Xs.ExtractArray
        Exit Function
    End If
    Vals.CompareMode = TextCompare
    If LastRow = 0 Then LastRow = DataRowCount
    If LastRow < FirstRow Then
        GetUnique = Xs.ExtractArray
        Exit Function
    End If
    Set PartRange = HdrRange.Offset(FirstRow).Resize(LastRow - FirstRow + 1)
    
    If IsArray(Field) Then
        Field2 = Field
    Else
        Field2 = Array(Field)
    End If
    If TypeName(Criteria) = "Range" Then
        Criteria2 = Xs.ExtractArray
        For I = 1 To Criteria.Columns.Count
            If Not IsEmpty(Criteria(2, I)) Then
                ReDim Preserve Criteria2(UBound(Criteria2) + 3)
                If UBound(Criteria2) = 2 Then
                    Criteria2(UBound(Criteria2) - 2) = "("
                Else
                    Criteria2(UBound(Criteria2) - 2) = "*("
                End If
                Criteria2(UBound(Criteria2) - 1) = Criteria(1, I)
                If VarType(Criteria(2, I).Value2) = vbDouble Then
                    Criteria2(UBound(Criteria2)) = "=" & Criteria(2, I).Value2
                Else
                    Criteria2(UBound(Criteria2)) = "=""" & CStr(Criteria(2, I)) & """"
                End If
                Criteria2(UBound(Criteria2)) = Criteria2(UBound(Criteria2)) & ")"
            End If
        Next I
    ElseIf IsArray(Criteria) Then
        Criteria2 = Criteria
    ElseIf IsMissing(Criteria) Then
        Criteria2 = Array(Criteria)
    ElseIf Left(Criteria, 1) = "=" Then
        Criteria2 = Split(Mid(Criteria, 2), "`")
    Else
        Criteria2 = Array(Criteria)
    End If
    
    If LBound(Criteria2) > UBound(Criteria2) Then
        JoinFormula = "IF(TRUE"
    ElseIf IsMissing(Criteria2(LBound(Criteria2))) Then
        JoinFormula = "IF(TRUE"
    Else
        JoinFormula = "IF("
        For I = LBound(Criteria2) To UBound(Criteria2)
            If IsNumeric(Criteria2(I)) Then
                JoinFormula = JoinFormula & CStr(Criteria2(I))
            ElseIf IsDate(Criteria2(I)) Then
                JoinFormula = JoinFormula & CStr(Criteria2(I))
            ElseIf Left(CStr(Criteria2(I)), 1) = """" And Right(CStr(Criteria2(I)), 1) = """" Then
                JoinFormula = JoinFormula & CStr(Criteria2(I))
            ElseIf FieldList.Exists(Criteria2(I)) Then
                Addr = PartRange.Columns(FieldCol(Criteria2(I))).Address(True, True)
                JoinFormula = JoinFormula & Addr
            Else
                JoinFormula = JoinFormula & CStr(Criteria2(I))
            End If
        Next I
    End If
    JoinFormula = JoinFormula & ","
    
    If IsMissing(Field2(LBound(Field2))) Then
        Addr = PartRange.Columns(1).Address(True, True)
        JoinFormula = JoinFormula & "ROW(" & Addr & ")-" & HdrRange.Row
    Else
        For I = LBound(Field2) To UBound(Field2)
            If FieldCheck(Field2(I)) Then
                If I > LBound(Field2) Then JoinFormula = JoinFormula & "&""" & ChrW(SepChar) & """&"
                Addr = PartRange.Columns(FieldCol(Field2(I))).Address(True, True)
                JoinFormula = JoinFormula & Addr
            Else
                HasError = True
                GetUnique = Xs.ExtractArray
                Exit Function
            End If
        Next I
    End If
    JoinFormula = JoinFormula & ",""" & ChrW(SepChar2) & """)"
    
    If LastRow - FirstRow > 0 Then
        Values = HdrRange.Worksheet.Evaluate(JoinFormula)
    Else
        ReDim Values(1 To 1, 1 To 1)
        Values(1, 1) = HdrRange.Worksheet.Evaluate(JoinFormula)
    End If
    
    If IsError(Values) Then
        MsgBox "Formula evaluates to an error attempting to obtain unique values"
        Debug.Print JoinFormula
        HasError = True
        GetUnique = Xs.ExtractArray
        Exit Function
    ElseIf Not IsArray(Values) Then
        MsgBox "Formula does not evaluate to an array attempting to obtain unique values"
        Debug.Print JoinFormula
        HasError = True
        GetUnique = Xs.ExtractArray
        Exit Function
    End If
    
    'Now includes blanks.
    For I = 1 To UBound(Values, 1)
        If IsError(Values(I, 1)) Then
            MsgBox "Error value encountered attempting to obtain unique values"
            HasError = True
            GetUnique = Xs.ExtractArray
            Exit Function
        ElseIf CStr(Values(I, 1)) <> ChrW(SepChar2) Then
            If Not Vals.Exists(CStr(Values(I, 1))) Then Vals.Add CStr(Values(I, 1)), 0
        End If
    Next I
    Keys = Vals.Keys
    If UBound(Keys) < 0 Then
        GetUnique = Xs.ExtractArray
        Exit Function
    End If
    If SortOutput Then
        If UBound(Field2) > LBound(Field2) Then
            QuickSortSep Keys, ChrW(SepChar)
        Else
            QuickSortArray Keys
        End If
    End If
    
    If UBound(Field2) = LBound(Field2) Then
        If IsMissing(Field2(LBound(Field2))) Then
            ReDim Results(UBound(Keys)) As Long
            For I = 0 To UBound(Keys)
                Results(I) = CLng(Keys(I))
            Next I
        Else
            ReDim Results(UBound(Keys)) As String
            For I = 0 To UBound(Keys)
                Results(I) = Keys(I)
            Next I
        End If
    Else
        ReDim Results(UBound(Keys), UBound(Field2) - LBound(Field2)) As String
        For I = 0 To UBound(Keys)
            SplitKey = Split(Keys(I), ChrW(SepChar))
            For J = LBound(Field2) To UBound(Field2)
                Results(I, J - LBound(Field2)) = SplitKey(J - LBound(Field2))
            Next J
        Next I
    End If
    GetUnique = Results
End Function

Public Function Find(Field, Value As Variant) As Long
'Gets row of found cell (first is 1) or returns 0 if not found.

    Dim FoundCell As Range
    
    If HasError Then Exit Function
    If FieldCheck(Field) And DataRowCount > 0 Then
        Set FoundCell = DataRange.Resize(, 1).Offset(, FieldCol(Field) - 1).Find(What:=Value, _
             After:=DataRange.Cells(1, 1).Offset(DataRowCount - 1, FieldCol(Field) - 1), _
            LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, _
            SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False)
        If Not FoundCell Is Nothing Then Find = FoundCell.Row - DataRange.Row + 1
    End If
End Function

Public Function ReplaceAll(Field, Value1 As Variant, Value2 As Variant) As Boolean
'Replaces all occurrences of Value1 with Value2 in the field.
'Returns True if successful

    
    If HasError Then Exit Function
    If FieldCheck(Field) And DataRowCount > 0 Then
        DataRange.Resize(, 1).Offset(, FieldCol(Field) - 1).Replace What:=Value1, _
            Replacement:=Value2, LookAt:=xlWhole, SearchOrder:=xlByColumns, _
            MatchCase:=False, MatchByte:=False
    End If
    ReplaceAll = True
End Function

Public Function CountGroup(Field, Row As Long) As Long
'Counts the adjacent cells in the column with the same value as the cell represented by Row.
'Retuns 0 if no data or Row is 0.
    
    Dim ThisCol As Integer
    Dim Value As Variant
    Dim I As Long

    If HasError Then Exit Function
    If FieldCheck(Field) And DataRowCount > 0 And Row > 0 Then
        ThisCol = FieldCol(Field)
        Value = DataRange.Cells(Row, ThisCol)
        Do
            CountGroup = CountGroup + 1
            If Row + CountGroup > DataRowCount Then Exit Do
        Loop While TestEqual(Value, DataRange.Cells(Row + CountGroup, ThisCol))
    End If
End Function

Public Function Sort(Optional Key1, Optional Order1 As XlSortOrder = xlAscending, _
    Optional Key2, Optional Order2 As XlSortOrder = xlAscending, _
    Optional Key3, Optional Order3 As XlSortOrder = xlAscending) As Boolean
'Sort on up to 3 keys.
'Returns True if successful.
'Removes any index.
'Adjusted to avoid crashes on the Mac.

    Dim WbName As String
    Dim RangeAdr As String
    Dim Key1Adr As String
    Dim Key2Adr As String
    Dim Key3Adr As String
    
    If HasError Then Exit Function
    If DataRowCount = 0 Then
        Sort = True
        Exit Function
    End If
    RemoveIndex
    
    With HdrRange.Resize(DataRowCount + 1)
        WbName = .Parent.Parent.Name
        RangeAdr = .Address(True, True)
        If IsMissing(Key3) Or Len(CStr(Key3)) = 0 Then
            If IsMissing(Key2) Or Len(CStr(Key2)) = 0 Then
                If IsMissing(Key1) Or Len(CStr(Key1)) = 0 Then
                    Sort = True
                Else
                    If FieldCheck(Key1) Then
                        Key1Adr = .Resize(, 1).Offset(, FieldCol(Key1) - 1).Address(True, True)
                        Workbooks(WbName).Worksheets(SheetName).Range(RangeAdr).Sort _
                            Key1:=Workbooks(WbName).Worksheets(SheetName).Range(Key1Adr), Order1:=Order1, Header:=xlYes
                        Sort = True
                    End If
                End If
            ElseIf FieldCheck(Key1) And FieldCheck(Key2) Then
                Key1Adr = .Resize(, 1).Offset(, FieldCol(Key1) - 1).Address(True, True)
                Key2Adr = .Resize(, 1).Offset(, FieldCol(Key2) - 1).Address(True, True)
                Workbooks(WbName).Worksheets(SheetName).Range(RangeAdr).Sort _
                    Key1:=Workbooks(WbName).Worksheets(SheetName).Range(Key1Adr), Order1:=Order1, _
                    Key2:=Workbooks(WbName).Worksheets(SheetName).Range(Key2Adr), Order2:=Order2, Header:=xlYes
                Sort = True
            End If
        ElseIf FieldCheck(Key1) And FieldCheck(Key2) And FieldCheck(Key3) Then
            Key1Adr = .Resize(, 1).Offset(, FieldCol(Key1) - 1).Address(True, True)
            Key2Adr = .Resize(, 1).Offset(, FieldCol(Key2) - 1).Address(True, True)
            Key3Adr = .Resize(, 1).Offset(, FieldCol(Key3) - 1).Address(True, True)
            Workbooks(WbName).Worksheets(SheetName).Range(RangeAdr).Sort _
                Key1:=Workbooks(WbName).Worksheets(SheetName).Range(Key1Adr), Order1:=Order1, _
                Key2:=Workbooks(WbName).Worksheets(SheetName).Range(Key2Adr), Order2:=Order2, _
                Key3:=Workbooks(WbName).Worksheets(SheetName).Range(Key3Adr), Order3:=Order3, Header:=xlYes
            Sort = True
        End If
    End With
End Function

Public Function SetBorders(LineStyle As XlLineStyle, BorderWeight As XlBorderWeight, _
    Optional IncludeHeaders As Boolean = False) As Boolean
'Sets all borders.

    If HasError Then Exit Function
    If DataRowCount = 0 Then
        SetBorders = True
        Exit Function
    End If
    With IIf(IncludeHeaders, Union(HdrRange, DataRange), DataRange)
        With .Borders(xlEdgeTop)
            .LineStyle = LineStyle
            .Weight = BorderWeight
        End With
        With .Borders(xlEdgeLeft)
            .LineStyle = LineStyle
            .Weight = BorderWeight
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = LineStyle
            .Weight = BorderWeight
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = LineStyle
            .Weight = BorderWeight
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = LineStyle
            .Weight = BorderWeight
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = LineStyle
            .Weight = BorderWeight
        End With
    End With
    SetBorders = True
End Function

Public Function GetRange() As Range
'Returns the complete range including headers and data.

    If HasError Then Exit Function
    If DataRowCount = 0 Then
        Set GetRange = HdrRange
    Else
        Set GetRange = Union(HdrRange, DataRange)
    End If
End Function

Public Function Compare(Row1 As Long, Row2 As Long, ParamArray Headers()) As Boolean
'Compares two rows for the given fields.

    Dim Headers2 As Variant
    Dim Headersx As Variant
    Dim Xs As New XShared
    Dim I As Long, J As Long

    If HasError Then Exit Function
    Headersx = Headers
    Headers2 = Xs.ExtractArray(Headersx)
    If UBound(Headers2) < LBound(Headers2) Then
        Xs.MatchFields Me, Nothing, Xs.ExtractArray, Headers2
    End If
    For I = LBound(Headers2) To UBound(Headers2)
        If Not FieldCheck(Headers2(I)) Then Exit Function
        J = FieldCol(Headers2(I))
        If Not TestEqual(DataRange(Row1, J), DataRange(Row2, J)) Then Exit Function
    Next I
    Compare = True
End Function

Public Function CheckBlank(Row As Long) As Boolean
'Checks if all the cells in the specified row are empty, empty strings, | or 0.
    
    Dim Val As Variant
    Dim I As Long
    
    If HasError Then Exit Function
    For I = 1 To DataRange.Columns.Count
        Val = DataRange.Cells(Row, I)
        Select Case VarType(Val)
        Case vbEmpty
        Case vbString
            Select Case Val
            Case vbNullString, "|"
            Case Else
                Exit Function
            End Select
        Case vbDouble
            If Val <> 0 Then Exit Function
        Case Else
            Exit Function
        End Select
    Next I
    CheckBlank = True
End Function

Public Function AdvancedFilter(Source As Range, Optional Criteria As Range = Nothing, _
    Optional Unique As Boolean = False, Optional Columns As Integer = 0) As Boolean
'Advanced filter with this as the destination.
'Targets number of columns if specified.

    Dim J As Long, K As Long, M As Long
    
    If HasError Then Exit Function
    If Columns = 0 Then Columns = HdrRange.Columns.Count
    If Not ListObj Is Nothing Then ListObj.ShowTotals = False
    Clear True, , True, True
    AddRow
    If Criteria Is Nothing Then
        Source.AdvancedFilter xlFilterCopy, , HdrRange.Resize(, Columns), Unique
    Else
        Source.AdvancedFilter xlFilterCopy, Criteria, HdrRange.Resize(, Columns), Unique
    End If
    
    With HdrRange
        M = .Row
        For J = .Column To .Column + Columns - 1
            K = .Worksheet.Cells(.Worksheet.Rows.Count, J).End(xlUp).Row
            If M < K Then M = K
        Next J
        If M < .Row Then M = .Row
        DataRowCount = M - .Row
        If DataRowCount > 0 Then
            Set DataRange = HdrRange.Offset(1, 0).Resize(DataRowCount)
        Else
            Set DataRange = Nothing
        End If
    End With
    
    If Not ListObj Is Nothing Then
        Select Case DataRowCount
        Case 0
            ListObj.ListRows(1).Delete
        Case 1
            If CheckBlank(1) Then
                ListObj.ListRows(1).Delete
                DataRowCount = 0
                Set DataRange = Nothing
            End If
        Case Else
            ListObj.Resize ListObj.HeaderRowRange.Resize(DataRowCount + 1)
        End Select
        If ShowTotals Then ListObj.ShowTotals = True
    End If
    
    ClearBuf
    RemoveIndex
    Arr = Empty
    ArrPos = 0
    RecordPos = 0
    AdvancedFilter = True
End Function

Public Function RemoveSubtotal() As Boolean
'Remove subtotals.

    If HasError Then Exit Function
    If DataRowCount = 0 Then
        RemoveSubtotal = True
        Exit Function
    End If
    If Not ListObj Is Nothing Then
        MsgBox "Cannot subtotal a table"
        HasError = True
        Exit Function
    End If
    RemoveIndex
    
    Union(HdrRange, DataRange).RemoveSubtotal
    If Not Initialise(HdrRange, , , True) Then Exit Function
    RemoveSubtotal = True
End Function

Public Function Subtotal(ByVal GroupBy As Variant, SubFunction As XlConsolidationFunction, Replace As Boolean, _
    PageBreaks As Boolean, SummaryBelowData As XlSummaryRow, ParamArray TotalList()) As Boolean
'Create subtotal.

    Dim TotalList2 As Variant
    Dim TotalListx As Variant
    Dim I As Long

    If HasError Then Exit Function
    If DataRowCount = 0 Then
        Subtotal = True
        Exit Function
    End If
    If Not ListObj Is Nothing Then
        MsgBox "Cannot subtotal a table"
        HasError = True
        Exit Function
    End If
    RemoveIndex
    
    If Not FieldCheck(GroupBy) Then Exit Function
    GroupBy = FieldCol(GroupBy)
    
    TotalListx = TotalList
    TotalList2 = Xs.ExtractArray(TotalListx)
    For I = 0 To UBound(TotalList2)
        If Not FieldCheck(TotalList2(I)) Then Exit Function
        TotalList2(I) = FieldCol(TotalList2(I))
    Next I
    
    Union(HdrRange, DataRange).Subtotal GroupBy, SubFunction, TotalList2, Replace, _
        PageBreaks, SummaryBelowData
    If Not Initialise(HdrRange, , , True) Then Exit Function
    
    Subtotal = True
End Function

Public Function PageBreak(Field As Variant) As Boolean
'Inserts page breaks on change of Field.

    Dim ColRange As Range
    Dim ColVals As Variant
    Dim PArea As String
    Dim I As Long

    If HasError Then Exit Function
    HdrRange.Worksheet.DisplayPageBreaks = False
    HdrRange.Worksheet.ResetAllPageBreaks
    If DataRowCount <= 1 Then
        PageBreak = True
        Exit Function
    End If
    Set ColRange = GetColumn(Field)
    If ColRange Is Nothing Then Exit Function
    
    HdrRange.Worksheet.DisplayPageBreaks = False
    PArea = HdrRange.Worksheet.PageSetup.PrintArea
    HdrRange.Worksheet.PageSetup.PrintArea = vbNullString
    ColVals = ColRange.Value
    For I = 2 To DataRowCount
        If ColVals(I, 1) <> ColVals(I - 1, 1) Then
            GetRow(I).EntireRow.PageBreak = xlPageBreakManual
        End If
    Next I
    HdrRange.Worksheet.PageSetup.PrintArea = PArea
    PageBreak = True
End Function

Public Function ToHeader(Field As Variant) As Boolean
'Converts a column to headers on change of value.

    Dim ThisCol As Range
    Dim ColVals As Variant
    Dim I As Long
    
    If HasError Then Exit Function
    Set ThisCol = GetColumn(Field)
    ColVals = ThisCol.Value
    ThisCol.ClearContents
    
    For I = DataRowCount - 1 To 1 Step -1
        If ColVals(I, 1) <> ColVals(I + 1, 1) Then
            InsertRows 1, I + 1
            Value(Field, I + 1) = ColVals(I + 1, 1)
        End If
        If I = 1 Then
            InsertRows 1, 1
            Value(Field, 1) = ColVals(1, 1)
        End If
        If HasError Then Exit Function
    Next I
    ToHeader = True
End Function

Public Function TextFields(ParamArray Field()) As Boolean
'Call before Export to specify list of fields where text values should be protected on export.
'Call with no parameters to protect all text fields.

    Dim Field2 As Variant
    Dim Fieldx As Variant
    Dim I As Long

    If HasError Then Exit Function
    Fieldx = Field
    Field2 = Xs.ExtractArray(Fieldx)
    If UBound(Field2) < 0 Then
        For I = 1 To UBound(Fld)
            Fld(I).IsText = True
        Next I
    Else
        For I = LBound(Field2) To UBound(Field2)
            If Not FieldCheck(Field2(I)) Then Exit Function
            Fld(FieldCol(Field2(I))).IsText = True
        Next I
    End If
End Function

Public Function Export(Name As String, ParamArray Field()) As Boolean
'Export all or named fields to external workbook.

    Dim Ws As Worksheet
    Dim Wb2 As Workbook
    Dim Field2 As Variant
    Dim Fieldx As Variant
    Dim I As Long
    
    If HasError Then Exit Function
    Set Ws = Xs.CreateSource(Me, Name)
    If Ws Is Nothing Then Exit Function
    DoEvents
    Fieldx = Field
    Field2 = Xs.ExtractArray(Fieldx)
    
    If UBound(Field2) < 0 Then
        If DataRowCount > 0 Then
            For I = 1 To UBound(Fld)
                If Fld(I).IsText Then
                    Ws.Cells(2, I).Resize(DataRowCount).NumberFormat = "@"
                End If
            Next I
        End If
        Ws.Cells(1, 1).Resize(DataRowCount + 1, UBound(Fld)).Value = _
            HdrRange.Resize(DataRowCount + 1).Value
    Else
        For I = LBound(Field2) To UBound(Field2)
            If Not FieldCheck(Field2(I)) Then Exit Function
            If DataRowCount > 0 Then
                If Fld(FieldCol(Field2(I))).IsText Then
                    Ws.Cells(2, I - LBound(Field2) + 1).Resize(DataRowCount).NumberFormat = "@"
                End If
            End If
            Ws.Cells(1, I - LBound(Field2) + 1).Resize(DataRowCount + 1).Value = _
                HdrRange.Offset(, FieldCol(Field2(I)) - 1).Resize(DataRowCount + 1).Value
        Next I
    End If
    If DataRowCount > 0 Then Ws.Cells(2, 1).Resize(DataRowCount).EntireRow.RowHeight = Ws.Rows(1).RowHeight
    Export = True
End Function

Public Function SaveXML(Filename As Variant) As Boolean
'Saves as a XML file in PersistXML format.

    Dim xlxml As Object
    
    If HasError Then Exit Function
    
    If IsObject(Filename) Then
        Filename.LoadXML GetRange.Value(xlRangeValueMSPersistXML)
        SaveXML = True
        Exit Function
    End If
    
    Set xlxml = CreateObject("MSXML2.DOMDocument.6.0")
    xlxml.LoadXML GetRange.Value(xlRangeValueMSPersistXML)
    
    On Error Resume Next
    Kill Filename
    Err.Clear
    xlxml.Save Filename
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox Err.Description
        On Error GoTo 0
        HasError = True
        Exit Function
    End If
    SaveXML = True
End Function

Public Function OpenXML(Filename As Variant, Optional ByVal TableStart As Range, _
    Optional VariableWidth As Boolean = False) As Boolean
'Opens a XML file in PersistXML format.
'Recordset is opened forward only.

    Dim Rd As Object

    If HasError Then Exit Function
    Set Rd = Xs.GetRd
    Rd.OpenXML Filename
    If Rd.HasError Then Exit Function
    CreateFrom Rd, TableStart, VariableWidth
    If HasError Then Exit Function
    Set Rd = Xs.GetRd(Rd)
    Rd.RsetClose
    OpenXML = True
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

Private Function TestEqual(Value1, Value2) As Boolean
    'Returns TRUE if Value1 and Value2 have the same data type and value.
    
    Select Case VarType(Value1)
    Case vbNull
        If IsNull(Value2) Then
            TestEqual = True
            Exit Function
        End If
    Case VarType(Value2)
        If Value1 = Value2 Then
            TestEqual = True
            Exit Function
        End If
    End Select
    TestEqual = False
End Function

Private Function QuickSortArray(varArray As Variant, _
                        Optional lngFirst As Long = -1, _
                        Optional lngLast As Long = -1) As Variant
                            
' QuickSort algorithm used to sort array.
    
    Dim lngLow      As Long
    Dim lngHigh     As Long
    Dim lngMiddle   As Long
    Dim varTempVal  As Variant
    Dim varTestVal  As Variant
    
    If lngFirst = -1 Then lngFirst = LBound(varArray)
    If lngLast = -1 Then lngLast = UBound(varArray)
        
    If lngFirst < lngLast Then
        lngMiddle = (lngFirst + lngLast) / 2
        varTestVal = varArray(lngMiddle)
        lngLow = lngFirst
        lngHigh = lngLast
        Do
            Do While varArray(lngLow) < varTestVal
                lngLow = lngLow + 1
            Loop
            Do While varArray(lngHigh) > varTestVal
                lngHigh = lngHigh - 1
            Loop
            If (lngLow <= lngHigh) Then
                varTempVal = varArray(lngLow)
                varArray(lngLow) = varArray(lngHigh)
                varArray(lngHigh) = varTempVal
                lngLow = lngLow + 1
                lngHigh = lngHigh - 1
            End If
        Loop While (lngLow <= lngHigh)
        If lngFirst < lngHigh Then QuickSortArray varArray, lngFirst, lngHigh
        If lngLow < lngLast Then QuickSortArray varArray, lngLow, lngLast
    End If
End Function

Private Function QuickSortSep(varArray As Variant, Sep As String, _
                        Optional lngFirst As Long = -1, _
                        Optional lngLast As Long = -1) As Variant
                            
    ' QuickSort algorithm used to sort array.
    
    Dim lngLow      As Long
    Dim lngHigh     As Long
    Dim lngMiddle   As Long
    Dim varTempVal  As Variant
    Dim varTestVal  As Variant
    
    If lngFirst = -1 Then lngFirst = LBound(varArray)
    If lngLast = -1 Then lngLast = UBound(varArray)
        
    If lngFirst < lngLast Then
        lngMiddle = (lngFirst + lngLast) / 2
        varTestVal = varArray(lngMiddle)
        lngLow = lngFirst
        lngHigh = lngLast
        Do
            Do While CompStrings(varArray(lngLow), varTestVal, Sep) < 0
                lngLow = lngLow + 1
            Loop
            Do While CompStrings(varArray(lngHigh), varTestVal, Sep) > 0
                lngHigh = lngHigh - 1
            Loop
            If (lngLow <= lngHigh) Then
                varTempVal = varArray(lngLow)
                varArray(lngLow) = varArray(lngHigh)
                varArray(lngHigh) = varTempVal
                lngLow = lngLow + 1
                lngHigh = lngHigh - 1
            End If
        Loop While (lngLow <= lngHigh)
        If lngFirst < lngHigh Then QuickSortSep varArray, Sep, lngFirst, lngHigh
        If lngLow < lngLast Then QuickSortSep varArray, Sep, lngLow, lngLast
    End If
End Function

Private Function CompStrings(Str1, Str2, Sep As String) As Integer
'Compares strings with separator
'Returns -1 if Str1<str2, 0 if Str1=Str2, 1 if Str1>Str2

    Dim Split1 As Variant
    Dim Split2 As Variant
    Dim Cnt As Long
    Dim I As Long
    
    If Len(Str1) = 0 Then
        If Len(Str2) = 0 Then
            CompStrings = 0
            Exit Function
        End If
        CompStrings = -1
        Exit Function
    ElseIf Len(Str2) = 0 Then
        CompStrings = 1
        Exit Function
    Else
        Split1 = Split(Str1, Sep)
        Split2 = Split(Str2, Sep)
        Cnt = UBound(Split1)
        If Cnt > UBound(Split2) Then Cnt = UBound(Split2)
        For I = 0 To Cnt
            Select Case Split1(I)
            Case Is < Split2(I)
                CompStrings = -1
                Exit Function
            Case Is > Split2(I)
                CompStrings = 1
                Exit Function
            End Select
        Next I
    End If
    Select Case UBound(Split1)
    Case Is < UBound(Split2)
        CompStrings = -1
    Case Is > UBound(Split2)
        CompStrings = 1
    End Select
End Function
