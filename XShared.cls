VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XShared"
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

#Const ArrMode = True 'True if ArrData is included
#Const RsetMode = True 'True if RsetData is included
#Const TableMode = True 'True if TableData is included
#Const RegexMode = True 'True to use Microsoft VBScript Regular Expressions 5.5, otherwise uses Jscript.
#Const UserFunctionMode = False 'True if user functions are defined in XUserFunctions class module
Private Const DefShowMode As Boolean = False 'Show in IE window

Dim LastShow As Date
Public NoShow As Boolean
Public ShowMode As Boolean
Const RepChar1 As String = "\1" 'replaces \
Const RepChar2 As String = "\2" 'replaces ,
Const RepChar3 As String = "\3" 'replaces ;
Const RepChar4 As String = "\4" 'replaces "

#If UserFunctionMode Then
Dim Uf As New XUserFunctions
#End If

Private Sub Class_Initialize()
    ShowMode = DefShowMode
End Sub

#If ArrMode Then

'-----------------------------------------------------
'Calculation Functions
'-----------------------------------------------------
Public Function RecordNumber(Ad As ArrData, Params As Variant) As Variant

    RecordNumber = Ad.RecPos
End Function

Public Static Function Expand(Ad As ArrData, Params As Variant) As Variant
'Creates an array of all the integer values between the two parameters.

    Dim Initialised As Boolean
    Dim Field1 As Long
    Dim Field2 As Long
    Dim Field As Variant
    Dim Text As String
    Dim I As Long
    
    If Not Initialised Then
        If UBound(Params) <> 1 Then
            Msg Ad, "Wrong number of parameters in Expand"
            Exit Function
        End If
        If Not FieldCheck(Ad, Params(0)) Then Exit Function
        If Not FieldCheck(Ad, Params(1)) Then Exit Function
        Field1 = Ad.Column(Params(0))
        Field2 = Ad.Column(Params(1))
        For Each Field In Params
            Select Case Ad.GetType(Field)
            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbByte
            Case Else
                Msg Ad, "Field " & Ad.GetHeader(Field) & " not numeric in Expand"
                Exit Function
            End Select
        Next Field
        Initialised = True
    End If
    
    If CLng(Ad.RecValue(Field2)) < CLng(Ad.RecValue(Field1)) Then
        Msg Ad, "First value is greater than second in Expand"
        Exit Function
    End If
    
    Text = vbNullString
    For I = CLng(Ad.RecValue(Field1)) To CLng(Ad.RecValue(Field2))
        If Len(Text) > 0 Then Text = Text & ","
        Text = Text & CStr(I)
    Next I
    Expand = "{" & Text & "}"
End Function

Public Static Function Combine(Ad As ArrData, Params As Variant) As Variant
'Creates an array combining the parameters (which may be arrays).
'Parameters must be either field names/numbers or literals
'Literals may be preceded by $ to ensure they are not interpreted as field names.
'Previously ignored empty entries, this now removed.

    Dim Initialised As Boolean
    Dim Fields As Variant
    Dim Texts As Variant
    Dim Text As String
    Dim I As Long
    
    If Not Initialised Then
        ReDim Fields(LBound(Params) To UBound(Params)) As Long
        ReDim Texts(LBound(Params) To UBound(Params)) As String
        For I = LBound(Params) To UBound(Params)
            If Left(CStr(Params(I)), 1) = "$" Then
                Texts(I) = Right(CStr(Params(I)), Len(CStr(Params(I))) - 1)
                If Left(Texts(I), 1) = "{" And Right(Texts(I), 1) = "}" Then Texts(I) = Mid(Texts(I), 2, Len(Texts(I)) - 2)
            ElseIf Ad.Exists(Params(I)) Then
                Fields(I) = Ad.Column(Params(I))
            Else
                Texts(I) = CStr(Params(I))
                If Left(Texts(I), 1) = "{" And Right(Texts(I), 1) = "}" Then Texts(I) = Mid(Texts(I), 2, Len(Texts(I)) - 2)
            End If
        Next I
        Initialised = True
    End If
    
    Text = vbNullString
    For I = LBound(Params) To UBound(Params)
        If Fields(I) > 0 Then
            Texts(I) = CStr(Ad.RecValue(Fields(I)))
            If Left(Texts(I), 1) = "{" And Right(Texts(I), 1) = "}" Then Texts(I) = Mid(Texts(I), 2, Len(Texts(I)) - 2)
        End If
        If I > LBound(Params) Then Text = Text & ","
        Text = Text & Texts(I)
    Next I
    Combine = "{" & Text & "}"
End Function

Public Static Function ToTextArray(Ad As ArrData, Params As Variant)
'Convert the parameter field to a single element text array.

    Dim Initialised As Boolean
    Dim Field As Long
    Dim Text As String

    If Not Initialised Then
        If Not FieldCheck(Ad, Params(0)) Then Exit Function
        Field = Ad.Column(Params(0))
        Initialised = True
    End If
    
    Text = CStr(Ad.RecValue(Field))
    Text = Replace(Text, "\", RepChar1)
    Text = Replace(Text, ",", RepChar2)
    Text = Replace(Text, ";", RepChar3)
    Text = Replace(Text, """", RepChar4)
    Text = "{""" & Text & """}"
    ToTextArray = Text
End Function

Public Static Function Total(Ad As ArrData, Params As Variant) As Variant
'Returns the total.
'Parameters are Field and GroupFunction.

    Dim Initialised As Boolean
    Dim TotValue As Variant
    Dim Field As Variant
    Dim GroupFunction As String

    If Not Initialised Then
        If UBound(Params) <> 1 Then
            Msg Ad, "Wrong number of parameters in Total"
            Exit Function
        End If
        If Not FieldCheck(Ad, Params(0)) Then Exit Function
        Field = Params(0)
        GroupFunction = Params(1)
        TotValue = Ad.Total(Field, GroupFunction)
        If Ad.HasError Then Exit Function
        Initialised = True
    End If
    Total = TotValue
End Function

Public Static Function ConvertFrac(Ad As ArrData, Params As Variant) As Variant
'First time checks parameters.
'Sets error if a problem.

    Dim Initialised As Boolean
    Dim FieldNo As Long
    
    If Not Initialised Then
        If UBound(Params) <> 0 Then
            Msg Ad, "Wrong number of parameters in ConvertFrac"
            Exit Function
        End If
        If Not FieldCheck(Ad, CStr(Params(0))) Then Exit Function
        FieldNo = Ad.Column(CStr(Params(0)))
        Initialised = True
    End If
    
    On Error Resume Next
    ConvertFrac = Frac2Num(CStr(Ad.RecValue(FieldNo)))
    If Err = 0 Then
        On Error GoTo 0
    Else
        Msg Ad, Err.Description & " error calculating ConvertFrac"
        On Error GoTo 0
    End If
End Function

#If RegexMode Then
Public Static Function RegExpr(Ad As ArrData, Params As Variant) As String
'Calculates regular expression using Microsoft VBscript Regular Expressions 5.5.
'Parameters: Field, Pattern, IgnoreCase, Global, Match number
'Only looks at first match.

    Dim Regex As Object
    Dim Matches As Object
    Dim Initialised As Boolean
    Dim FieldNo As Long
    
    If Not Initialised Then
        If UBound(Params) <> 4 Then
            Msg Ad, "Wrong number of parameters in RegExpr"
            Exit Function
        End If
        If Not FieldCheck(Ad, CStr(Params(0))) Then Exit Function
        FieldNo = Ad.Column(CStr(Params(0)))
        On Error Resume Next
        Set Regex = CreateObject("VBScript.RegExp")
        Regex.Pattern = CStr(Params(1))
        Regex.IgnoreCase = CBool(Params(2))
        Regex.Global = CBool(Params(3))
        If Err = 0 Then
            On Error GoTo 0
        Else
            On Error GoTo 0
            Msg Ad, "Error initialising regular expression " & CStr(Params(1))
            Exit Function
        End If
        Initialised = True
    End If
    
    On Error Resume Next
    Set Matches = Regex.Execute(CStr(Ad.RecValue(FieldNo)))
        If Err = 0 Then
            On Error GoTo 0
        Else
            On Error GoTo 0
            Msg Ad, "Error matching regular expression " & CStr(Params(1))
            Exit Function
        End If
        If Matches Is Nothing Then
        Msg Ad, "Regular expression " & CStr(Params(1)) & " not matched for " & CStr(Params(0)) & " " & CStr(Ad.RecValue(FieldNo))
    ElseIf Matches.Count = 0 Then
        Msg Ad, "Regular expression " & CStr(Params(1)) & " not matched for " & CStr(Params(0)) & " " & CStr(Ad.RecValue(FieldNo))
    Else
        RegExpr = Matches(0).SubMatches(CLng(Params(4)))
    End If
End Function

#Else
Public Static Function RegExpr(Ad As ArrData, Params As Variant) As String
'Calculates regular expression using Jscript.
'Parameters: Field, Pattern, IgnoreCase, Global, Match number
'Global parameter is ignored

    Dim FieldNo As Long
    Dim Initialised As Boolean
    Dim IgnoreCase As String
    Dim Matches As Variant

    If Not Initialised Then
        If UBound(Params) <> 4 Then
            Msg Ad, "Wrong number of parameters in RegExpr"
            Exit Function
        End If
        If Not FieldCheck(Ad, CStr(Params(0))) Then Exit Function
        FieldNo = Ad.Column(CStr(Params(0)))
        IgnoreCase = IIf(CBool(Params(2)), "i", vbNullString)
        Initialised = True
    End If
    
    If Not JsRmatch(CStr(Params(1)), IgnoreCase, CStr(Ad.RecValue(FieldNo)), Matches) Then
        Ad.HasError = True
    ElseIf UBound(Matches) < CLng(Params(4) + 1) Then
        Msg Ad, "Regular expression " & CStr(Params(1)) & " not matched for " & CStr(Params(0)) & " " & CStr(Ad.RecValue(FieldNo))
    Else
        RegExpr = Matches(CLng(Params(4)) + 1)
    End If
End Function
#End If

Private Static Function JsRmatch(Pattern As String, Modifiers As String, InputData As String, ReturnVal As Variant) As Boolean
'Matches regular expression using Javascript and returns array of match and submatches
    
    Dim Scontrol As Object
    Dim Mcount As Long
    Dim I As Long
    Const Script As String = "var matches;" & _
        "  function reMatch(thisPattern, thisModifiers, thisData) {" & _
        "  var regex = new RegExp(thisPattern, thisModifiers);" & _
        "  matches = regex.exec(thisData);" & _
        "  if(matches) {" & _
        "    return matches.length;" & _
        "  } else {" & _
        "    return 0;" & _
        " }}"
        
    If Scontrol Is Nothing Then
        Set Scontrol = CreateObject("ScriptControl")
        With Scontrol
            .Language = "JScript"
            .AddCode Script
        End With
    End If

    With Scontrol
        On Error Resume Next
        Mcount = .Run("reMatch", Pattern, Modifiers, InputData)
        If Err = 0 Then
            On Error GoTo 0
        Else
            On Error GoTo 0
            MsgBox "Error matching regular expression " & Pattern
            ReturnVal = Array()
            Exit Function
        End If
        
        If Mcount = 0 Then
            ReturnVal = Array()
        Else
            ReDim ReturnVal(0 To Mcount - 1) As String
            For I = 0 To Mcount - 1
                ReturnVal(I) = Scontrol.Eval("matches[" & I & "]")
            Next I
        End If
    End With
    JsRmatch = True
End Function

'-----------------------------------------------------
'Group functions.
'Ad.GroupData is a 1-based array of values
'-----------------------------------------------------
Function Sum(Ad As ArrData) As Variant
'On text, concatenates.
    
    Dim GroupData As Variant
    Dim Result As Variant
    Dim I As Long
    
    GroupData = Ad.GroupData
    For I = 1 To UBound(GroupData)
        Result = Result + GroupData(I)
    Next I
    Sum = Result
End Function

Function Count(Ad As ArrData) As Long
    
    Count = UBound(Ad.GroupData)
End Function

Function Average(Ad As ArrData) As Variant
    
    Dim GroupData As Variant
    Dim Result As Variant
    Dim I As Long
    
    GroupData = Ad.GroupData
    For I = 1 To UBound(GroupData)
        Result = Result + GroupData(I)
    Next I
    Average = Result / (UBound(GroupData))
End Function

Function Max(Ad As ArrData) As Variant

    Dim GroupData As Variant
    Dim Result As Variant
    Dim I As Long
    
    GroupData = Ad.GroupData
    Result = GroupData(1)
    For I = 2 To UBound(GroupData)
        If Result < GroupData(I) Then Result = GroupData(I)
    Next I
    Max = Result
End Function

Function Min(Ad As ArrData) As Variant

    Dim GroupData As Variant
    Dim Result As Variant
    Dim I As Long
    
    GroupData = Ad.GroupData
    Result = GroupData(1)
    For I = 2 To UBound(GroupData)
        If Result > GroupData(I) Then Result = GroupData(I)
    Next I
    Min = Result
End Function

Function First(Ad As ArrData) As Variant

    First = Ad.GroupData(1)
End Function

Function Last(Ad As ArrData) As Variant

    Last = Ad.GroupData(UBound(Ad.GroupData))
End Function

Function NumArray(Ad As ArrData) As Variant
'Creates a numeric array.

    Dim GroupData As Variant
    Dim I As Long
    
    If VarType(Ad.GroupData) <> vbArray + vbString Then
        Msg Ad, "Can only create an array from a text field"
        Exit Function
    End If
    GroupData = Ad.GroupData
    For I = 1 To UBound(GroupData)
        GroupData(I) = Trim(GroupData(I))
    Next I
    NumArray = "{" & Join(GroupData, ",") & "}"
End Function

Function TextArray(Ad As ArrData) As Variant
'Creates a text array.
'Replaces \ , ; with \1 \2 \3 (reversed on ungroup).
    
    Dim GroupData As Variant
    Dim I As Long

    If VarType(Ad.GroupData) <> vbArray + vbString Then
        Msg Ad, "Can only create an array from a text field"
        Exit Function
    End If
    GroupData = Ad.GroupData
    For I = 1 To UBound(GroupData)
        GroupData(I) = Replace(GroupData(I), "\", RepChar1)
        GroupData(I) = Replace(GroupData(I), ",", RepChar2)
        GroupData(I) = Replace(GroupData(I), ";", RepChar3)
        GroupData(I) = Replace(GroupData(I), """", RepChar4)
        GroupData(I) = """" & GroupData(I) & """"
    Next I
    TextArray = "{" & Join(GroupData, ",") & "}"
End Function

Function UniqueList(Ad As ArrData) As Variant
'Creates a comma-separated unique list (with space after comma).

    Dim GroupData As Variant
    Dim dValues As New Dictionary
    Dim Value As String
    Dim Values As Variant
    Dim I As Long

    If VarType(Ad.GroupData) <> vbArray + vbString Then
        Msg Ad, "Can only create a list from a text field"
        Exit Function
    End If
    
    GroupData = Ad.GroupData
    dValues.CompareMode = TextCompare
    For I = 1 To UBound(GroupData)
        Value = CStr(GroupData(I))
        Value = Trim(Value)
        If Not dValues.Exists(Value) Then
            dValues.Add Value, Value
        End If
    Next I
    Values = dValues.Keys
    UniqueList = Join(Values, ", ")
End Function

Function ArrArray(Ad As ArrData) As Variant
'Creates a 2-dimensional array.

    Dim GroupData() As String
    Dim I As Long
    
    If VarType(Ad.GroupData) <> vbArray + vbString Then
        Msg Ad, "Can only create an array from a text field"
        Exit Function
    End If
    GroupData = Ad.GroupData
    For I = 1 To UBound(GroupData)
        If Left(GroupData(I), 1) = "{" Then GroupData(I) = Mid(GroupData(I), 2)
        If Right(GroupData(I), 1) = "}" Then GroupData(I) = _
            Left(GroupData(I), Len(GroupData(I)) - 1)
    Next I
    If UBound(GroupData) = 1 Then
        ArrArray = "{" & GroupData(1) & " }"
    Else
        ArrArray = "{" & Join(GroupData, ";") & "}"
    End If
End Function

'-----------------------------------------------------
'XML functions
'-----------------------------------------------------

Function SaveAllXML(Filename As String, ParamArray AllArr()) As Boolean

    Dim XMLDoc0 As Object
    Dim XMLDoc1 As Object
    Dim Element0 As Object
    Dim Element1 As Object
    Dim Element2 As Object
    Dim ProcInst As Object
    Dim ArrObjects As Variant
    Dim AllArr2 As Variant
    Dim ArrObj As Variant
    Dim Ad As ArrData
    
    AllArr2 = AllArr
    ArrObjects = ExtractArray(AllArr2)
    Set XMLDoc0 = CreateObject("MSXML2.DOMDocument.6.0")
    Set ProcInst = XMLDoc0.createProcessingInstruction("xml", " version='1.0' encoding='UTF-16'")
    XMLDoc0.appendChild ProcInst
    Set Element0 = XMLDoc0.createElement("AllTables")
    
    For Each ArrObj In ArrObjects
        Set Ad = ArrObj
        If Ad.HasError Then Exit Function
        If Len(Ad.ThisObjName) = 0 Then
            MsgBox "Missing ad name"
            Exit Function
        End If
        Set XMLDoc1 = CreateObject("MSXML2.DOMDocument.6.0")
        Ad.SaveXML XMLDoc1
        If Ad.HasError Then Exit Function
        
        Set Element1 = XMLDoc1.FirstChild
        Set Element2 = XMLDoc0.createElement(Ad.ThisObjName)
        Element2.appendChild Element1
        Element0.appendChild Element2
    Next ArrObj
    
    XMLDoc0.appendChild Element0
    On Error Resume Next
    Kill Filename
    Err.Clear
    XMLDoc0.Save Filename
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox Err.Description
        On Error GoTo 0
        Exit Function
    End If
    SaveAllXML = True
End Function

#End If

'-----------------------------------------------------
'Common Code
'-----------------------------------------------------

Public Function Copy(Xd1 As XData, Xd2 As XData, ByVal Start As Long, _
    ByVal Recs As Long, Fields2 As Variant) As Boolean
'Copy TableData, Rsetdata or ArrData object and append.
'Xd1 is the destination from which this is called.
'Xd2 is the source.
'If Start and Recs are both 0, assumes start at the first record and finises at the last record.
'If Start is <1, assumes starts at the first record.
'If Recs is negative, assumes finishes at the last record.
'Fields2 is an array of fields.
'If no list of fields supplied then all fields are copied if present on Xd2.
'If Fields = * then all fields are copied and must be present on Xd2.
    
    Dim Fields3 As Variant 'Fields columns
    Dim Fields4 As Variant 'Fields columns on Xd2
    Dim I As Long, J As Long
    
    If Xd2.HasError Then
        Xd1.HasError = True
        Exit Function
    End If
    If Start = 0 And Recs = 0 Then
        Start = 1
        Recs = Xd2.DataRowCount - Start + 1
        If Not Xd2.RecStart Then
            Copy = True
            Exit Function
        End If
    Else
        If Start <= 1 Then
            Start = 1
            If Not Xd2.RecStart Then
                Copy = True
                Exit Function
            End If
        ElseIf Start > Xd2.DataRowCount Then
            Copy = True
            Exit Function
        Else
            Xd2.RecPos = Start - 1
            If Xd2.HasError Then Exit Function
        End If
        
        If Recs < 0 Then
            Recs = Xd2.DataRowCount - Start + 1
        ElseIf Recs = 0 Then
            Copy = True
            Exit Function
        ElseIf Recs > Xd2.DataRowCount - Start + 1 Then
            Recs = Xd2.DataRowCount - Start + 1
        End If
    End If
    
    If UBound(Fields2) < LBound(Fields2) Then
        MatchFields Xd1, Xd2, ExtractArray, Fields3, , Fields4
    ElseIf Fields2(LBound(Fields2)) = "*" Then
        MatchFields Xd1, Nothing, ExtractArray, Fields3
        If Not CheckFields(Xd1, Xd2, Fields3, Fields4) Then Exit Function
    Else
        If Not ProcessFields(Xd1, Fields2, Fields3) Then Exit Function
        If Not CheckFields(Xd1, Xd2, Fields2, Fields4) Then Exit Function
    End If
    If UBound(Fields3) < LBound(Fields3) Then
        Msg Xd1, "No fields on data"
        Exit Function
    End If

    Do While Xd2.RecNext
            If Xd2.HasError Then
                Xd1.HasError = True
                Xd1.RecClose True
                Exit Function
        End If
        J = J + 1
        If J > Recs Then Exit Do
        If Not Xd1.RecAdd Then
            Exit Function
        End If
        For I = 1 To UBound(Fields3)
            Xd1.RecValue2(Fields3(I)) = Xd2.RecValue2(Fields4(I))
            If Xd1.HasError Then
                Exit Function
            End If
            If Xd2.HasError Then
                Xd1.HasError = True
                Xd1.RecClose True
                Exit Function
            End If
        Next I
    Loop
    If Not Xd1.RecClose Then Exit Function
    Copy = True
End Function

Public Function ProcessFields(Xd As XData, Fields As Variant, FieldCols As Variant, _
    Optional FieldNames As Variant) As Boolean
'Checks and converts an array of fields.
'Checks for duplicates.
'Returns FieldCols and (optionally) FieldNames as 1-based arrays.
    
    Dim I As Long, J As Long
    
    If UBound(Fields) < LBound(Fields) Then
        FieldCols = ExtractArray
        If Not IsMissing(FieldNames) Then _
            FieldNames = ExtractArray
        ProcessFields = True
        Exit Function
    End If
    
    ReDim FieldCols(1 To UBound(Fields) - LBound(Fields) + 1) As Long
    If Not IsMissing(FieldNames) Then _
        ReDim FieldNames(1 To UBound(Fields) - LBound(Fields) + 1) As String
    For I = LBound(Fields) To UBound(Fields)
        J = J + 1
        If Not FieldCheck(Xd, Fields(I)) Then Exit Function
        FieldCols(J) = Xd.Column(Fields(I))
        If Not IsMissing(FieldNames) Then _
            FieldNames(J) = Xd.GetHeader(Fields(I))
    Next I
    
    For I = 1 To UBound(FieldCols)
        For J = I + 1 To UBound(FieldCols)
            If FieldCols(I) = FieldCols(J) Then
                Msg Xd, "Duplicate field name " & Xd.GetHeader(FieldCols(I))
                Exit Function
            End If
        Next J
    Next I
    ProcessFields = True
End Function

Public Function CheckFields(Xd1 As XData, Xd2 As XData, Fields As Variant, _
    Optional XdFields As Variant) As Boolean
'Check fields exist on a Td, Ad or Rd object.
'Optionally returns column positions of fields on Xd2 as a 1-based array
'Xd1 is the caller.
'Xd2 is the target.

    Dim I As Long, J As Long
    
    If Not IsMissing(XdFields) Then _
        ReDim XdFields(1 To UBound(Fields) - LBound(Fields) + 1)
    For I = LBound(Fields) To UBound(Fields)
        If FieldCheck(Xd2, Xd1.GetHeader(Fields(I))) Then
            J = J + 1
            If Not IsMissing(XdFields) Then _
                XdFields(J) = Xd2.Column(Xd1.GetHeader(Fields(I)))
        Else
            Xd1.HasError = True
            Exit Function
        End If
    Next I
    CheckFields = True
End Function

Public Sub MatchFields(Xd1 As XData, Xd2 As XData, Fields As Variant, FieldCols As Variant, _
    Optional FieldNames As Variant, Optional XdCols As Variant)
'Gets all the fields which exist here and on a Td, Ad or Rd object but are not in Fields.
'Xd1 is the caller.
'Xd2 is the target and can be Nothing to ignore this comparison.
'Fields can be empty array to ignore this comparison.
'Returns FieldCols and (optionally) FieldNames and XdCols as 1-based arrays.

    Dim I As Long, J As Long, K As Long
    
    If Xd2 Is Nothing Then Set Xd2 = Xd1
    ReDim FieldCols(1 To Xd1.Columns) As Long
    If Not IsMissing(FieldNames) Then _
        ReDim FieldNames(1 To Xd1.Columns) As String
    If Not IsMissing(XdCols) Then _
        ReDim XdCols(1 To Xd1.Columns) As Long
        
    For I = 1 To Xd1.Columns
        If Xd2.Exists(Xd1.GetHeader(I)) Then
            For J = LBound(Fields) To UBound(Fields)
                If I = Xd1.Column(Fields(J)) Then Exit For
            Next J
            If J > UBound(Fields) Then
                K = K + 1
                FieldCols(K) = I
                If Not IsMissing(FieldNames) Then _
                    FieldNames(K) = Xd1.GetHeader(I)
                If Not IsMissing(XdCols) Then _
                    XdCols(K) = Xd2.Column(Xd1.GetHeader(I))
            End If
        End If
    Next I
    
    If K > 0 Then
        ReDim Preserve FieldCols(1 To K)
        If Not IsMissing(FieldNames) Then _
            ReDim Preserve FieldNames(1 To K)
        If Not IsMissing(XdCols) Then _
            ReDim Preserve XdCols(1 To K)
    Else
        FieldCols = ExtractArray
        If Not IsMissing(FieldNames) Then _
            FieldNames = ExtractArray
        If Not IsMissing(XdCols) Then _
            XdCols = ExtractArray
    End If
End Sub

Public Function CheckDuplicates(Xd As XData, Fields1 As Variant, Fields2 As Variant) As Boolean
'Check for duplicate fields.

    Dim I As Long, J As Long
    
    For I = LBound(Fields1) To UBound(Fields1)
        For J = LBound(Fields2) To UBound(Fields2)
            If Xd.Column(Fields1(I)) = Xd.Column(Fields2(J)) Then
                Msg Xd, "Duplicate field name " & Fields1(I)
                Exit Function
            End If
        Next J
    Next I
    CheckDuplicates = True
End Function

Public Function Exists(Xd1 As XData, Xd2 As XData, Fields As Variant) As Boolean
'Checks that all the fields in the array Fields (which are fields in the calling object Xd1)
'exist on Xd2.
    
    Dim FldName As String
    Dim I As Long
    
    For I = LBound(Fields) To UBound(Fields)
        If Not FieldCheck(Xd1, Fields(I)) Then Exit Function
        FldName = Xd1.GetHeader(Fields(I))
        If Not Xd2.Exists(FldName) Then Exit Function
    Next I
    Exists = True
End Function

Public Function ExtractArray(Optional InputData As Variant) As Variant
'Returns an array from an array, a single element array containing an array or a single value
'Byval removed, InputData now a copy of original ParamArray

    Const SplitChar As String = "|"

    If IsMissing(InputData) Then
        ExtractArray = Split(vbNullString, SplitChar)
    ElseIf IsNull(InputData) Then
        ExtractArray = Split(vbNullString, SplitChar)
    ElseIf IsEmpty(InputData) Then
        ExtractArray = Split(vbNullString, SplitChar)
    ElseIf Not IsArray(InputData) Then
        ExtractArray = Array(InputData)
    ElseIf UBound(InputData) < LBound(InputData) Then
        ExtractArray = Split(vbNullString, SplitChar)
    ElseIf UBound(InputData) > LBound(InputData) Then
        ExtractArray = InputData
    ElseIf IsArray(InputData(UBound(InputData))) Then
        ExtractArray = InputData(UBound(InputData))
    Else
        ExtractArray = Array(InputData(UBound(InputData)))
    End If
End Function

Public Function ConvertType(DataValue As Variant, DataType As VbVarType, Optional Silent As Boolean) As Boolean
'Converts DataValue to specified data type. returns True if successful.
'Converts empty string as if it were Empty.

    Dim DataValue2 As Variant
    
    DataValue2 = DataValue
    If VarType(DataValue) = vbString Then
        If Len(DataValue) = 0 Then DataValue2 = Empty
    End If
    On Error Resume Next
    Select Case DataType
    Case vbInteger
        DataValue = CInt(DataValue2)
    Case vbLong
        DataValue = CLng(DataValue2)
    Case vbSingle
        DataValue = CSng(DataValue2)
    Case vbDouble
        DataValue = CDbl(DataValue2)
    Case vbCurrency
        DataValue = CCur(DataValue2)
    Case vbDate
        DataValue = CDate(DataValue2)
    Case vbString
        DataValue = CStr(DataValue2)
    Case vbBoolean
        DataValue = CBool(DataValue2)
    Case vbByte
        DataValue = CByte(DataValue2)
    Case Else
    End Select
    If Err = 0 Then
        On Error GoTo 0
    Else
        On Error GoTo 0
        If Not Silent Then MsgBox "Bad data type converting " & DataValue
        Exit Function
    End If
    ConvertType = True
End Function

Public Function EncodeXML(ByVal OldText As String) As String
'Encodes special characters for xml.
'Converts back <br/> tags.

    OldText = Replace(OldText, "&", "&amp;")
    OldText = Replace(OldText, """", "&quot;")
    OldText = Replace(OldText, "'", "&apos;")
    OldText = Replace(OldText, "<", "&lt;")
    OldText = Replace(OldText, ">", "&gt;")
    OldText = Replace(OldText, "&lt;br/&gt;", "<br/>")
    EncodeXML = OldText
End Function

Private Function FieldCheck(Xd As XData, Field As Variant) As Boolean
'Check if a field exists. May be name or number.

    If Xd.Exists(Field) Then
        FieldCheck = True
    ElseIf Not Xd.HasError Then
        Msg Xd, "Cannot find field " & Field & " on data"
    End If
End Function

Private Sub Msg(Xd As XData, ErrMsg As String)
'Displays error message unless HasError is already true, then sets HasError.

    If Not Xd.HasError Then MsgBox ErrMsg
    Xd.HasError = True
End Sub

Public Function CreateSource(Xd As XData, Name As String) As Worksheet
'Return data source with a new worksheet if one is open, otherwise create one.
'If Name already exists, it is replaced.

    Dim Wb As Workbook
    Dim Wb2 As Workbook
    Dim Ws As Worksheet
    Dim Ws2 As Worksheet

    For Each Wb In Application.Workbooks
        If Wb.BuiltinDocumentProperties("Subject") = "Export" Then
            On Error Resume Next
            Set Ws2 = Wb.Worksheets(Name)
            On Error GoTo 0

            Set Ws = Wb.Worksheets.Add(After:=Wb.Worksheets(Wb.Worksheets.Count))
            If Not Ws2 Is Nothing Then
                Application.DisplayAlerts = False
                Ws2.Delete
                Application.DisplayAlerts = True
            End If

            On Error Resume Next
            Ws.Name = Name
            If Err = 0 Then
                On Error GoTo 0
            Else
                On Error GoTo 0
                Msg Xd, "Invalid sheet name on export"
                Exit Function
            End If
            Set CreateSource = Ws
            Exit Function
        End If
    Next Wb
    Set Wb2 = ActiveWorkbook
    Set Wb = Application.Workbooks.Add(xlWorksheet)
    Wb.BuiltinDocumentProperties("Subject") = "Export"
    Set Ws = Wb.Worksheets(1)
    On Error Resume Next
    Ws.Name = Name
    If Err = 0 Then
        On Error GoTo 0
    Else
        On Error GoTo 0
        Msg Xd, "Invalid sheet name on export"
        Wb2.Activate
        Exit Function
    End If
    Set CreateSource = Ws
    Wb2.Activate
End Function

Public Function Show(Xd As XData)
'Show unless NoShow is True (which is cleared)

    If ShowMode Then
        If NoShow Then
            NoShow = False
        Else
            ShowIE Xd, 100
        End If
    Else
        NoShow = False
    End If
End Function

Public Sub ShowIE(Xd As XData, Optional MaxRows As Long)
'Show Xd in an internet explorer window.

    Dim IeObj As Object
    Dim IeObjName As String
    Dim IeObjBusy As Boolean
    Dim IeObjReadyState
    Dim I As Long
    Const READYSTATE_COMPLETE As Integer = 4
    Const MaxIters As Long = 25

    CreateWebPage Xd, , , MaxRows
    GoSub GetIeObj
    
    If IeObj Is Nothing Then
        Set IeObj = CreateObject("InternetExplorer.Application") 'Create
        IeObj.AddressBar = False
        IeObj.MenuBar = False
        IeObj.StatusBar = False
        IeObj.Toolbar = False
    End If
    
    For I = 1 To MaxIters
        On Error Resume Next
        IeObj.Navigate2 Environ("Temp") & "\_temp.htm"
        If Err = 0 Then
            On Error GoTo 0
            Exit For
        Else
            On Error GoTo 0
            DoEvents
        End If
    Next I
    If I > MaxIters Then Exit Sub
'    Debug.Print "iters " & I
    
    For I = 1 To MaxIters
        GoSub GetIeObj
        If IeObj Is Nothing Then Exit Sub
        On Error Resume Next
        IeObjBusy = IeObj.Busy
        IeObjReadyState = IeObj.readyState
        If Err = 0 Then
            On Error GoTo 0
            If Not IeObjBusy And IeObjReadyState = READYSTATE_COMPLETE Then Exit For
        Else
            On Error GoTo 0
        End If
        DoEvents
    Next I
    If I > MaxIters Then Exit Sub
'    Debug.Print "iters2 " & I
    
    'This is for putting the IE window on top.
'    If Now - LastShow > 5 / 86400 Then IeObj.Visible = False
    IeObj.Visible = True
    LastShow = Now
    Exit Sub
    
GetIeObj:
    IeObjName = vbNullString
    Set IeObj = Nothing
    For Each IeObj In CreateObject("Shell.Application").Windows() 'Loop to find
        If Not IeObj Is Nothing Then
            On Error Resume Next
            IeObjName = IeObj.Name
            On Error GoTo 0
            If IeObjName Like "*Internet Explorer*" Then
                Exit For
            Else
                Set IeObj = Nothing
            End If
        End If
    Next IeObj
    Return
End Sub

Public Function CreateWebPage(Xd As XData, Optional AdName As String, Optional FontSize As Long = 12, Optional MaxRows As Long = 0)
'Create a web page from Xd.

    Dim MaxWidth As String
    Dim BgColour As String
    Dim Value As Variant
    Dim Value2 As Variant
    Dim FileNo As Long
    Dim I As Long, J As Long, K As Long
    
    If (Xd.Columns <= 5) Then
        MaxWidth = "max-width: 285px;"
    Else
        MaxWidth = "width: " & 60 * Xd.Columns & "px;"
    End If
    
    FileNo = FreeFile()
    Open Environ("Temp") & "\_temp.htm" For Output As #FileNo
    Print #FileNo, "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">"
    Print #FileNo, "<html xmlns=""http://www.w3.org/1999/xhtml"" lang=""en"" xml:lang=""en"">"
    Print #FileNo, "<head>"
    Print #FileNo, "<meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1""/>"
    Print #FileNo, "<meta http-equiv=""X-UA-Compatible"" content=""IE=8""/>"
    Print #FileNo, "<title>" & AdName & "</title>"
    Print #FileNo, "<style type=""text/css"">"
    Print #FileNo, "table, th, td {"
    Print #FileNo, "border: 1px solid gray;"
    Print #FileNo, "border-collapse: collapse;"
    Print #FileNo, "padding: 5px;"
    Print #FileNo, "font-family: Calibri, sans-serif;"
    Print #FileNo, "font-size: " & FontSize & "px;"
    Print #FileNo, "word-wrap: break-word;"
    Print #FileNo, "}"
    Print #FileNo, "</style>"
    Print #FileNo, "</head>"
    
    Print #FileNo, "<body>"
    Print #FileNo, "<table id='Tbl2' style='" & MaxWidth & ";'>"
    Print #FileNo, "<tr>"
    For I = 1 To Xd.Columns
        Print #FileNo, "<th style='background-color:steelblue; color:white;'>" & Xd.GetHeader(I) & "</th>"
    Next I
    Print #FileNo, "</tr>"
    Xd.RecStart
    Do While Xd.RecNext
        If MaxRows > 0 And Xd.RecPos > MaxRows Then Exit Do
        If Xd.RecPos Mod 2 = 0 Then
            BgColour = " style='background-color: #eef5f8;'"
        Else
            BgColour = vbNullString
        End If
        Print #FileNo, "<tr>"
        For I = 1 To Xd.Columns
            Value = Xd.RecValue2(I)
            Select Case VarType(Value)
            Case vbSingle, vbDouble
                Value = Format(Value, "#,##0.####")
                If Right(Value, 1) = "." Then Value = Left(Value, Len(Value) - 1)
            Case vbCurrency
                Value = Format(Value, "#,##0.00")
            Case vbByte
                Value = Format(Value, "0")
            Case vbString
                If Left(Value, 1) = "{" And Right(Value, 1) = "}" And Len(Value) > 2 Then
                    Value = Mid(Value, 2, Len(Value) - 2)
                    Value = Split(Value, ";")
                    For K = 0 To UBound(Value)
                        Value2 = Split(Value(K), ",")
                        For J = 0 To UBound(Value2)
                            If IsNumeric(Value2(J)) Then
                                Value2(J) = Format(Value2(J), "#,##0.####")
                                If Right(Value2(J), 1) = "." Then Value2(J) = Left(Value2(J), Len(Value2(J)) - 1)
                            End If
                        Next J
                        Value(K) = Join(Value2, ",<br/>")
                    Next K
                    Value = "{" & Join(Value, ";<br/>") & "}"
                End If
            End Select
            Print #FileNo, "<td" & BgColour & ">" & EncodeXML(CStr(Value)) & "</td>"
        Next I
        Print #FileNo, "</tr>"
    Loop
    Print #FileNo, "</table>"
    Print #FileNo, "</body>"
    Print #FileNo, "</html>"
    Close #FileNo
End Function

#If ArrMode And RsetMode Then
Public Function RefreshPivot(Xd As XData, Ws As Worksheet, PTName As String) As Boolean
'Create pivot table, or refresh if already exists.

    Dim Rd As New RsetData
    Dim Ad As New ArrData
    Dim Pt As PivotTable
    Dim Pc As PivotCache
    Dim PTFound As Boolean
    
    If Xd.HasError Then Exit Function
    Select Case Xd.ObjType
    Case "RsetData"
        Set Rd = GetRd(Xd)
    Case "ArrData"
        Set Ad = GetAd(Xd)
        If Not Rd.CreateFrom(Ad) Then Exit Function
    Case Else
        Msg Xd, "RefreshPivot only supports RsetData and ArrData objects"
        Exit Function
    End Select
    
    For Each Pt In Ws.PivotTables
        If Pt.Name = PTName Then
            PTFound = True
            Exit For
        End If
    Next Pt
    
    If PTFound Then
        Set Pt.PivotCache.Recordset = Rd.Rsett
        Pt.RefreshTable
    Else
        Set Pc = Ws.Parent.PivotCaches.Create(xlExternal)
        Set Pc.Recordset = Rd.Rsett
        Set Pt = Ws.PivotTables.Add(Pc, Ws.Cells(1, 1), PTName)
    End If
    Rd.RsetClose True
    RefreshPivot = True
    
End Function

#Else
Public Function RefreshPivot(Ad As XData, Ws As Worksheet, PTName As String) As Boolean

    Msg Ad, "RefreshPivot not enabled"
End Function
#End If

Public Sub ReplaceName(NameName As String, NameText As String, Optional IsVisible As Boolean = False)
'Replace a name, removing it first (which seems occasionally necessary)

    On Error Resume Next
    Do
        ThisWorkbook.Names(NameName).Delete
        If Err <> 0 Then Exit Do
    Loop
    On Error GoTo 0
    ThisWorkbook.Names.Add NameName, "=""" & NameText & """", IsVisible
End Sub

#If RsetMode Then
Public Function GetRd(Optional Rd As Object) As RsetData
'Creates a new RsetData object, or extracts the RsetData object from Rd.

    If Rd Is Nothing Then
        Set GetRd = New RsetData
    Else
        Set GetRd = Rd
    End If
End Function

#Else
Public Function GetRd(Optional Rd As Object) As Object
End Function
#End If

#If ArrMode Then
Public Function GetAd(Optional Ad As Object) As ArrData
'Creates a new ArrData object, or extracts the ArrData object from Ad.

    If Ad Is Nothing Then
        Set GetAd = New ArrData
    Else
        Set GetAd = Ad
    End If
End Function

#Else
Public Function GetAd(Optional Ad As Object) As Object
End Function
#End If

#If TableMode Then
Public Function GetTd(Optional Td As Object) As TableData
'Creates a new TableData object, or extracts the TableData object from Td.

    If Td Is Nothing Then
        Set GetTd = New TableData
    Else
        Set GetTd = Td
    End If
End Function

#Else
Public Function GetTd(Optional Td As Object) As Object
End Function
#End If

#If UserFunctionMode Then
Public Function GetUf() As XUserFunctions

    If Uf Is Nothing Then
        Set Uf = New XUserFunctions
    End If
    Set GetUf = Uf
End Function

#Else
Public Function GetUf() As Object

    Set GetUf = Nothing
End Function
#End If

Public Function HasError(ParamArray Xds() As Variant) As Boolean
'Returns True if any of the parameters has HasError=True.

    Dim Xd As XData
    Dim I As Long
    
    For I = 0 To UBound(Xds)
        Set Xd = Xds(I)
        If Xd.HasError Then
            HasError = True
            Exit Function
        End If
    Next I
End Function

Public Function Frac2Num(ByVal x As String) As Double
   Dim P As Integer
   Dim N As Double
   Dim Num As Double
   Dim Den As Double

   x = Trim$(x)
   P = InStr(x, "/")

   If P = 0 Then
      N = Val(x)
   Else
      Den = Val(Mid$(x, P + 1))
      If Den = 0 Then Err.Raise 11  ' Divide by zero
      x = Trim$(Left$(x, P - 1))
      P = InStr(x, " ")

      If P = 0 Then
         Num = Val(x)
      Else
         Num = Val(Mid$(x, P + 1))
         N = Val(Left$(x, P - 1))
      End If
   End If

   If Den <> 0 Then
      N = N + Num / Den
   End If

   Frac2Num = N
End Function

#If RsetMode And ArrMode Then
Public Function RdUpdate(Rd As RsetData, Xd As XData, Keys As Variant, AutoNumField As String) As Boolean
'Updates a RsetData recordset.
'Xd must have matching fields.
'Requires ArrData.

    Dim Ad As New ArrData
    Dim KeyValues
    Dim Lookup As Variant
    Dim FieldCols As Variant
    Dim GroupFns As Variant
    Dim Xs As New XShared
    Dim I As Long
    
    If Xd.HasError Then
        Rd.HasError = True
        Exit Function
    End If
    
    Ad.Create "Matched", vbBoolean
    Ad.AddFields Rd
    MatchFields Rd, Ad, ExtractArray(), FieldCols
    If UBound(FieldCols) <> Rd.Columns Then
        Msg Rd, "Fields do not match in update operation"
        Exit Function
    End If
    Ad.CopyFrom Xd, 0, 0
    Ad.CreateIndex Keys
    ReDim KeyValues(LBound(Keys) To UBound(Keys)) As Variant
    
    Rd.RecStart
    Do While Rd.RecNext
        If Rd.HasError Then Exit Function
        For I = LBound(Keys) To UBound(Keys)
            KeyValues(I) = Rd.RecValue(Keys(I))
        Next I
        Lookup = Ad.Lookup(KeyValues)
        If UBound(Lookup) = 0 Then
            Ad.RecPos = Lookup(0)
            Ad.RecValue("Matched") = True
            For I = 1 To Rd.Columns
                If Rd.RecValue(I) <> Ad.RecValue(I + 1) Then
                    Rd.RecValue(I) = Ad.RecValue(I + 1)
                End If
            Next I
        Else
            Rd.RecDelete
        End If
    Loop
    If Rd.HasError Then Exit Function
    Ad.Filter "Matched", "=FALSE"
    If Ad.HasError Then Exit Function
    
    'Avoid duplicate keys by taking last record.
    'Drop autonumber field.
    If Ad.DataRowCount = 0 Then
        Rd.RecClose
    Else
        MatchFields Ad, Nothing, Keys, FieldCols
        ReDim GroupFns(1 To UBound(FieldCols)) As String
        For I = 1 To UBound(FieldCols)
            GroupFns(I) = "Last"
        Next I
        Ad.Group FieldCols, GroupFns
        If Len(AutoNumField) > 0 Then Ad.DropField AutoNumField
        Rd.CopyFrom Ad, 0, 0
        If Rd.HasError Then Exit Function
    End If
    RdUpdate = True
End Function

#Else
Public Function RdUpdate(Rd As XData, Xd As XData, ParamArray KeySet())

    Msg Rd, "RsetData/ArrData not available"
End Function

#End If

