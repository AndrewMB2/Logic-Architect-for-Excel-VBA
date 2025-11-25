VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RsetData"
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

Private Type FieldType
    Name As String
End Type

Public HasError As Boolean
Public Connx As Object
Public CnFileName As String
Public Rsett As Object
Public DataRowCount As Long
Public ErrNumber As Long 'set for connect or recordset open or execute
Dim Xs As New XShared
Dim objCatalog As Object
Dim StrCon As String
Dim DefCursorLocation As adCursorLocation
Dim Dirty As Boolean
Dim FieldList As Dictionary
Dim Fld() As FieldType '1-based
Private Const adPersistXML As Integer = 1
Private Const adKeyPrimary As Integer = 1
Private Const adColNullable As Integer = 2
Private Const adFldIsNullable As Long = &H20
Private Const adExecuteNoRecords = &H80
Private Const adSchemaIndexes = 12
Private Const adSchemaPrimaryKeys = 28
Private Type TypeGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "OLE32.DLL" (pGuid As TypeGUID) As Long
#Else
Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As TypeGUID) As Long
#End If

Public Enum adDataType
    adSmallInt = 2
    adInteger = 3
    adSingle = 4
    adDouble = 5
    adCurrency = 6
    adDate = 7
    adDBDate = 133
    adNumeric = 131
    adDBTimeStamp = 135
    adChar = 129
    adVarChar = 200
    adLongVarChar = 201
    adVarWChar = 202
    adLongVarWChar = 203
    adBoolean = 11
    adTinyInt = 16
End Enum
Public Enum adCursorType
    adOpenForwardOnly = 0
    adOpenkeyset = 1
    adOpenDynamic = 2
    adOpenStatic = 3
'    adOpenUnspecified = -1
End Enum
Public Enum adLockType
    adLockReadOnly = 1
    adLockPessimistic = 2
    adLockOptimistic = 3
    adLockBatchOptimistic = 4
    adLockUnspecified = -1
End Enum
Public Enum adObjectState
    adStateClosed = 0
    adStateOpen = 1
    adStateConnecting = 2
    adStateExecuting = 4
    adStateFetching = 8
End Enum
Public Enum CnType
    cnExcel12xml = 1
    cnExcel12 = 2
    cnAccess = 3
    cnSQLServer = 4
    cnMySql = 5
    cnDBISAM = 6
End Enum
Public Enum adCursorLocation
    adUseClient = 3
    adUseNone = 1
    adUseServer = 2
End Enum

Function ObjType() As String
    ObjType = "RsetData"
End Function

Private Function XData_ObjType() As String
    XData_ObjType = "RsetData"
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

Public Function CreateDatabase(Filename As String) As Boolean
'Creates a new Access database using ADOX and stores the catalog and the connection.

    If HasError Then Exit Function
    GetConnectionString Filename, cnAccess
    Set objCatalog = CreateObject("ADOX.Catalog")
    On Error Resume Next
    Set Connx = objCatalog.Create(StrCon)
    If Err = 0 Then
        On Error GoTo 0
        CreateDatabase = True
    Else
        MsgBox Err.Description
        On Error GoTo 0
        Set objCatalog = Nothing
        Set Connx = Nothing
        HasError = True
        Exit Function
    End If
End Function

Public Function CreateTable(Xd As XData, TblName As String, Optional LongTextFields As Variant, Optional Keys As Variant)
'Creates an Access table using ADOX.
'Connection must have been previously opened.
'Sets long text fields to Memo, otherwise text fields to Text.
'Sets primary key. Field is required if a primary key but not otherwise.

    Dim Tbl As Object
    Dim TCol As Object
    Dim Tkey As Object
    Dim Fields2 As Variant
    Dim Fields3 As Variant 'Long text fields as field numbers
    Dim Fields4() As Boolean 'Long text fields as array containing True or False
    Dim Fields5 As Variant 'Keys as field numbers
    Dim Fields6 As Variant 'Keys as field names
    Dim Fields7() As Boolean 'Keys as array containing True or False
    Dim I As Long
    
    If HasError Then Exit Function
    Fields2 = Xs.ExtractArray(LongTextFields)
    If Not Xs.ProcessFields(Xd, Fields2, Fields3) Then Exit Function
    ReDim Fields4(1 To Xd.Columns)
    For I = 1 To UBound(Fields3)
        Fields4(Fields3(I)) = True
    Next I
    
    Fields2 = Xs.ExtractArray(Keys)
    If Not Xs.ProcessFields(Xd, Fields2, Fields5, Fields6) Then Exit Function
    ReDim Fields7(1 To Xd.Columns)
    For I = 1 To UBound(Fields5)
        Fields7(Fields5(I)) = True
    Next I
    
    On Error Resume Next
    objCatalog.Tables.Delete TblName
    Select Case Err
    Case 0, 3265
        On Error GoTo 0
    Case Else
        MsgBox Err.Description
        On Error GoTo 0
        Set objCatalog = Nothing
        Set Connx = Nothing
        HasError = True
        Exit Function
    End Select
    
    Set Tbl = CreateObject("ADOX.Table")
    Tbl.Name = TblName
    For I = 1 To Xd.Columns
        On Error Resume Next
        Tbl.Columns.Append Xd.GetHeader(I), ConvertVtype(Xd.GetType(I), Fields4(I))
        Set TCol = Tbl.Columns(I - 1)
        If Not Fields7(I) Then TCol.Attributes = adColNullable
        If Err = 0 Then
            On Error GoTo 0
        Else
            MsgBox Err.Description
            On Error GoTo 0
            Set objCatalog = Nothing
            Set Connx = Nothing
            HasError = True
            Exit Function
        End If
    Next I
    
    If UBound(Fields6) > 0 Then
        Set Tkey = CreateObject("ADOX.Key")
        With Tkey
            .Name = "PrimaryKey"
            .Type = adKeyPrimary
            For I = 1 To UBound(Fields6)
                .Columns.Append Fields6(I)
            Next I
        End With
        On Error Resume Next
        Tbl.Keys.Append Tkey
        If Err = 0 Then
            On Error GoTo 0
        Else
            MsgBox Err.Description
            On Error GoTo 0
            Set objCatalog = Nothing
            Set Connx = Nothing
            HasError = True
            Exit Function
        End If
    End If
    
    On Error Resume Next
    objCatalog.Tables.Append Tbl
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox Err.Description
        On Error GoTo 0
        Set objCatalog = Nothing
        Set Connx = Nothing
        HasError = True
        Exit Function
    End If
    
    RsetOpen "SELECT * FROM " & TblName
    CopyFrom Xd, 0, 0
    If HasError Then Exit Function
    
    Dirty = False
    CreateTable = True
End Function

Public Function Connect(Filename As String, ConnType As CnType, Optional IMEX As Integer = 0, _
    Optional InitialCatalog As String = vbNullString, Optional UserId As String = vbNullString, _
    Optional Password As String = vbNullString, Optional ByVal CursorLocation As adCursorLocation = -1) As Boolean
    
'Create a connection.
'IMEX=1 produces a read-only connection.

    If HasError Then Exit Function
    ErrNumber = 0
    GetConnectionString Filename, ConnType, IMEX, InitialCatalog, UserId, Password
    Set Connx = CreateObject("ADODB.Connection")
    If CursorLocation < 0 Then CursorLocation = DefCursorLocation
    Connx.CursorLocation = CursorLocation
    On Error Resume Next
    Connx.Open StrCon
    If Err = 0 Then
        On Error GoTo 0
        Connect = True
    Else
        ErrNumber = Err.Number
        MsgBox Err.Description
        On Error GoTo 0
        Set objCatalog = Nothing
        Set Connx = Nothing
        HasError = True
        Exit Function
    End If
    Set objCatalog = CreateObject("ADOX.Catalog")
    Set objCatalog.ActiveConnection = Connx
End Function

Public Function ConnectFrom(ConnectionName As String, _
    Optional ByVal CursorLocation As adCursorLocation = -1) As Boolean
'Create a connection from an existing Workbook OLEDB connection.

    Dim Wc As WorkbookConnection

    If HasError Then Exit Function
    ErrNumber = 0
    On Error Resume Next
    Set Wc = ThisWorkbook.Connections(ConnectionName)
    If Err = 0 Then
        On Error GoTo 0
    Else
        On Error GoTo 0
        MsgBox "Cannot connect to " & ConnectionName
        HasError = True
        Exit Function
    End If
    
    StrCon = Wc.OLEDBConnection.Connection
    If Left(StrCon, 6) = "OLEDB;" Then
        StrCon = Mid(StrCon, 7)
    Else
        MsgBox "Not a OLEDB connection string"
        HasError = True
        Exit Function
    End If
        
    Set Connx = CreateObject("ADODB.Connection")
    If CursorLocation < 0 Then
        If Right(Wc.Description, 3) Like "(#)" Then
            DefCursorLocation = CInt(Mid(Wc.Description, Len(Wc.Description) - 1, 1))
        Else
            MsgBox "No cursor location"
            HasError = True
            Exit Function
        End If
        CursorLocation = DefCursorLocation
    End If
    Connx.CursorLocation = CursorLocation
    On Error Resume Next
    Connx.Open StrCon
    If Err = 0 Then
        On Error GoTo 0
        ConnectFrom = True
    Else
        ErrNumber = Err.Number
        MsgBox Err.Description
        On Error GoTo 0
        Set objCatalog = Nothing
        Set Connx = Nothing
        HasError = True
        Exit Function
    End If
    Set objCatalog = CreateObject("ADOX.Catalog")
    Set objCatalog.ActiveConnection = Connx
End Function

Public Function GetConnectionString(Filename As String, ConnType As CnType, Optional IMEX As Integer = 0, _
    Optional InitialCatalog As String = vbNullString, Optional UserId As String = vbNullString, _
    Optional Password As String = vbNullString) As Boolean
'Create a connection string.
    
    If HasError Then Exit Function
    CnFileName = Filename
    Select Case ConnType
    Case cnExcel12xml
        StrCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=""" & Filename & """;" & _
            "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=" & IMEX & """;"
        DefCursorLocation = adUseServer
    Case cnExcel12
        StrCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=""" & Filename & """;" & _
            "Extended Properties=""Excel 12.0;HDR=YES;IMEX=" & IMEX & """;"
        DefCursorLocation = adUseServer
    Case cnAccess
        StrCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=""" & Filename & """;"
        DefCursorLocation = adUseServer
    Case cnSQLServer
        StrCon = "Provider=SQLOLEDB.1;"
        If Len(Password) > 0 Then StrCon = StrCon & "Password=" & Password & ";"
        StrCon = StrCon & "Persist Security Info=True;"
        If Len(UserId) > 0 Then StrCon = StrCon & "User ID=" & UserId & ";"
        If Len(InitialCatalog) > 0 Then StrCon = StrCon & "Initial Catalog=" & InitialCatalog & ";"
        StrCon = StrCon & "Data Source=" & Filename & ";"
        StrCon = StrCon & "Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;"
        StrCon = StrCon & "Workstation ID=" & Environ$("ComputerName") & ";"
        StrCon = StrCon & "Use Encryption for Data=False;Tag with column collation when possible=False"
        DefCursorLocation = adUseClient
    Case cnMySql
        StrCon = "Driver={MySQL ODBC 8.0 Unicode Driver};Server=" & Filename & "; " & _
            "Database=" & InitialCatalog & ";" & _
            "UID=" & UserId & ";PWD=" & Password & ";FOUND_ROWS=1;READTIMEOUT=20;WRITETIMEOUT=20;"
        DefCursorLocation = adUseClient
    Case cnDBISAM
        StrCon = "DRIVER={DBISAM 4 ODBC Driver};ConnectionType=Local;CatalogName=" & Filename & ";"
        DefCursorLocation = adUseClient
    End Select
    GetConnectionString = True
End Function

Public Function CreateConnection(ConnectionName As String, Optional Description As String, _
    Optional iCmdtype As Integer = 2, Optional CommandText As String = vbNullString) As Boolean
'Create a workbook connection from a previously created connection string.
'iCmdtype default 2 is for SQL. 3 for Table, 1 for OLAP cube.
'Places default cursor location in brackets at the end of the description.

    Dim Wc As WorkbookConnection
    
    If HasError Then Exit Function
    If Len(StrCon) = 0 Then
        MsgBox "No connection string"
        HasError = True
        Exit Function
    End If
    
    For Each Wc In ThisWorkbook.Connections
        If Wc.Name = ConnectionName Then
            Wc.Delete
        End If
    Next Wc
    
    On Error Resume Next
    ThisWorkbook.Connections.Add ConnectionName, Description & " (" & DefCursorLocation & ")", "OLEDB;" & StrCon, _
        CommandText, iCmdtype
    If Err = 0 Then
        On Error GoTo 0
    Else
        On Error GoTo 0
        MsgBox "Cannot create connection"
        HasError = True
        Exit Function
    End If
    CreateConnection = True
End Function

Public Function ConnectionStatus() As Boolean
'Check connection status

    If Not Connx Is Nothing Then
        Select Case Connx.Properties("DBMS Name")
        Case "MySQL"
            If Connx.State = adStateOpen Then
                If Connx.Properties("Connection Status") = 1 Then
                    ConnectionStatus = True
                End If
            End If
        Case Else
            ConnectionStatus = True
        End Select
    End If
        
End Function

Public Function GetQueryTable(SQL As String, TableRange As Range, Optional TableName As String, _
    Optional TableStyle As String, Optional AdjustColumnWidth As Boolean = True) As Boolean
'Create a query table.

    Dim ResultRange As Range
    Dim Qt As QueryTable
    Dim ListObj As ListObject
    Dim Wc As WorkbookConnection
    
    If HasError Then Exit Function
    For Each Qt In TableRange.Parent.QueryTables
        Qt.Delete
    Next Qt
    If Len(TableName) > 0 Then
        On Error Resume Next
        Set ListObj = TableRange.Parent.ListObjects(TableName)
        On Error GoTo 0
        If Not ListObj Is Nothing Then ListObj.Delete
    End If

    With TableRange.Parent.QueryTables.Add("OLEDB;" & StrCon, TableRange)
        .CommandType = xlCmdSql
        .CommandText = SQL
        .PreserveFormatting = True
        .FieldNames = True
        .FillAdjacentFormulas = True
        .RefreshOnFileOpen = False
        .AdjustColumnWidth = AdjustColumnWidth
        .RefreshStyle = xlInsertDeleteCells
        .MaintainConnection = False
        .BackgroundQuery = False
        .EnableEditing = True
        .EnableRefresh = False
        .SaveData = False
        .SavePassword = False
        Set Wc = .WorkbookConnection
        On Error GoTo IfError
        .Refresh BackgroundQuery:=False
        On Error GoTo 0
        If HasError Then
            Wc.Delete
            .Delete
            Exit Function
        End If
        Set ResultRange = .ResultRange
        .Delete
    End With
    If Len(TableName) > 0 Then
        With TableRange.Parent.ListObjects.Add(xlSrcRange, ResultRange, , xlYes)
            .Name = TableName
            .TableStyle = TableStyle
            .Range.AutoFilter
        End With
    End If
    GetQueryTable = True
    Exit Function
    
IfError:
    MsgBox "Error attempting query" & vbLf & Err.Description
    HasError = True
    Resume Next
End Function

Public Function RsetOpen(SQL As String, Optional CursorType As adCursorType = adOpenStatic, _
    Optional LockType As adLockType = adLockOptimistic, Optional CacheSize As Variant, _
    Optional Silent As Boolean = False) As Boolean
'Open recordset from SQL.
'For updating, use adLockOptimistic
    
    Dim I As Long
    
    If HasError Then Exit Function
    ErrNumber = 0
    Dirty = False
    Set Rsett = CreateObject("ADODB.Recordset")
    If Not IsMissing(CacheSize) Then Rsett.CacheSize = CacheSize
    On Error Resume Next
    Rsett.Open SQL, Connx, CursorType, LockType
    If Err = 0 Then
        On Error GoTo 0
        If Rsett.RecordCount >= 0 Then
            DataRowCount = Rsett.RecordCount
            RsetOpen = True
        Else
            If Silent Then
                Debug.Print "Record count not set"
            Else
                MsgBox "Record count not set"
            End If
            Set Rsett = Nothing
            DataRowCount = 0
            HasError = True
            Exit Function
        End If
    Else
        ErrNumber = Err.Number
        If Silent Then
            Debug.Print Err.Description
        Else
            MsgBox Err.Description
        End If
        On Error GoTo 0
        Set Rsett = Nothing
        DataRowCount = 0
        HasError = True
        Exit Function
    End If
    
    Set FieldList = New Dictionary
    FieldList.CompareMode = TextCompare
    ReDim Fld(1 To Rsett.Fields.Count)
    For I = 1 To Rsett.Fields.Count
        FieldList.Add Rsett.Fields(I - 1).Name, I
        Fld(I).Name = Rsett.Fields(I - 1).Name
    Next I
End Function

Private Function RsetOpenSchema(QueryType As Long) As Boolean
'Returns a recordset with schema infomation.
'DataRowCount is not returned.
'Available for Access but may not be available for all databases.
    
    Dim I As Long
    
    If HasError Then Exit Function
    ErrNumber = 0
    Dirty = False
    On Error Resume Next
    Set Rsett = Nothing
    Set Rsett = Connx.OpenSchema(QueryType)
    If Err = 0 Then
        On Error GoTo 0
        If Rsett.RecordCount >= 0 Then
            DataRowCount = Rsett.RecordCount
        Else
            DataRowCount = 0
        End If
    Else
        ErrNumber = Err.Number
        MsgBox Err.Description
        On Error GoTo 0
        Set Rsett = Nothing
        DataRowCount = 0
        HasError = True
        Exit Function
    End If
    
    Set FieldList = New Dictionary
    FieldList.CompareMode = TextCompare
    ReDim Fld(1 To Rsett.Fields.Count)
    For I = 1 To Rsett.Fields.Count
        FieldList.Add Rsett.Fields(I - 1).Name, I
        Fld(I).Name = Rsett.Fields(I - 1).Name
    Next I
    RsetOpenSchema = True
End Function

Public Function GetKeys(TableName As String) As Variant
'Returns the keys of the specified table.
'Requires ADODB OpenSchema(adSchemaPrimaryKeys) - works for Access.
'Replaces Rsett

    Dim KeySet As Variant
    Dim Xs As New XShared
    
    ReDim KeySet(0 To 0) As String
    RsetOpenSchema adSchemaPrimaryKeys
    If HasError Then Exit Function
    RecStart
    Do While RecNext
        If RecValue("TABLE_NAME") = TableName Then
            KeySet(UBound(KeySet)) = RecValue("COLUMN_NAME")
            ReDim Preserve KeySet(UBound(KeySet) + 1)
        End If
    Loop
    If HasError Then Exit Function
    RsetClose
    ReDim Preserve KeySet(UBound(KeySet) - 1)
    GetKeys = KeySet
End Function

Public Function Execute(SQL As String) As Long
'Executes a SQL statement.
'Returns the number of records affected.

    Dim Records As Long

    If HasError Then Exit Function
    On Error Resume Next
    Connx.Execute SQL, Records, adExecuteNoRecords
    If Err = 0 Then
        On Error GoTo 0
    Else
        ErrNumber = Err.Number
        MsgBox Err.Description
        On Error GoTo 0
        HasError = True
        Exit Function
    End If
    Execute = Records
    
End Function

Public Function Create(Headers, RecordTypes) As Boolean
'Create a new recordset.
    
    Dim Vtype As Integer
    Dim I As Long
    
    If HasError Then Exit Function
    Dirty = False
    Set objCatalog = Nothing
    Set Connx = Nothing
    CnFileName = vbNullString
    StrCon = vbNullString
    
    If Not IsArray(Headers) Then
        MsgBox "Headers must be an array"
        HasError = True
        Exit Function
    ElseIf UBound(Headers) < LBound(Headers) Then
        MsgBox "No fields"
        HasError = True
        Exit Function
    ElseIf Not IsArray(RecordTypes) Then
        MsgBox "RecordTypes must be an array"
        HasError = True
        Exit Function
    ElseIf UBound(RecordTypes) <> UBound(Headers) Or _
        LBound(RecordTypes) <> LBound(Headers) Then
        MsgBox "RecordTypes and Headers must match"
        HasError = True
        Exit Function
    End If
    
    Set FieldList = New Dictionary
    FieldList.CompareMode = TextCompare
    ReDim Fld(1 To UBound(Headers) - LBound(Headers) + 1)
    Set Rsett = CreateObject("ADODB.Recordset")
    DataRowCount = 0
    For I = LBound(Headers) To UBound(Headers)
        If FieldList.Exists(Headers(I)) Then
            MsgBox "Duplicate header " & Headers(I)
            HasError = True
            Exit Function
        Else
            FieldList.Add Headers(I), I - LBound(Headers) + 1
        End If
        Fld(I - LBound(Headers) + 1).Name = Headers(I)
        Vtype = ConvertVtype(CInt(RecordTypes(I)))
        If Vtype = adLongVarWChar Then
            Rsett.Fields.Append Headers(I), Vtype, 536870910
        Else
            Rsett.Fields.Append Headers(I), Vtype
        End If
    Next I
    Rsett.Open , , , adLockOptimistic
    Create = True
End Function

Public Function CreateFrom(Xd As XData) As Boolean
'Create a recordset from a TableData or ArrData object.
'Leaves recordset positioned at BOF.
'If Xd is TableData, returns False if no data (in which case no recordset is created)

    Dim xlxml As Object
    Dim Headers As Variant
    Dim DataTypes As Variant
    Dim I As Long
    
    If HasError Then Exit Function
    If Xd.HasError Then
        HasError = True
        Exit Function
    End If
    Select Case Xd.ObjType
    Case "TableData"
        Dirty = False
        Set objCatalog = Nothing
        Set Connx = Nothing
        CnFileName = vbNullString
        StrCon = vbNullString
        
        Set FieldList = New Dictionary
        FieldList.CompareMode = TextCompare
        ReDim Fld(1 To Xd.Columns)
        For I = 1 To Xd.Columns
            FieldList.Add Xd.GetHeader(I), I
            Fld(I).Name = Xd.GetHeader(I)
        Next I
            
        If Xd.DataRowCount > 0 Then
            Set Rsett = CreateObject("ADODB.Recordset")
            Set xlxml = CreateObject("MSXML2.DOMDocument.6.0")
            xlxml.LoadXML Xs.GetTd(Xd).GetRange.Value(xlRangeValueMSPersistXML)
            Rsett.Open xlxml, , adOpenForwardOnly, adLockOptimistic
            If Rsett.RecordCount = Xd.DataRowCount Then
                Rsett.MovePrevious
                CreateFrom = True
                DataRowCount = Xd.DataRowCount
            Else
                MsgBox "Cannot create recordset"
                HasError = True
                Rsett.Close
                Set Rsett = Nothing
                DataRowCount = 0
            End If
        End If
    Case "ArrData"
        ReDim Headers(1 To Xd.Columns) As String
        ReDim DataTypes(1 To Xd.Columns) As VbVarType
        For I = 1 To Xd.Columns
            Headers(I) = Xd.GetHeader(I)
            DataTypes(I) = Xd.GetType(I)
        Next I
        If Not Create(Headers, DataTypes) Then Exit Function
        If Not CopyFrom(Xd, 0, 0) Then Exit Function
        RecStart
        CreateFrom = True
    End Select

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
'Returns VarType of field (Index is 1-based) (-1 if an error).

    Dim DataType As VbVarType
        
    DataType = XData_GetType(Index)
    GetType = DataType
End Function


Private Function XData_GetType(Index As Variant) As Variant

    Dim AdType As adDataType
    Dim DataType As VbVarType
        
    If HasError Then
        XData_GetType = -1
        Exit Function
    End If
    If FieldCheck(Index) Then
        AdType = Rsett.Fields(FieldCol(Index) - 1).Type
        DataType = ConvertAdtype(AdType)
        If DataType = -1 Then
            MsgBox "AdType not recognised " & AdType
            HasError = True
        End If
        XData_GetType = DataType
    Else
        XData_GetType = -1
    End If
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

Function RecEOF() As Boolean
'Check if recordset is at EOF.

    If HasError Then Exit Function
    On Error Resume Next
    RecEOF = Rsett.EOF
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox "Error checking for EOF" & vbLf & Err.Description
        On Error GoTo 0
        HasError = True
    End If
End Function

Function RecBOF() As Boolean
'Check if recordset is at BOF.

    If HasError Then Exit Function
    On Error Resume Next
    RecBOF = Rsett.Bof
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox "Error checking for BOF" & vbLf & Err.Description
        On Error GoTo 0
        HasError = True
    End If
End Function

Public Function RecStart() As Boolean

    RecStart = XData_RecStart
End Function

Private Function XData_RecStart() As Boolean
'Move to the BOF position of the recordset and returns True, or False if no records.
    
    If HasError Then Exit Function
    If Not (Rsett.Bof And Rsett.EOF) Then
        If HasError Then Exit Function
        On Error Resume Next
        Rsett.MoveFirst
        If Err <> 0 Then
            MsgBox "Error moving to start" & vbLf & Err.Description
            HasError = True
            Exit Function
        End If
        If Not Rsett.Bof Then Rsett.MovePrevious
        If Err = 0 Then
            On Error GoTo 0
            XData_RecStart = True
        Else
            MsgBox "Error moving to start" & vbLf & Err.Description
            On Error GoTo 0
            HasError = True
        End If
    End If
End Function

Function RecEnd() As Boolean
'Move to the EOF position of the recordset and returns True, or False if no records.
    
    If HasError Then Exit Function
    If Not (Rsett.Bof And Rsett.EOF) Then
        If HasError Then Exit Function
        On Error Resume Next
        Rsett.MoveLast
        If Err <> 0 Then
            MsgBox "Error moving to end" & vbLf & Err.Description
            HasError = True
            Exit Function
        End If
        If Not Rsett.EOF Then Rsett.MoveNext
        If Err = 0 Then
            On Error GoTo 0
            RecEnd = True
        Else
            MsgBox "Error moving to end" & Err.Description
            On Error GoTo 0
            HasError = True
        End If
    End If
End Function

Public Function RecNext() As Boolean

    RecNext = XData_RecNext
End Function

Private Function XData_RecNext() As Boolean
'Move to the next record of the recordset and returns True, or False if at EOF.
'Skip deleted records.
    
    If HasError Then Exit Function
    Do
        If RecEOF Then
            Exit Function
        Else
            On Error Resume Next
            Rsett.MoveNext
            If Err = 0 Then
                On Error GoTo 0
            Else
                MsgBox "Error moving to next record" & vbLf & Err.Description
                Rsett.CancelUpdate
                On Error GoTo 0
                HasError = True
                Exit Function
            End If
            If RecEOF Then
                Exit Function
            ElseIf RecExists Then
                XData_RecNext = True
                Exit Function
            End If
        End If
    Loop
End Function

Function RecPrev() As Boolean
'Move to the previous record of the recordset and returns True, or False if at BOF
    
    If HasError Then Exit Function
    If Not Rsett.Bof Then
        On Error Resume Next
        Rsett.MovePrevious
        If Err = 0 Then
            On Error GoTo 0
        Else
            MsgBox "Error moving to previous record" & vbLf & Err.Description
            Rsett.CancelUpdate
            On Error GoTo 0
            HasError = True
            Exit Function
        End If
        RecPrev = Not RecBOF
    End If
End Function

Public Property Get RecPos() As Long

    If HasError Then Exit Property
    On Error Resume Next
    RecPos = Rsett.AbsolutePosition
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox "Error retrieving record position" & vbLf & Err.Description
        Rsett.CancelUpdate
        On Error GoTo 0
        HasError = True
    End If
End Property

Private Property Get XData_RecPos() As Long
'Returns absolute position.

    If HasError Then Exit Property
    On Error Resume Next
    XData_RecPos = Rsett.AbsolutePosition
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox "Error retrieving record position" & vbLf & Err.Description
        Rsett.CancelUpdate
        On Error GoTo 0
        HasError = True
    End If
End Property

Public Property Let RecPos(Pos As Long)

    If HasError Then Exit Property
    On Error Resume Next
    Rsett.AbsolutePosition = Pos
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox "Error setting record position" & vbLf & Err.Description
        Rsett.CancelUpdate
        On Error GoTo 0
        HasError = True
    End If
End Property

Private Property Let XData_RecPos(Pos As Long)
'Sets absolute position.

    If HasError Then Exit Property
    On Error Resume Next
    Rsett.AbsolutePosition = Pos
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox "Error setting record position" & vbLf & Err.Description
        Rsett.CancelUpdate
        On Error GoTo 0
        HasError = True
    End If
End Property

Property Get RecBookMark() As Variant
'Returns bookmark.

    If HasError Then Exit Property
    On Error Resume Next
    RecBookMark = Rsett.BookMark
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox "Error retrieving bookmark" & vbLf & Err.Description
        Rsett.CancelUpdate
        On Error GoTo 0
        HasError = True
    End If
End Property

Property Let RecBookMark(BookMark As Variant)
'Sets bookmark.

    If HasError Then Exit Property
    On Error Resume Next
    Rsett.BookMark = BookMark
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox "Error setting bookmark" & vbLf & Err.Description
        Rsett.CancelUpdate
        On Error GoTo 0
        HasError = True
    End If
End Property

Public Function RecAdd() As Boolean

    RecAdd = XData_RecAdd
End Function

Private Function XData_RecAdd() As Boolean
'Add a record. Returns True if successful.
'Cancels update on failure.
    
    If HasError Then Exit Function
    On Error Resume Next
    Rsett.AddNew
    If Err = 0 Then
        On Error GoTo 0
        DataRowCount = DataRowCount + 1
        Dirty = True
        XData_RecAdd = True
    Else
        MsgBox "Error adding record to recordset" & vbLf & Err.Description
        Rsett.CancelUpdate
        On Error GoTo 0
        Dirty = False
        HasError = True
    End If
End Function

Public Function RecClose(Optional Cancel As Boolean = False) As Boolean

    RecClose = XData_RecClose(Cancel)
End Function

Private Function XData_RecClose(Optional Cancel As Boolean = False) As Boolean
'Write changes to data source, or cancels if Cancel is true.
'Can be run when HasError is true.

    If Not Connx Is Nothing Then
        If Dirty Then
            If RecEOF Then
                If RecBOF Then
                    If HasError Then Exit Function
                    Dirty = False
                    XData_RecClose = True
                    Exit Function
                Else
                    RecStart
                    RecNext
                End If
            End If
            If HasError Then Exit Function
            
            On Error Resume Next
            If Cancel Then
                Rsett.CancelUpdate
                If Err = 0 Then
                    On Error GoTo 0
                    XData_RecClose = True
                Else
                    MsgBox "Error closing recordset" & vbLf & Err.Description
                    On Error GoTo 0
                    HasError = True
                    Exit Function
                End If
                
            Else
                Rsett.Update
                If Err = 0 Then
                    On Error GoTo 0
                    XData_RecClose = True
                Else
                    MsgBox "Error closing recordset" & vbLf & Err.Description
                    Rsett.CancelUpdate
                    On Error GoTo 0
                    HasError = True
                    Exit Function
                End If
            End If
            DataRowCount = Rsett.RecordCount
            Dirty = False
        Else
            XData_RecClose = True
        End If
    Else
        XData_RecClose = True
    End If
End Function

Public Function RecDelete() As Boolean

    RecDelete = XData_RecDelete
End Function

Private Function XData_RecDelete() As Boolean
'Delete record - Update required afterwards.

    If HasError Then Exit Function
    On Error Resume Next
    Rsett.Delete
    If Err = 0 Then
        On Error GoTo 0
        DataRowCount = DataRowCount - 1
        Dirty = True
        XData_RecDelete = True
    Else
        MsgBox "Error deleting record from recordset" & vbLf & Err.Description
        Rsett.CancelUpdate
        On Error GoTo 0
        Dirty = False
        HasError = True
    End If
End Function

Public Function RecExists() As Boolean

    Dim Rval As Variant

    If HasError Then Exit Function
    On Error Resume Next
    Rval = Rsett.Fields(0).Value
    If Err = 0 Then RecExists = True
    On Error GoTo 0
End Function

Public Sub RsetClose(Optional CloseConnection As Boolean = False)
'Close the recordset if any and optionally any connection.
'Can be run when HasError is True.

    If Not Rsett Is Nothing Then
        If Rsett.State <> adStateClosed Then
            Rsett.Close
        End If
        Set Rsett = Nothing
        DataRowCount = 0
    End If
    If CloseConnection Then
        If Not Connx Is Nothing Then
            If Connx.State <> adStateClosed Then
                Connx.Close
            End If
            Set Connx = Nothing
        End If
    End If
End Sub

Property Get RecValue(Field As Variant) As Variant
'Get the value of a field in the current record.

    Dim Field2 As Long

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Field2 = FieldCol(Field)
        RecValue = XData_RecValue2(Field2)
    End If
End Property

Property Get XData_RecValue(Field As Variant) As Variant
'Get the value of a field in the current record.

    Dim Field2 As Long

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Field2 = FieldCol(Field)
        XData_RecValue = XData_RecValue2(Field2)
    End If
End Property

Private Property Get XData_RecValue2(Field As Variant) As Variant
'Get the value of a field in the current record.
'Assumes Field is a valid column number.

    If HasError Then Exit Property
    With Rsett.Fields(Field - 1)
        On Error Resume Next
        XData_RecValue2 = .Value
        If IsNull(XData_RecValue2) Then XData_RecValue2 = Empty
        If Err = 0 Then
            On Error GoTo 0
            If .Type = adVarWChar And Not IsEmpty(XData_RecValue2) Then
                XData_RecValue2 = Left(CStr(XData_RecValue2), .DefinedSize)
            End If
        Else
            MsgBox "Error retrieving recordset field " & Field & vbLf & Err.Description
            On Error GoTo 0
            HasError = True
        End If
    End With
End Property

Property Let RecValue(Field As Variant, FieldValue As Variant)
Attribute RecValue.VB_UserMemId = 0
'Set the value of a field in the current record.

    Dim Field2 As Long

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Field2 = FieldCol(Field)
        XData_RecValue2(Field2) = FieldValue
    End If
End Property

Property Let XData_RecValue(Field As Variant, FieldValue As Variant)
'Set the value of a field in the current record.

    Dim Field2 As Long

    If HasError Then Exit Property
    If FieldCheck(Field) Then
        Field2 = FieldCol(Field)
        XData_RecValue2(Field2) = FieldValue
    End If
End Property

Private Property Let XData_RecValue2(Field As Variant, FieldValue As Variant)
'Set the value of a field in the current record.
'Assumes Field is a valid column number.
'Cancels update on failure.

    If HasError Then Exit Property
    On Error Resume Next
    If IsEmpty(FieldValue) And CBool(Rsett.Fields(Field - 1).Attributes And adFldIsNullable) Then
        Rsett.Fields(Field - 1).Value = Null
    Else
        Rsett.Fields(Field - 1).Value = FieldValue
    End If
    If Err = 0 Then
        Dirty = True
        On Error GoTo 0
    Else
        MsgBox "Error creating recordset field " & Field & " value " & CStr(FieldValue) & vbLf & Err.Description
        On Error GoTo 0
        Rsett.CancelUpdate
        DataRowCount = Rsett.RecordCount
        Dirty = False
        HasError = True
    End If
End Property

Public Function GetArray(Optional ByVal MaxCols As Long = 0, Optional ByVal MaxRows As Long = 0, _
    Optional Default = Empty) As Variant
'Returns the recordset as an array (1-based) suitable for return to a worksheet.
'If row or column counts are greater than maxima specified, returns #N/A.
'If row or column counts are less than maxima specified, fills the unused cells with Default specified.
'Truncates long text to 255 characters unless all the fields and Default (if specified) are text.

    Dim DataSet As Variant
    Dim DataSet2 As Variant
    Dim Value As Variant
    Dim MixedType As Boolean
    Dim I As Long, J As Long
    
    If HasError Then
        GetArray = Xs.ExtractArray
        Exit Function
    End If
    If MaxCols = 0 Then MaxCols = UBound(Fld)
    If MaxRows = 0 Then MaxRows = DataRowCount + 1
    For I = 1 To Rsett.Fields.Count
        Select Case Rsett.Fields(I - 1).Type
        Case adVarWChar, adLongVarWChar
        Case Else
            MixedType = True
        End Select
    Next I
    If Not IsEmpty(Default) Then
        If VarType(Default) <> vbString Then MixedType = True
    End If
    
    If MixedType Then
        ReDim DataSet2(1 To MaxRows, 1 To MaxCols) As Variant
    Else
        ReDim DataSet2(1 To MaxRows, 1 To MaxCols) As String
    End If
    For I = 1 To MaxCols
        If I <= UBound(Fld) Then
            DataSet2(1, I) = Fld(I).Name
        Else
            DataSet2(1, I) = Default
        End If
    Next I
    
    If MaxCols < UBound(Fld) Or MaxRows < DataRowCount + 1 Then
        For I = 1 To MaxCols
            For J = 2 To MaxRows
                DataSet2(J, I) = CVErr(xlErrNA)
            Next J
        Next I
        Debug.Print "Space allowed is too small to return results of query"
    
    Else
        If Rsett.RecordCount > 0 Then
            Rsett.MoveFirst
            DataSet = Rsett.GetRows
        End If
        
        For I = 1 To MaxCols
            For J = 2 To MaxRows
                If I <= UBound(Fld) And J <= Rsett.RecordCount + 1 Then
                    Value = DataSet(LBound(DataSet, 1) + I - 1, LBound(DataSet, 2) + J - 2)
                    Select Case VarType(Value)
                    Case vbString
                        If MixedType And Len(Value) > 255 Then Value = Left(Value, 255)
                    Case vbNull
                        Value = vbNullString
                    End Select
                    DataSet2(J, I) = Value
                Else
                    DataSet2(J, I) = Default
                End If
            Next J
        Next I
    End If
    GetArray = DataSet2
End Function

Public Function Export(Name As String, ParamArray Field()) As Boolean
'Export all or named fields to external workbook.
'Exports string fields as text regardless of content.

    Dim Ws As Worksheet
    Dim Wb2 As Workbook
    Dim Ws2 As Worksheet
    Dim Field2 As Variant
    Dim Fieldx As Variant
    Dim DataSet As Variant
    Dim DataSet2 As Variant
    Dim I As Long, J As Long
    
    If HasError Then Exit Function
    Set Ws = Xs.CreateSource(Me, Name)
    If Ws Is Nothing Then Exit Function
    DoEvents
    Fieldx = Field
    Field2 = Xs.ExtractArray(Fieldx)
    
    If Rsett.RecordCount = 0 Then
        If UBound(Field2) < 0 Then
            For I = 1 To UBound(Fld)
                Ws.Cells(1, I).Value = Fld(I).Name
            Next I
        Else
            For I = LBound(Field2) To UBound(Field2)
                If Not FieldCheck(Field2(I)) Then
                    HasError = True
                    Exit Function
                End If
                Ws.Cells(1, I - LBound(Field2) + 1) = FieldName(Field2(I))
            Next I
        End If
        
    Else
        Rsett.MoveFirst
        If UBound(Field2) < 0 Then
            Set Wb2 = ActiveWorkbook
            Set Ws2 = ActiveSheet
            Ws.Parent.Activate
            Ws.Activate
            For I = 1 To UBound(Fld)
                Ws.Cells(1, I).Value = Fld(I).Name
            Next I
            Ws.Cells(2, 1).CopyFromRecordset Rsett
            Ws.Cells(2, 1).Resize(Rsett.RecordCount).EntireRow.RowHeight = Ws.Rows(1).RowHeight
            Wb2.Activate
            Ws2.Activate
        Else
            For I = LBound(Field2) To UBound(Field2)
                If Not FieldCheck(Field2(I)) Then
                    HasError = True
                    Exit Function
                End If
                Ws.Cells(1, I - LBound(Field2) + 1) = FieldName(Field2(I))
                Field2(I) = FieldCol(Field2(I)) - 1
            Next I
            DataSet = Rsett.GetRows(, , Field2)
            ReDim DataSet2(LBound(DataSet, 2) To UBound(DataSet, 2), LBound(DataSet, 1) To UBound(DataSet, 1))
            For I = LBound(DataSet, 1) To UBound(DataSet, 1)
                For J = LBound(DataSet, 2) To UBound(DataSet, 2)
                    If VarType(DataSet(I, J)) = vbString Then
                        DataSet2(J, I) = "'" & DataSet(I, J)
                    Else
                        DataSet2(J, I) = DataSet(I, J)
                    End If
                Next J
            Next I
            With Ws.Cells(2, 1).Resize(Rsett.RecordCount, UBound(Field2) - LBound(Field2) + 1)
                .Value = DataSet2
                .EntireRow.RowHeight = Ws.Rows(1).RowHeight
            End With
        End If
    End If
    Export = True
End Function

Public Function ExpOpen(Filename As String) As Boolean
'Reopen exported workbook.

    Dim Wb As Workbook
    
    If HasError Then Exit Function
    Set Wb = ActiveWorkbook
    On Error Resume Next
    Workbooks.Open Filename, False, True
    If Err = 0 Then
        On Error GoTo 0
    Else
        On Error GoTo 0
        MsgBox "Cannot open exported file"
        HasError = True
        Exit Function
    End If
    Wb.Activate
    ExpOpen = True
End Function

Function ExpAdjust() As Boolean
'Moves rows to the top of the exported sheet as necessary for data type recognition.
'Also handles long text (>255 characters).

    Dim Wb2 As Workbook
    Dim Ws As Worksheet
    Dim Rcount As Long
    Dim Ccount As Long
    Dim Changed As Boolean
    Dim I As Long, J As Long, K As Long
    Const TypeGuessRows As Integer = 8
    
    If HasError Then Exit Function
    Set Wb2 = GetSource
    If Wb2 Is Nothing Then Exit Function
    
    Set Ws = Wb2.Worksheets(Wb2.Worksheets.Count)
    With Ws.UsedRange
        Rcount = .Rows.Count
        Ccount = .Columns.Count
    End With
    
    With Ws
        Do
            Changed = False
            For I = 1 To Ccount
                If IsEmpty(.Cells(2, I)) Then
                    J = .Cells(2, I).End(xlDown).Row
                Else
                    J = 2
                End If
                If VarType(.Cells(J, I)) = vbString Then
                    J = 0
                    On Error Resume Next
                    J = Application.Evaluate("=MATCH(TRUE,LEN(" & _
                        .Cells(2, I).Resize(Rcount - 1, 1).Address(, , , True) & ")>255,0)") + 1
                    On Error GoTo 0
                End If
                If J > TypeGuessRows + 1 And J <= Rcount Then
                    K = K + 1
                    If K > TypeGuessRows Then
                        MsgBox "Too many columns to adjust"
                        HasError = True
                        Exit Function
                    End If
                    .Range(.Cells(J, 1), .Cells(J, Ccount)).Cut
                    .Range(.Cells(K + 1, 1), .Cells(K + 1, Ccount)).Insert xlShiftDown
                    Application.CutCopyMode = False
                    Changed = True
                    Exit For
                End If
            Next I
        Loop While Changed
    End With
    ExpAdjust = True
End Function

Function ExpClose(Optional TempDays As Long = 1) As String
'Close export workbook and save as Excel binary workbook, returning path and name.
'Cleans up any files more than TempDays day old.

    Dim Wb2 As Workbook
    Dim FilePath As String
    Dim Filename As String
    Dim OldNames() As String
    Dim I As Long
    Const SubFldr As String = "\{5F7C1E42-9770-4305-9A47-B3939E267E9B}\"
    
    If HasError Then Exit Function
    Set Wb2 = GetSource
    If Wb2 Is Nothing Then Exit Function
    
    'Clear out old workbooks (except any still in an existing connection string) and save the new one.
    FilePath = Environ("Temp") & SubFldr
    If Len(Dir(FilePath, vbDirectory)) = 0 Then MkDir FilePath
    ReDim OldNames(1 To 1)
    Filename = Dir(FilePath)
    If Len(Filename) > 0 Then
        OldNames(1) = Filename
        Do
            Filename = Dir()
            If Len(Filename) = 0 Then Exit Do
            ReDim Preserve OldNames(1 To UBound(OldNames) + 1)
            OldNames(UBound(OldNames)) = Filename
        Loop
    End If
    For I = 1 To UBound(OldNames)
        If OldNames(I) Like "*.xls?" Then
            If Now - FileDateTime(FilePath & OldNames(I)) > TempDays Then
                On Error Resume Next
                Kill FilePath & OldNames(I)
                On Error GoTo 0
            End If
        End If
    Next I
    Filename = X_GetGUID & ".xlsb"
    Calculate 'Excel 2016 bug
    Wb2.SaveAs FilePath & Filename, xlExcel12, , , True
    Wb2.Close False
    ExpClose = FilePath & Filename
End Function

Public Function SaveXML(Filename As Variant) As Boolean
'Saves as a XML file in PersistXML format.
'Do not save more than once to the same destination.

    If HasError Then Exit Function
    
    On Error Resume Next
    Kill Filename
    Err.Clear
    Rsett.Save Filename, adPersistXML
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

Public Function OpenXML(Filename As Variant) As Boolean
'Opens a XML file in PersistXML format.
'Recordset is opened forward only.

    Dim I As Long

    If HasError Then Exit Function

    Set Rsett = CreateObject("ADODB.Recordset")
    On Error Resume Next
    Rsett.Open Filename, , adOpenForwardOnly, adLockOptimistic
    If Err = 0 Then
        On Error GoTo 0
    Else
        MsgBox Err.Description
        On Error GoTo 0
        HasError = True
        Set Rsett = Nothing
        Exit Function
    End If
        
    If Rsett.Bof And Rsett.EOF Then
        'No records
    Else
        Rsett.MovePrevious
    End If
    DataRowCount = Rsett.RecordCount
    Dirty = False
    Set objCatalog = Nothing
    Set Connx = Nothing
    CnFileName = vbNullString
    StrCon = vbNullString
    Set FieldList = New Dictionary
    FieldList.CompareMode = TextCompare
    ReDim Fld(1 To Rsett.Fields.Count)
    For I = 1 To Rsett.Fields.Count
        FieldList.Add Rsett.Fields(I - 1).Name, I
        Fld(I).Name = Rsett.Fields(I - 1).Name
    Next I
    OpenXML = True
End Function

Public Function UpdateFrom(Xd As XData, ParamArray KeySet()) As Boolean
'Updates open recordsset from Xd
'Keyset 0 indicates first field is autonumber key.

    Dim Xs As New XShared
    Dim Keyset2 As Variant
    Dim Keyset3 As Variant
    Dim AutoNumField As String
    
    If HasError Then Exit Function
    Keyset2 = KeySet
    Keyset3 = Xs.ExtractArray(Keyset2)
    If UBound(Keyset3) = 0 And Keyset3(0) = 0 Then
        Keyset3 = Array(GetHeader(1))
        AutoNumField = GetHeader(1)
    End If
    Xs.RdUpdate Me, Xd, Keyset3, AutoNumField
    If Not HasError Then UpdateFrom = True

End Function

Public Function RefreshPivot(Ws As Worksheet, PTName As String) As Boolean
'Create pivot table, or refresh if already exists.

    Dim Xs As New XShared

    If HasError Then Exit Function
    RefreshPivot = Xs.RefreshPivot(Me, Ws, PTName)
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

Private Function ConvertVtype(Vtype As VbVarType, Optional LongText As Boolean = True) As Integer
'Convert variable type to that approriate to a recordset.

    Select Case Vtype
    Case vbInteger
        ConvertVtype = adSmallInt
    Case vbLong
        ConvertVtype = adInteger
    Case vbSingle
        ConvertVtype = adSingle
    Case vbDouble
        ConvertVtype = adDouble
    Case vbCurrency
        ConvertVtype = adCurrency
    Case vbDate
        ConvertVtype = adDate
    Case vbString
        ConvertVtype = IIf(LongText, adLongVarWChar, adVarWChar)
    Case vbBoolean
        ConvertVtype = adBoolean
    End Select
End Function

Private Function ConvertAdtype(AdType As adDataType) As VbVarType
'Convert recordset datatype back to VbVarType

 Select Case AdType
    Case adSmallInt
        ConvertAdtype = vbInteger
    Case adInteger
        ConvertAdtype = vbLong
    Case adSingle
        ConvertAdtype = vbSingle
    Case adDouble, adNumeric
        ConvertAdtype = vbDouble
    Case adCurrency
        ConvertAdtype = vbCurrency
    Case adDate, adDBTimeStamp, adDBDate
        ConvertAdtype = vbDate
    Case adChar, adVarChar, adLongVarChar, adVarWChar, adLongVarWChar
        ConvertAdtype = vbString
    Case adBoolean
        ConvertAdtype = vbBoolean
    Case adTinyInt
        ConvertAdtype = vbByte
    Case Else
        ConvertAdtype = -1
    End Select
End Function

Private Function GetSource() As Workbook
'Returns data source if one is open with error if not found.

    Dim Wb As Workbook
    Dim Wb2 As Workbook

    For Each Wb In Application.Workbooks
        If Wb.BuiltinDocumentProperties("Subject") = "Export" Then
            Set Wb2 = Wb
            Exit For
        End If
    Next Wb
    If Wb2 Is Nothing Then
        MsgBox "File not found"
        HasError = True
        Exit Function
    End If
    Set GetSource = Wb2
End Function

Public Function GetExport() As Workbook
'Returns data source if one is open or nothing if not found.

    Dim Wb As Workbook
    Dim Wb2 As Workbook

    For Each Wb In Application.Workbooks
        If Wb.BuiltinDocumentProperties("Subject") = "Export" Then
            Set Wb2 = Wb
            Exit For
        End If
    Next Wb
    Set GetExport = Wb2
End Function

Private Function X_GetGUID() As String
'(c) 2000 Gus Molina

    Dim udtGUID As TypeGUID
    
    If (CoCreateGuid(udtGUID) = 0) Then
        X_GetGUID = "{" & _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & "-" & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & "-" & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & "-" & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & "-" & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7)) & "}"
    End If
End Function

Private Function XData_CreateIndex(ParamArray Fields() As Variant) As Boolean

    MsgBox "Indexing not available"
    HasError = True
End Function

Private Function XData_Lookup(ParamArray Values() As Variant) As Variant

    MsgBox "Indexing not available"
    HasError = True
End Function

Private Function XData_RemoveIndex() As Boolean

    MsgBox "Indexing not available"
    HasError = True
End Function
