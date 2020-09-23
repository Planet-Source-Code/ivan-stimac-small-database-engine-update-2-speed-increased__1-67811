VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
'       MODULE      : clsDatabase
'       DESCRIPTION : database engine
'       VERSION     : 1.0.2
'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
'
'
'
'--------- how it works (sorry for my english) ----------------------------------------------
'
'   :: DB FILE
'       * at the start of file is strAppIdent = MDBE, we use this to check is opened file
'         correct db file
'       * after this is encrypted user name and passwor to access (each uses 10 spaces)
'       * then data thad contains list of tables, table columns, where starts (in file)
'         first and last data, this uses 5000 spaces (enought for about 50 tables)
'         I'm decided to reserve const space because it's more simple if all tables in
'         one place
'       * then we have all data. All row contains place where previous row starts
'         so this si link betveen data.
'
'   :: WRITING DATA
'       * all data writes to end of file and stores start of previous data
'
'   :: WRITING TABLES
'       * we first from table informations create one string:
'         tblName[(collNum)(colName1,colName2,...):(firstItem,lastItem)]tbl2Name[(colName1,colName2,...),:(firstItem,lastItem)],...
'        * then we encrypt that string and write it in file, at position after user name and possword
'
'   :: READING DATA
'       * for reading data we need to enter table name. Then program reads all data and store
'         them to recordset.
'       * principe: in tables defs string is data that tell us where starts last row.
'                   then we read this row. This row contains data obout this where starts
'                   previous data and search previous row for all rows until we got 0
'                   -> first row have as previous row 0
'
'   :: DELETING DATA
'       * this is problematic. You can't delete data from middle of file. So I got one idea.
'       * And idea is: create one more data for each row. If this data is + then row in not
'         deleted, if is * then it's deleted.
'       * but data is still in file, it's just invisible for reading. The problem solve is
'         clsDBAdvanced. This module reads all tables and record and writes them to another file
'         So if all success you can delete old file and copy new as this name
'
'
'--------- changes ----------------------------------------------------------------------------
'
'   [] 2/7/2007
'       * fix slow reading
'       * fix slow deleting
'       * increase write speed
'       * fix bug that create some strange table names if you try more then once
'         to create tables (happens because I didn't reset all values)
'

Private ffDB As Integer

'data for access to db
Private Type dbParams
    dbUsername As String * 10
    dbPassword As String * 10
End Type

'list of tables
'   tblName[(collNum)(colName1,colName2,...):(firstItem,lastItem)]tbl2Name[(colName1,colName2,...),:(firstItem,lastItem)],...
Private Type dbTbls
    'there is 5000 spaces reserved for tables defs
    dbTables As String * 5000
End Type
'
'read row
Private Type dbTableRow
    tblFieldsData As String
    tblPrevRow As Long
    tblFieldDeleted As String * 1 ' -> * deleted, + not deleted
End Type

'to store username and password
Private mDBParams As dbParams
'string with tables defs
Private mDBTbls As dbTbls
'array that countains table names, colum names... for each table
Private mTblParams() As tblParams
'to store row data (prew row, data of this row)
Private mTblRow As dbTableRow
'where in file starts string with tables defs
Private tblDefStart As Integer
'
Private haveTables As Boolean

'to store where each row starts in file (need for delete rows)
Private lngColl2() As Long
'
Private dbFileName As String, dbAcsPass As String * 10, dbAcsUsrName As String * 10
'Private mColl3 As New Collection
'this is on file start, we using tis data to check is right file open
Private Const strAppIdent As String * 4 = "MDBE"



'----------------------------------------------------------------------------------------
' function messages:
'   -1 : invalid file
'   -2 : fail to open
'   -3 : invalid username or password
'    0 : open success
Public Function OpenDB(ByVal dbFile As String, Optional dbAccess As String = "") As Integer
    'On Error Resume Next
    Dim mstrAppIdent As String * 4
    Dim strUsrName As String, strUsrPass As String
    'get input username and pass
    parseDbAccess dbAccess, strUsrName, strUsrPass
    
    ffDB = FreeFile
    Open dbFile For Binary As #ffDB
    'error when atempt to open file
    If Err.Number <> 0 Then
        OpenDB = -2
        Exit Function
    End If
    'check is opened invalid file
    Get ffDB, , mstrAppIdent
    If mstrAppIdent <> strAppIdent Then
        OpenDB = -1
        Close #ffDB
        ffDB = 0
        Exit Function
    'if all ok then continue read
    Else
        'read username and password
        mDBParams.dbPassword = Space(10)
        mDBParams.dbUsername = Space(10)
        Get ffDB, , mDBParams
        mDBParams.dbPassword = Decrypt(mDBParams.dbPassword)
        mDBParams.dbUsername = Decrypt(mDBParams.dbUsername)
        '
        dbFileName = dbFile
        dbAcsPass = mDBParams.dbPassword
        dbAcsUsrName = mDBParams.dbUsername
        '
        'then check it
        If strUsrName <> Trim(mDBParams.dbUsername) Or strUsrPass <> Trim(mDBParams.dbPassword) Then
            OpenDB = -3
            Close #ffDB
            ffDB = 0
            Exit Function
        End If
        '
        'if username and password is correct then read tables
        '   note: there is 5000 spaces reserved for table list and
        '         table definitions
        
        ' mDBTbls.dbTables = Space(5000)
        tblDefStart = Loc(ffDB) + 1
        Get ffDB, , mDBTbls
        mDBTbls.dbTables = Decrypt(Trim(mDBTbls.dbTables), 50)
        'then we need parse this string and add tables to variable
        parseTables mDBTbls.dbTables, mTblParams
    End If
End Function

'messages:
'   -1 : cant write
'    0 : success
Public Function CreateDB(ByVal dbFile As String, Optional dbAccess As String = "") As Integer
    On Error Resume Next
    Dim mFFdb As Integer
    mFFdb = FreeFile
    'Kill dbFile
    'load username and password into dbParams variable
    If dbAccess = "" Then
        mDBParams.dbPassword = Space(10)
        mDBParams.dbUsername = Space(10)
    Else
        parseDbAccess dbAccess, mDBParams.dbUsername, mDBParams.dbPassword
        mDBParams.dbUsername = Crypt(mDBParams.dbUsername)
        mDBParams.dbPassword = Crypt(mDBParams.dbPassword)
    End If
    
    'write appIdent, dbParams (user name and pass) and tables into file
    Open dbFile For Binary As mFFdb
        Put mFFdb, , strAppIdent
        Put mFFdb, , mDBParams
        Put mFFdb, , mDBTbls
    Close mFFdb
    'messages
    If Err.Number <> 0 Then
        ffDB = 0
        CreateDB = -1
    Else
        CreateDB = 0
    End If
End Function
' function messages:
'   -1 : there is no selected DB
'   -2 : fail to create
'   -3 : table exists
'    0 : execute success
Public Function CreateTable(ByVal strSQL As String) As Integer
    On Error Resume Next
    Dim tmpStr As String, fReturn As String
    Dim i As Integer, strChr As String, tmpTblName  As String
    If ffDB = 0 Then
        CreateTable = -1
    Else
        fReturn = getTblParamsFromSQL(strSQL)
        If fReturn <> "INVALID SQL QUERY" Then
            'find table name
            tmpTblName = ""
            For i = 1 To Len(fReturn)
                strChr = Mid(fReturn, i, 1)
                If strChr = "[" Then Exit For
                tmpTblName = tmpTblName & strChr
            Next i
            
           ' check if table with this name exists
            For i = 0 To UBound(mTblParams)
                If Err.Number <> 0 Then Exit For
                If mTblParams(i).tblName = Trim(tmpTblName) Then
                    'MsgBox mTblParams(i).tblName & vbCrLf & Trim(tmpTblName)
                    CreateTable = -3
                    Exit Function
                End If
            Next i
            Err.Clear
            'save table
            'MsgBox Asc(Mid(mDBTbls.dbTables, 1, 1))
            If Asc(Mid(mDBTbls.dbTables, 1, 1)) = 0 Then
                tmpStr = fReturn
            Else
                tmpStr = Trim(mDBTbls.dbTables) & fReturn
            End If
            mDBTbls.dbTables = Crypt(tmpStr, 50)
            Put ffDB, tblDefStart, mDBTbls.dbTables
            mDBTbls.dbTables = tmpStr 'Decrypt(tmpStr, 50)
        Else
            CreateTable = -2
        End If
    End If
    
    
End Function
' function messages:
'   -1 : there is no selected DB
'   -2 : fail to write
'   -3 : invalid table name
'    0 : execute success
Public Function InsertRecord(ByVal strSQL As String) As Integer
    'On Error Resume Next
    Dim fReturn As String, tmpTblName As String
    Dim mTblInd As Integer
    If ffDB = 0 Then
        InsertRecord = -1
    Else
        fReturn = parseInsetIntoSql(strSQL, mTblParams, mTblInd)
        If mTblInd >= 0 Then
            mTblRow.tblFieldsData = fReturn
            mTblRow.tblPrevRow = mTblParams(mTblInd).tblLastData
            mTblRow.tblFieldDeleted = "+"
            'MsgBox mTblRow.tblPrevRow
            '
            mTblParams(mTblInd).tblLastData = LOF(ffDB) + 1
            If mTblParams(mTblInd).tblFirstData = 0 Then mTblParams(mTblInd).tblFirstData = LOF(ffDB) + 1
            If mTblParams(mTblInd).tblLastData = 0 Then mTblParams(mTblInd).tblLastData = LOF(ffDB) + 1
            'save row
            Put ffDB, LOF(ffDB) + 1, mTblRow
'            'save tables def to string
'            mDBTbls.dbTables = CreateStrDbTablesDef(mTblParams)
'            tmpTblName = mDBTbls.dbTables
'            mDBTbls.dbTables = Crypt(Trim(mDBTbls.dbTables), 50)
'            'write to file
'            Put ffDB, tblDefStart, mDBTbls.dbTables
'            'restore (encriped) tables def to string
'            mDBTbls.dbTables = tmpTblName
'            'save tbl info
        Else
            InsertRecord = -3
        End If
    End If
End Function

Public Sub UpdateAfterWrite()
    On Error Resume Next
    'save tables def to string
    mDBTbls.dbTables = CreateStrDbTablesDef(mTblParams)
    tmpTblName = mDBTbls.dbTables
    mDBTbls.dbTables = Crypt(Trim(mDBTbls.dbTables), 50)
    'write to file
    Put ffDB, tblDefStart, mDBTbls.dbTables
    'restore (encriped) tables def to string
    mDBTbls.dbTables = tmpTblName
    'save tbl info
End Sub

' function messages:
'   -1 : there is no selected DB
'   -2 : fail to delete
'    0 : execute success
'   whereWhat: (colName1, colName2) VALUES (='val1',<>'val2')
Public Function DeleteRecords(ByVal tblName As String, Optional whereWhat As String = "") As Integer
    'On Error Resume Next
    Dim mTmpRecSet As New clsRecSet
    Dim tmpStr As String, strCols() As String, strVals() As String, strOperator() As String
    Dim strChr As String
    Dim i As Long, j As Long, k As Long, m As Long, readLevel As Integer
    Dim tblIndex As Integer
    
    
    If ffDB = 0 Then
        DeleteRecord = -1
    Else
        'read all data from table to record set
        'MsgBox "IDE1"
        Set mTmpRecSet = Me.ReadRecords(tblName, True)
        'MsgBox "IDE2"
        'rewrite all deleted data (deleted = *)
        If whereWhat = "" Then
            For i = 0 To mTmpRecSet.Rows - 1
                tmpStr = ""
                For j = 0 To mTmpRecSet.Columns - 1
                    If j > 0 Then tmpStr = tmpStr & "|"
                    tmpStr = tmpStr & "'" & mTmpRecSet.Data(i, j) & "'"
                Next j
                mTblRow.tblFieldDeleted = "*"
                mTblRow.tblFieldsData = tmpStr
                If i < mTmpRecSet.Rows - 1 Then
                    mTblRow.tblPrevRow = mColl2.Item(i + 2)
                Else
                    mTblRow.tblPrevRow = 0
                End If
                'Put #ffDB, mColl2.Item(i + 1), mTblRow
                Put #ffDB, lngColl2(i), mTblRow
            Next i
        Else
            'get data count in strWhat
            j = 0
            For i = 1 To Len(whereWhat)
                If Mid(whereWhat, i, 1) = "'" Then j = j + 1
            Next i
            j = j / 2
            
            ReDim strCols(j)
            ReDim strVals(j)
            ReDim strOperator(j)
            j = 0
            k = 0
            readLevel = 0
            'save to vars
            For i = 1 To Len(whereWhat)
                strChr = Mid(whereWhat, i, 1)
                If strChr = "(" Then
                    readLevel = readLevel + 1
                    tmpStr = ""
                    z = 0
                ElseIf strChr = "," Or strChr = ")" Then
                    If readLevel = 1 Then
                        strCols(z) = Trim(tmpStr)
                        ''''''''''''''''''''''''''''''''
                        'MsgBox tmpStr
                    End If
                    tmpStr = ""
                    z = z + 1
                ElseIf strChr = "," Then
                    z = z + 1
                ElseIf strChr = "'" Then
                    k = k + 1
                    If k = 1 Then
                        strOperator(z) = Trim(tmpStr)
                        ''''''''''''''''''''''''''''''''
                        'MsgBox tmpStr
                    ElseIf k = 2 Then
                        strVals(z) = Trim(tmpStr)
                        ''''''''''''''''''''''''''''''''
                        'MsgBox tmpStr
                        k = 0
                    End If
                    tmpStr = ""
                ElseIf strChr = "(" Then
                    tmpStr = ""
                Else
                    tmpStr = tmpStr & strChr
                End If
            Next i
            '
            'find table index
            For i = 0 To UBound(mTblParams)
                If mTblParams(i).tblName = tblName Then Exit For
            Next i
            
            If i = UBound(mTblParams) Then
                'can't find table
                DeleteRecords = -2
            Else
                tblIndex = i
            End If
            
            
            For i = 0 To mTmpRecSet.Rows - 1
                tmpStr = ""
                mTblRow.tblFieldDeleted = "*"
                For j = 0 To mTmpRecSet.Columns - 1
                    If j > 0 Then tmpStr = tmpStr & "|"
                    tmpStr = tmpStr & "'" & mTmpRecSet.Data(i, j) & "'"
                    '
                    For k = 0 To mTblParams(tblIndex).tblColCnt - 1
                        For m = 0 To UBound(strCols) - 1
                            If strCols(m) = mTblParams(tblIndex).tblCols(k) Then
                                Select Case strOperator(m)
                                    Case "="
                                        If mTmpRecSet.Data(i, m) <> strVals(m) Then
                                            mTblRow.tblFieldDeleted = "+"
                                            GoTo exitCheck
                                        End If
                                    Case ">"
                                        'MsgBox mTmpRecSet.Data(i, m) & vbCrLf & Val(strVals(m)) & vbCrLf & i & "  -   " & j
                                        If Val(mTmpRecSet.Data(i, m)) < Val(strVals(m)) Or mTmpRecSet.Data(i, m) = strVals(m) Then
                                            mTblRow.tblFieldDeleted = "+"
                                            
                                            GoTo exitCheck
                                        End If
                                    Case "<"
                                        If Val(mTmpRecSet.Data(i, m)) > Val(strVals(m)) Or mTmpRecSet.Data(i, m) = strVals(m) Then
                                            mTblRow.tblFieldDeleted = "+"
                                            GoTo exitCheck
                                        End If
                                    Case "<>"
                                        If mTmpRecSet.Data(i, m) = strVals(m) Then
                                            mTblRow.tblFieldDeleted = "+"
                                            GoTo exitCheck
                                        End If
                                End Select
                            End If
                        Next m
                    Next k
exitCheck:
                Next j
                'check is for delete

                '
                mTblRow.tblFieldsData = tmpStr
                If i < mTmpRecSet.Rows - 1 Then
                    'mTblRow.tblPrevRow = mColl2.Item(i + 2)
                    mTblRow.tblPrevRow = lngColl2(i + 1)
                Else
                    mTblRow.tblPrevRow = 0
                End If
                'Put #ffDB, mColl2.Item(i + 1), mTblRow
                Put #ffDB, lngColl2(i), mTblRow
            Next i
            'now we must check all fields
            
        End If
    End If
End Function

Public Function ReadRecords(ByVal tblName As String, Optional forDelete As Boolean = False) As clsRecSet
    Dim i As Long, j As Long, k As Long, c As Long
    Dim lastS As Integer
    Dim prev As Long, rowNum As Long
    Dim strChr As String, tmpStr As String
    'Dim mColl As New Collection
    Dim mRecSet As New clsRecSet
    Dim strColl() As String
    
    For i = 0 To UBound(mTblParams) - 1
        If tblName = mTblParams(i).tblName Then Exit For
    Next i
    '
    If mTblParams(i).tblLastData = 0 Then Exit Function
    'MsgBox mTblParams(i).tblName
    '

    mRecSet.Columns = mTblParams(i).tblColCnt

    
    'first count rows
    prev = mTblParams(i).tblLastData
    Set mColl2 = Nothing
    rowNum = 0
    Do While prev > 0
        Get #ffDB, prev, mTblRow
        'if not deleted then
        If mTblRow.tblFieldDeleted <> "*" Then
            rowNum = rowNum + 1
        End If
        prev = mTblRow.tblPrevRow
    Loop
    'then read them to string array
    If rowNum < 1 Then Exit Function
    'set arrays lenght
    ReDim strColl(rowNum - 1)
    ReDim lngColl2(rowNum - 1)
    '
    mRecSet.Rows = rowNum
    rowNum = 0
    prev = mTblParams(i).tblLastData
    Do While prev > 0
        Get #ffDB, prev, mTblRow
        'if not deleted then
        If mTblRow.tblFieldDeleted <> "*" Then
            If forDelete = True Then lngColl2(rowNum) = prev 'mColl2.Add prev
            strColl(rowNum) = mTblRow.tblFieldsData
            rowNum = rowNum + 1
        End If
        prev = mTblRow.tblPrevRow
    Loop
    
    'before (used collection) - this part is no problematic
    'problem is reading from collection (it's slow)
'    Do While prev > 0
'        Get #ffDB, prev, mTblRow
'        'if not deleted then
'        If mTblRow.tblFieldDeleted <> "*" Then
'            If forDelete = True Then mColl2.Add prev
'            mColl.Add mTblRow.tblFieldsData
'            rowNum = rowNum + 1
'        End If
'        prev = mTblRow.tblPrevRow
'    Loop
    
   ' mRecSet.Rows = mColl.Count
    
    
    'there has been problem (reading from collection is much slower than array)
    tmpStr = ""
    'at the end parse and save to record set
    For i = 0 To UBound(strColl)  'mColl.Count
        j = 0
        k = 1
        c = 0
        'InStr is faster than checking all characters
        Do While k > 0
            k = InStr(k, strColl(i), "'") 'mColl.Item(i), "'")
            If k > 0 Then
                c = c + 1
                If c = 2 Then
                    j = j + 1
                    mRecSet.Data(i, j - 1) = Mid(strColl(i), lastS, k - lastS)
                    c = 0
                Else
                    lastS = k + 1
                End If
                k = k + 1
            End If
        Loop
        'before (slow)
'        For k = 1 To Len(mColl.Item(i))
'            strChr = Mid(mColl.Item(i), k, 1)
'            If strChr = "'" Then
'                c = c + 1
'                If c = 2 Then
'                    j = j + 1
'                    'If j < 0 Then MsgBox j
'                    mRecSet.Data(i - 1, j - 1) = tmpStr
'                    'MsgBox tmpStr
'                    c = 0
'                End If
'                tmpStr = ""
'            Else
'                tmpStr = tmpStr & strChr
'            End If
'        Next k
        
        
    Next i
    'Tick = GetTickCount - Tick
    'MsgBox "Elapsed " & Format$(Tick / 1000, "0.000000000") & " second", 32
    
    Erase strColl
    Set ReadRecords = mRecSet
End Function
Public Function CloseDB() As Integer
    Erase lngColl2
    Erase mTblParams
    mDBTbls.dbTables = ""
    tblDefStart = 0
    haveTables = False
    Close #ffDB
End Function

Public Function getTablesCount() As Integer
    getTablesCount = UBound(mTblParams)
End Function
Public Function getTableName(ByVal Index As Integer) As String
    getTableName = mTblParams(Index).tblName
End Function
Public Function getColCount(ByVal Index As Integer) As Integer
    getColCount = mTblParams(Index).tblColCnt
End Function
Public Function getColName(ByVal Index As Integer, ByVal ColIndex As Integer) As String
    getColName = mTblParams(Index).tblCols(ColIndex)
End Function

' function messages:
'   -1 : there is no selected DB
'   -2 : fail to execute
'    0 : execute success
Public Function ExecuteSQL(ByVal sqlString As String) As Integer
    If ffDB = 0 Then
        ExecuteSQL = -1
    Else
        'check sql and call function for it
        If Format(Mid(sqlString, 1, Len("CREATE TABLE ")), ">") = "CREATE TABLE " Then
            ExecuteSQL = CreateTable(sqlString)
        ElseIf Format(Mid(sqlString, 1, Len("INSERT INTO ")), ">") = "INSERT INTO " Then
            ExecuteSQL = InsertRecord(sqlString)
        End If
    End If
End Function


Private Sub Class_Terminate()
    Erase mTblParams
    Erase lngColl2
End Sub
