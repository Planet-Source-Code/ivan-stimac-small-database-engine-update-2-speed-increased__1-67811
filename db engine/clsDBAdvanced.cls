VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim mDB As New clsDatabase
Dim mDBNew As New clsDatabase
Dim mRecSet As New clsRecSet
'
' function messages:
'   -1 : invalid db file
'   -2 : fail
'    0 : success
Public Function CompactDatabase(ByVal dbFile As String, ByVal saveNewFile As String, Optional dbAccess As String = "") As Integer
    On Error Resume Next
    Dim i As Long, j As Long, k As Long
    Dim strQuery As String, tmp1Str As String
    
    'create new database
    mDBNew.CreateDB saveNewFile, dbAccess
    If mDB.OpenDB(dbFile, dbAccess) < 0 Then
        CompactDatabase = -1
        mDBNew.CloseDB
        mDB.CloseDB
        Exit Function
    End If
    
    mDBNew.OpenDB saveNewFile, dbAccess
    'create tables
    For i = 0 To mDB.getTablesCount - 1
        strQuery = "CREATE TABLE " & mDB.getTableName(i) & "("
        For j = 0 To mDB.getColCount(i) - 1
            strQuery = strQuery & mDB.getColName(i, j)
            If j < mDB.getColCount(i) - 1 Then strQuery = strQuery & ","
        Next j
        strQuery = strQuery & ")"
        mDBNew.ExecuteSQL strQuery
        'msgBox strQuery
    Next i
    DoEvents
    mDBNew.CloseDB
    mDBNew.OpenDB saveNewFile, dbAccess
    '
    'fil data
    For i = 0 To mDB.getTablesCount - 1
        Set mRecSet = mDB.ReadRecords(mDB.getTableName(i))
        tmp1Str = "INSERT INTO " & mDB.getTableName(i) & "("
        For k = 0 To mDB.getColCount(i) - 1
            tmp1Str = tmp1Str & mDB.getColName(i, k)
            If k < mDB.getColCount(i) - 1 Then tmp1Str = tmp1Str & ","
        Next k
        tmp1Str = tmp1Str & ") "
        For j = 0 To mRecSet.Rows - 1
            strQuery = tmp1Str & "VALUES ("
            For k = 0 To mRecSet.Columns - 1
                strQuery = strQuery & "'" & mRecSet.Data(j, k) & "'"
                If k < mRecSet.Columns - 1 Then strQuery = strQuery & ","
            Next k
            strQuery = strQuery & ")"
            'MsgBox strQuery
            mDBNew.ExecuteSQL strQuery
        Next j
    Next i
    mDBNew.CloseDB
    mDB.CloseDB
End Function
