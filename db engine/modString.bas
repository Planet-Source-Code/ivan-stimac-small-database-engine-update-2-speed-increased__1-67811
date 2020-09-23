Attribute VB_Name = "modString"
Option Explicit
'get username and password from strDbAccess string (username=sdgasg;password=sdgsd)
Public Sub parseDbAccess(ByVal strDbAccess As String, ByRef strUserName As String, ByRef strPass As String)
    Dim i As Integer
    Dim tmpStr As String, strChr As String ', strPass As String, strUserName As String
    Dim startWrite As Boolean
    tmpStr = ""
    For i = 1 To Len(strDbAccess)
        strChr = Mid(strDbAccess, i, 1)
        If strChr = "=" Then
            startWrite = True
        ElseIf strChr = ";" Then
            strUserName = tmpStr
            tmpStr = ""
            startWrite = False
        ElseIf startWrite = True Then
            tmpStr = tmpStr & strChr
        End If
    Next i
    strPass = tmpStr
    'MsgBox strPass
End Sub

'parse string with tables params and store data in tblParams array
Public Sub parseTables(ByVal strToParse As String, ByRef storeParams() As tblParams)
    Dim i As Integer, colCnt As Integer, readLevel As Integer, readSubLevel As Integer
    Dim tblsNum As Integer, currTbl As Integer, z As Integer
    Dim strChr As String, tmpStr As String
    Dim tmpName As String, tmpFirst As Long, tmpLast As Long, tmpCols() As String
    readLevel = 0
    readSubLevel = 0
    tblsNum = 0
    For i = 1 To Len(strToParse)
        If Mid(strToParse, i, 1) = "]" Then tblsNum = tblsNum + 1
    Next i
    
    'MsgBox tblsNum
    
    If tblsNum = 0 Then Exit Sub
    ReDim storeParams(tblsNum)
    
    currTbl = 0
    For i = 1 To Len(strToParse)
        'MsgBox "POC:" & tmpStr
        strChr = Mid(strToParse, i, 1)
        If strChr = "[" Then
            tmpName = tmpStr
            tmpStr = ""
        ElseIf strChr = "(" Then
            readLevel = readLevel + 1
            readSubLevel = 0
            tmpStr = ""
        ElseIf strChr = ")" And readLevel = 1 Then
            colCnt = tmpStr
            ReDim tmpCols(colCnt)
            tmpStr = ""
        ElseIf strChr = ")" And readLevel = 2 Then
            tmpCols(readSubLevel) = tmpStr
            tmpStr = ""
        ElseIf strChr = ")" And readLevel = 3 Then
            tmpLast = tmpStr
            tmpStr = ""
        ElseIf strChr = "," Then
            If readLevel = 2 Then
                tmpCols(readSubLevel) = tmpStr
                tmpStr = ""
            ElseIf readLevel = 3 Then
                'If readSubLevel = 0 Then
                
                    tmpFirst = tmpStr
                    'MsgBox tmpFirst
               ' Else
                   ' tmpLast = tmpStr
                    'MsgBox tmpLast
               ' End If
                tmpStr = ""
            End If
            readSubLevel = readSubLevel + 1
        ElseIf strChr = "]" Then
            readLevel = 0
            readSubLevel = 0
            tmpStr = ""
        '    tblName As String
        '    tblColCnt As Integer
        '    tblCols() As String
        '    tblFirstData As Long
        '    tblLastData As Long
            storeParams(currTbl).tblName = tmpName
            storeParams(currTbl).tblColCnt = colCnt
            storeParams(currTbl).tblFirstData = tmpFirst
            storeParams(currTbl).tblLastData = tmpLast
            ReDim storeParams(currTbl).tblCols(colCnt)
            'save column names
            For z = 1 To colCnt
                storeParams(currTbl).tblCols(z - 1) = tmpCols(z - 1)
            Next z
            currTbl = currTbl + 1
        Else
            tmpStr = tmpStr & strChr
            'MsgBox tmpStr
        End If
    Next i
    
    Erase tmpCols
End Sub

Public Sub setLen(ByRef myStr As String, ByVal intLen As Integer)
    If Len(myStr) < intLen Then myStr = myStr & Space(intLen - Len(myStr))
End Sub

'create string from tblParams for write to db
Public Function CreateStrDbTablesDef(ByRef tbls() As tblParams) As String
    On Error Resume Next
    Dim i As Integer, z As Integer
    'check for error (if there is no tables)
    i = UBound(tbls)
    If Err.Number <> 0 Then Exit Function
    '
    CreateStrDbTablesDef = ""
    For i = 0 To UBound(tbls) - 1
        CreateStrDbTablesDef = CreateStrDbTablesDef & Trim(tbls(i).tblName) & "[(" & tbls(i).tblColCnt & ")("
        For z = 0 To tbls(i).tblColCnt - 1
            CreateStrDbTablesDef = CreateStrDbTablesDef & tbls(i).tblCols(z)
            If z < tbls(i).tblColCnt - 1 Then
                CreateStrDbTablesDef = CreateStrDbTablesDef & ","
            End If
        Next z
        CreateStrDbTablesDef = CreateStrDbTablesDef & "):(" & tbls(i).tblFirstData & "," & tbls(i).tblLastData & ")]"
    Next i
End Function

'
Public Function InStrCharCount(ByVal mString As String, ByVal strChar As String) As Integer
    Dim mPos As Long
    mPos = 1
    InStrCharCount = 0
    Do While mPos > 0
        mPos = InStr(mPos, mString, strChar, vbBinaryCompare)
        If mPos > 0 Then
            InStrCharCount = InStrCharCount + 1
            mPos = mPos + 1
        End If
    Loop
End Function





