Attribute VB_Name = "modSQL"
Option Explicit
'parse CREATE TABLE SQL query
'   CREATE TABLE tblName (colName1, colName2...) => TO :
'   tblName[(colNums)(colName1, colName2...):(firstItem_pos,lastItem_pos)]
Public Function getTblParamsFromSQL(ByVal strSQL As String) As String
    Dim i As Integer, z As Integer
    Dim strChr As String, tmpStr As String
    Dim tmpTblDef As tblParams
    tmpTblDef.tblColCnt = 0
    If Format(Mid(strSQL, 1, Len("CREATE TABLE ")), ">") = "CREATE TABLE " Then
        'first we need to find coll count
        For i = Len("CREATE TABLE ") To Len(strSQL)
            If Mid(strSQL, i, 1) = "," Then tmpTblDef.tblColCnt = tmpTblDef.tblColCnt + 1
        Next i
        tmpTblDef.tblColCnt = tmpTblDef.tblColCnt + 1
        '
        z = 0
        'and then reserve tblCols(col_count) for column names
        ReDim tmpTblDef.tblCols(tmpTblDef.tblColCnt - 1)
        'fisrs and last data is 0 because there is no data yet
        tmpTblDef.tblFirstData = 0
        tmpTblDef.tblLastData = 0
        'read table name and column names
        For i = Len("CREATE TABLE ") To Len(strSQL)
            strChr = Mid(strSQL, i, 1)
            If strChr = "(" Then
                tmpTblDef.tblName = Trim(tmpStr)
                tmpStr = ""
            ElseIf strChr = "," Or strChr = ")" Then
                tmpTblDef.tblCols(z) = Trim(tmpStr)
                tmpStr = ""
                z = z + 1
            Else
                tmpStr = tmpStr & strChr
            End If
        Next i
        'set return value for that will be encriped and writen to file
        getTblParamsFromSQL = tmpTblDef.tblName & "[(" & tmpTblDef.tblColCnt & ")("
        For i = 0 To tmpTblDef.tblColCnt - 1
            getTblParamsFromSQL = getTblParamsFromSQL & tmpTblDef.tblCols(i)
            If i < tmpTblDef.tblColCnt - 1 Then getTblParamsFromSQL = getTblParamsFromSQL & ","
        Next i
        getTblParamsFromSQL = getTblParamsFromSQL & "):(0,0)]"
    'if we have invalid query
    Else
        getTblParamsFromSQL = "INVALID SQL QUERY"
    End If
End Function
'parse INSERT INTO SQL query
'   INSERT INTO tblName (colName1, colName2...) VALUES ('value1','value2'...) => TO :
'   'value1'|'value2'|...
Public Function parseInsetIntoSql(ByVal strSQL As String, ByRef tbls() As tblParams, ByRef mTblInd As Integer) As String
    On Error Resume Next
    Dim i As Long, z As Long, k As Long, dataCnt As Integer, len1 As Integer, len2 As Integer
    Dim strChr As String, tmpStr As String
    Dim strFields() As String, strData() As String, strDataSorted() As String
    Dim mTbl As String, mTblIndex As Integer
    
    'get table name
    mTbl = Mid(strSQL, Len("INSERT INTO "), InStr(1, strSQL, "(") - Len("INSERT INTO "))
    mTbl = Trim(mTbl)
    
    'get table name (slow)
'    For i = Len("INSERT INTO ") To Len(strSQL)
'        strChr = Mid(strSQL, i, 1)
'        If strChr = "(" Then Exit For
'        mTbl = mTbl & strChr
'    Next i
'    mTbl = Trim(mTbl)
    '
    'find table index in tblParams
    For i = 0 To UBound(tbls)
        'MsgBox tbls(i).tblName & vbCrLf & mTbl
        If Trim(tbls(i).tblName) = mTbl Then
            mTblIndex = i
            mTblInd = i
            Exit For
        ElseIf i = UBound(tbls) Then
            parseInsetIntoSql = "INVALID TABLE NAME!"
            mTblInd = -1
            Exit Function
        End If
    Next i
    
    'count fields
    dataCnt = InStrCharCount(strSQL, ",")
    dataCnt = dataCnt + 1
    'count fields (slow)
'    For i = 1 To Len(strSQL)
'        'strChr = Mid(strSQL, i, 1)
'        If Mid(strSQL, i, 1) = "," Then
'            dataCnt = dataCnt + 1
'        ElseIf Mid(strSQL, i, 1) = ")" Then
'            dataCnt = dataCnt + 1
'            Exit For
'        End If
'    Next i
    '
    ReDim strFields(dataCnt)
    ReDim strData(dataCnt)

    'read fields
    z = 0
    i = InStr(1, strSQL, "(")
    k = InStr(1, strSQL, ")")
    len1 = 0
    Do While len1 < k
        len2 = InStr(len1 + 1, strSQL, ",")
        If len2 > k Then len2 = k
        If len1 = 0 Then
            strFields(z) = Trim(Mid(strSQL, i + 1, len2 - i - 1))
            '
        Else
            strFields(z) = Trim(Mid(strSQL, len1 + 1, len2 - len1 - 1))
        End If
       ' MsgBox strFields(z) & vbCrLf & k & "----" & i
        len1 = len2
        z = z + 1
    Loop

    'read fields (slow)
'    For i = 1 To Len(strSQL)
'        strChr = Mid(strSQL, i, 1)
'        If strChr = "(" Then
'            'startWrite = True
'            tmpStr = ""
'        ElseIf strChr = "," Then
'            strFields(z) = Trim(tmpStr)
'            tmpStr = ""
'            z = z + 1
'        ElseIf strChr = ")" Then
'            strFields(z) = Trim(tmpStr)
'            Exit For
'        Else
'            tmpStr = tmpStr & strChr
'        End If
'    Next i

    'find data values
    len1 = InStr(10, strSQL, "VALUES", vbTextCompare)
    z = 0
    i = 0
    'k = len1
    Do While len1 > 0
        len1 = InStr(len1, strSQL, "'") 'mColl.Item(i), "'")
        If len1 > 0 Then
            i = i + 1
            If i = 2 Then
                strData(z) = Mid(strSQL, len2, len1 - len2)
                z = z + 1
                i = 0
            Else
                len2 = len1 + 1
            End If
            len1 = len1 + 1
        End If
    Loop
    
'    'find data values (slow)
'    For i = len1 To Len(strSQL)
'        strChr = Mid(strSQL, i, 1)
'        If strChr = "(" Then
'            tmpStr = ""
'        ElseIf strChr = "'" Then
'            k = k + 1
'            If k = 2 Then
'                strData(z) = Trim(tmpStr)
'                z = z + 1
'                k = 0
'            End If
'            tmpStr = ""
'        Else
'            tmpStr = tmpStr & strChr
'        End If
'    Next i
    
    'now we must sort data at this order as tbl definition columns order
    ReDim strDataSorted(tbls(mTblIndex).tblColCnt)
    For i = 0 To tbls(mTblIndex).tblColCnt - 1
        For z = 0 To UBound(strFields)
            If tbls(mTblIndex).tblCols(i) = strFields(z) Then strDataSorted(i) = strData(z)
        Next z
    Next i
    
    'set return value
    parseInsetIntoSql = ""
    For i = 0 To UBound(strDataSorted) - 1
        If i > 0 Then parseInsetIntoSql = parseInsetIntoSql & "|"
        parseInsetIntoSql = parseInsetIntoSql & "'" & strDataSorted(i) & "'"
    Next i
    
    'at the end free memory
    Erase strFields
    Erase strData
End Function


