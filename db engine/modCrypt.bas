Attribute VB_Name = "modCrypt"
'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
'
'       MODULE      : modCrypt
'       VERSION     : 1.0.0
'       DESCRIPTION : encript/decript strings. This is not some hard encription.
'                     It only increase or decrease char asc for len_of_string/divLen
'                     NOTE: len_of_string/divLen should not be bigger than 255 or
'                           you will got error. So if you want use this module for
'                           your projects you must know len of string to encript
'                           and if it's bigger than 255 you must increase divLen
'                           enought to get len_of_string/divLen less than 255
'
'--------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------

Public Function Crypt(ByVal strToCrypt As String, Optional divLen As Integer = 1) As String
    Dim i As Integer, tmpASC As Integer
    Dim tmpStr As String, tmpChr As String
    tmpStr = ""
    For i = 1 To Len(strToCrypt)
        tmpChr = Mid(strToCrypt, i, 1)
        If Asc(tmpChr) <> 0 Then
            tmpASC = Asc(tmpChr) - Len(strToCrypt) / divLen
            If tmpASC < 1 Then tmpASC = 255 + tmpASC
            tmpStr = tmpStr & Chr(tmpASC)
        Else
            'tmpStr = tmpStr & tmpChr
            Exit Function
        End If
    Next i
    Crypt = tmpStr
End Function

Public Function Decrypt(ByVal strToDecrypt As String, Optional divLen As Integer = 1) As String
    Dim i As Integer, tmpASC As Integer
    Dim tmpStr As String, tmpChr As String
    tmpStr = ""
    For i = 1 To Len(strToDecrypt)
        tmpChr = Mid(strToDecrypt, i, 1)
        If Asc(tmpChr) <> 0 Then
            tmpASC = Asc(tmpChr) + Len(strToDecrypt) / divLen
            If tmpASC > 255 Then tmpASC = tmpASC - 255
            tmpStr = tmpStr & Chr(tmpASC)
        Else
            'tmpStr = tmpStr & tmpChr
            Exit Function
        End If
    Next i
    Decrypt = tmpStr
End Function
