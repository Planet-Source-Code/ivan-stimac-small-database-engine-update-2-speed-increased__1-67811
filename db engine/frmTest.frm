VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Small Database Engine"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelp 
      Caption         =   "?"
      Height          =   255
      Left            =   9120
      TabIndex        =   13
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton Command7 
      Caption         =   "If you want vote click me"
      Height          =   855
      Left            =   8400
      Picture         =   "frmTest.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox txtDelete 
      Height          =   285
      Left            =   6420
      TabIndex        =   11
      Text            =   "(ID) VALUES (<'400')"
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Open DB"
      Height          =   735
      Left            =   180
      Picture         =   "frmTest.frx":1982
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3300
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Compact"
      Height          =   735
      Left            =   9540
      Picture         =   "frmTest.frx":3304
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton DeleteRecords 
      Caption         =   "Delete records"
      Height          =   735
      Left            =   6420
      Picture         =   "frmTest.frx":3942
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create db"
      Height          =   735
      Left            =   180
      Picture         =   "frmTest.frx":4002
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add to list"
      Height          =   735
      Left            =   4860
      Picture         =   "frmTest.frx":4642
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   180
      TabIndex        =   3
      Top             =   1680
      Width           =   10395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read fields"
      Height          =   735
      Left            =   3300
      Picture         =   "frmTest.frx":497B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1740
      TabIndex        =   1
      Text            =   "10000"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add fields"
      Height          =   735
      Left            =   1740
      Picture         =   "frmTest.frx":4FC0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "####"
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   5160
      Width           =   6675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----- ABOUT --------------------------------------------------------------------------------
'
'   Author : Ivan Stimac, CROATIA
'   E-mail : ivan.stimac@po.htnet.hr or flashboy01@gmail.com
'
'
'   Short program description:
'   this is sample of fast db engine that support few sql queryes, muliple tables, username and
'   password for acces to db, add/delete data...
'
'
'----- COPYRIGHT ----------------------------------------------------------------------------
'
'   This program is free for any use. Just please contact me if you find some bugs or
'   with any idea how to improve it.
'
'
'-----SQL suported:--------------------------------------------------------------------------
'
'       CREATE TABLE tbl_name (colName1, colName2...)
'       INSERT INTO tbl_name (colName1, colName2...) VALUES ('val1','val2'...)
'           also: INSERT INTO tbl_name (colName2, colName1...) VALUES ('val2','val1'...)
'                 -> program will alone make correct array
'                   ---------------------------------
'                   NOTE: don't use ' in values data
'                   ---------------------------------
'
'
'-----Read records:--------------------------------------------------------------------------
'
'       you can use class clsRecSet to store data:
'          like:  Set mRecSet = mDB.ReadRecords(tbl_name)
'
'
'-----Delete records:------------------------------------------------------------------------
'
'       there is function DeleteRecords (tblName as String, whereWhat as String)
'
'           tblName   : table name whicj contains data to delete
'
'           whereWhat : delete only data that match condition
'                       this is not like Delete SQL query, using:
'                       (ColName1, colName2...) VALUES (operator1'val1', operator2'val2',...)
'                       operators: =, <, >, <>
'                       how it works:(if we use operator1 is = and operator2 is <>):
'                                   if colName1 = val1 AND colName2 <> val2 AND ...
'                           so you can see there is only AND operator between checks
'                           OR operator and other operators os no avaible
'
'           NOTE: erased data is still in file, it's only disabled for read
'                 for full clear you must do CompactDatabase
'
'
'---------------------------------------------------------------------------------------------
'-----IMPORTANT TO KNOW-----------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
'
'   To speed up I'm disabled update table after writing to it.
'   So You MUST DO IT ALONE WHEN YOU WRITE ALL DATA FOR THIS TABLE CALLING FUNCTION UpdateAfterWrite
'   You can write to another table before updating previous, you only need to update before
'   you close db (if you don't data will be in file but you can't read them (like deleted data))
'   This update can decrese write time few times.
'
'
'----Modules-----------------------------------------------------------------------------------
'
'   if you want you can code from all modules add to one
'   module. I'm created few modules becouse it's more easy
'   for coding
'
'
'----What I learned by creating this project----------------------------------------------------
'
'   * more about binary file access
'   * always avoid collection class (I have know it's some slower from arrays, but when I created test
'       I can't belive how much: for 10 000 records with collection takes 17 secs, without
'       collection about 0,3 sec)
'       Writing to collection is not problem, but reading is too slow
'   * when searching string don't never go throught it and checking char by char, using InStr function
'       (where you can) you got much faster code (see ReadRecords function in clsDatabase)
'
'   * hope that helps someone
'
'   * NOTE: you can find much commented code - commented code is used before, it's much slower
'
'
'----The end ----------------------------------------------------------------------------------
'
'   If you find some bugs please contact me on ivan.stimac@po.htnet.hr
'   or flashboy01@gmail.com or leave feedback on PSC.
'
'   Thanks for downloading!
'
'----------------------------------------------------------------------------------------------
'   a n d    :-(   sorry for my english. ;-)

'
Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_NORMAL = 1
'

Private mRec As New clsRecSet
Private mDB As New clsDatabase

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub cmdHelp_Click()
    MsgBox "This is something like DELTE QUERY but it's not same." & vbCrLf & vbCrLf & _
           "(col_name1, col_name2,...) VALUES (operator 'value1', operator 'value2',...)" & vbCrLf & _
           "Operator can be =, <, >, <>" & vbCrLf & vbCrLf & _
           "Example: (ID) VALUES (<'400') will delete all fields where ID is < 400" & vbCrLf & _
           "Example: (userName, password) VALUES (='ivan',='myPass') will delete all fields where userName is ivan and password is myPass" & vbCrLf & vbCrLf & _
           "NOTE: table names and column names are case sensitive"
           
End Sub

'add data to each table
Private Sub Command1_Click()
    Dim Tick As Long
    'open db
    mDB.OpenDB App.Path & "\testDB.txt", "userName=ivan;password=8547hhh"
    Tick = GetTickCount
    'save data for first database
    For i = 1 To Val(Me.Text1.Text) / 2
        mDB.ExecuteSQL ("INSERT INTO MyTable1 (ID, Password, userName) VALUES ('" & i & "','8547hhh', 'ivan')")
    Next i
    'mDB.UpdateAfterWrite
    'save data for second database
    For i = 1 To Val(Me.Text1.Text) / 2
        mDB.ExecuteSQL ("INSERT INTO MyTable2 (ID, carType, engine) VALUES ('" & i & "','Porsche 911 Turbo', 'this is some description of cars engine and gearbox. There is data about engine horse power, fuel...')")
    Next i
    'update tables after write (must be done)
    mDB.UpdateAfterWrite
    'get used time for execution
    Tick = GetTickCount - Tick
    MsgBox "Elapsed " & Format$(Tick / 1000, "0.000") & " second", 32
    MsgBox "Add " & Val(Me.Text1.Text) / 2 & "to each table!"
    mDB.CloseDB
End Sub

'read data and set them to recordset
Private Sub Command2_Click()
    Dim Tick As Long
    If mDB.OpenDB(App.Path & "\testDB.txt", "userName=ivan;password=8547hhh") = 0 Then
        Tick = GetTickCount
        Set mRec = mDB.ReadRecords(Me.Combo1.List(Me.Combo1.ListIndex))
        Tick = GetTickCount - Tick
        MsgBox "Elapsed " & Format$(Tick / 1000, "0.000") & " second", 32
    Else
        MsgBox "Can't open DB!"
    End If
    mDB.CloseDB
End Sub

'show data in list box
Private Sub Command3_Click()
    Dim i As Long, j As Long
    Dim tmpStr As String
    mRec.FirstRow
    Me.List1.Clear
    For i = 0 To mRec.Rows - 1
        tmpStr = ""
        For j = 0 To mRec.Columns - 1
            tmpStr = tmpStr & mRec.DataRow(j) & "      "
        Next j
        Me.List1.AddItem tmpStr
        mRec.NextRow
    Next i
    Me.Label1.Caption = "List contains " & Me.List1.ListCount & " records"
End Sub


'create db file and insert tables
Private Sub Command4_Click()
    mDB.CreateDB App.Path & "\testDB.txt", "userName=ivan;password=8547hhh"
    mDB.OpenDB App.Path & "\testDB.txt", "userName=ivan;password=8547hhh"
    mDB.ExecuteSQL ("CREATE TABLE MyTable1 (ID, userName, Password)")
    mDB.ExecuteSQL ("CREATE TABLE MyTable2 (ID, carType, engine)")
    'IMPORTANT:
    'close db for refresh (should do after creating tables)
    mDB.CloseDB
    MsgBox "Created!" & vbCrLf & _
           "Now, click on Open DB to load tabled in combo box, then Add fields button, after that read fields and add to list" & vbCrLf & _
           "Also you can change MyTable1 to MyTable2 and click Read Fields and Add to list"
End Sub
'delete unused (delited) data from file
Private Sub Command5_Click()
    Dim mDbCmp As New clsDBAdvanced
    Dim Tick As Long
    Tick = GetTickCount
    mDbCmp.CompactDatabase App.Path & "\testDB.txt", App.Path & "\testDB1.txt", "userName=ivan;password=8547hhh"
    mDB.CloseDB
    DoEvents
    Kill App.Path & "\testDB.txt"
    FileCopy App.Path & "\testDB1.txt", App.Path & "\testDB.txt"
    Kill App.Path & "\testDB1.txt"
    Tick = GetTickCount - Tick
    MsgBox "Elapsed " & Format$(Tick / 1000, "0.000") & " second", 32
End Sub

'open db file and all all tables to combo box
Private Sub Command6_Click()
    If mDB.OpenDB(App.Path & "\testDB.txt", "userName=ivan;password=8547hhh") = 0 Then
        Me.Combo1.Clear
        For i = 0 To mDB.getTablesCount - 1
            Me.Combo1.AddItem mDB.getTableName(i)
        Next i
        Me.Combo1.ListIndex = 0
    Else
        MsgBox "Can't open db (invalid file, username or password!)"
    End If
    mDB.CloseDB
End Sub
'open psc
Private Sub Command7_Click()
    ShellExecute hwnd, "open", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=67811&lngWId=1", vbNullString, vbNullString, conSwNormal
End Sub
'delete data
Private Sub DeleteRecords_Click()
    Dim Tick As Long
    Tick = GetTickCount
    mDB.OpenDB App.Path & "\testDB.txt", "userName=ivan;password=8547hhh"
    mDB.DeleteRecords Me.Combo1.List(Me.Combo1.ListIndex), txtDelete.Text
    mDB.CloseDB
    '
    Tick = GetTickCount - Tick
    MsgBox "Elapsed " & Format$(Tick / 1000, "0.000") & " second", 32
End Sub

Private Sub Form_Load()
    MsgBox "IMPORTANT TO KNOW" & vbCrLf & vbCrLf & _
   "To speed up I'm disabled update table after writing to it." & vbCrLf & _
   "So You MUST DO IT ALONE WHEN YOU WRITE ALL DATA FOR THIS TABLE CALLING FUNCTION UpdateAfterWrite" & vbCrLf & _
   "You can write to another table before updating previous, you only need to update before" & vbCrLf & _
   "you close db (if you don't data will be in file but you can't read them (like deleted data))" & vbCrLf & _
   "This update can decrese write time few times.", vbExclamation, "PLEASE READ"
End Sub

Private Sub Label4_Click()

End Sub
