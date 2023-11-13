VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cXpCapitalDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Const XPDB = "XpCapital.mdb"

Private t As Single

Private dbpath As String
Private cn As ADODB.Connection
Private rs As ADODB.Recordset
Private sql As String
Private lCount As Long
Public o As Scripting.Dictionary    '// structure of scanner tables, fields, and rs

Private Sub Class_Initialize()
    dbpath = Environ("appdata") & "\xpCapital\" & XPDB
    If Dir(dbpath) = vbNullString Then
        CreateDb
    Else
        CompactDb
        ConnectDb
    End If
    checkDb
End Sub
Private Sub Class_Terminate()
    If Not cn Is Nothing Then If cn.State = adStateOpen Then cn.Close
    Set cn = Nothing
End Sub

Private Sub ConnectDb()
    If cn Is Nothing Then Set cn = New ADODB.Connection
    With cn
        If .State <> adStateOpen Then
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .Open dbpath
            If .State <> adStateOpen Then MsgBox "Could NOT open xpdb - db state is " & .State
        End If
    End With
End Sub
Private Sub checkDb()       '// create tables if not found
    Dim k
    Dim i As Integer
    Dim tb As String        '// load xpDb.o with dictionary
    
    Set rs = cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
''    Stop
    rs.Filter = "TABLE_NAME Like '__*'"     '// filter on underscore tables, two underscores to escape
    Set o = New Dictionary
    Do Until rs.EOF         '// load table fields and table into dictionary xpDb.o
        tb = rs.Fields(2).Value
        Set o(right(tb, Len(tb) - 1)) = New Dictionary  '// each table is a dictionary
        rs.MoveNext
    Loop
    rs.Close
    
    For Each k In o     '// build a json structure {tablename: fieldnames: fieldIndex, "rs" = recordset}
        Set rs = New Recordset
        With rs
            .CursorLocation = adUseClient
            .Open "_" & k, cn, adOpenStatic, adLockBatchOptimistic, adCmdTableDirect
            Set .ActiveConnection = Nothing
            For i = 0 To .Fields.Count - 1
                o(k)(.Fields(i).name) = i
            Next i
            Set o(k)("rs") = rs
        End With
    Next k
End Sub
Private Sub CompactDb()     '// keeps at most 6 backup copies of db, compacts every day or so
    Dim bak As String
    bak = Replace(dbpath, ".mdb", "-" & Weekday(Now()) & ".mdb")    '// LiDb-6.mdb
    If Dir(bak) <> vbNullString Then
        If Now() - FileDateTime(bak) < 1 Then Exit Sub
        Kill bak
    End If
    Name dbpath As bak
    CreateObject("DAO.DBEngine.120").CompactDatabase bak, dbpath
    Debug.Print FileLen(bak), bak
    Debug.Print FileLen(dbpath), dbpath
End Sub

'// Data Definition Language DDL Section
'// SQL DDL Procedures
    
Private Sub CreateDb()
    mkDatabase    '// INT is 4 Bytes ~2B max for primary keys - no BIGINT avail
    mkContract
    mkAccountValue
    mkPorfolio
End Sub
Private Sub mkDatabase()
    If MsgBox(XPDB & " not found in: " & dbpath & vbNewLine & "Create it ?", vbYesNo) = vbYes Then
        With CreateObject("ADOX.Catalog")
            .Create ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath & ";")
            Set cn = .ActiveConnection
        End With
    Else
        Stop
    End If
End Sub

Public Sub uploadScan()
    Dim k
    
    If o Is Nothing Then Exit Sub
    For Each k In o
        With o(k)("rs")
            Set .ActiveConnection = cn
            .UpdateBatch
        End With
    Next k
End Sub
Public Sub syncUpload()
    t = timer
''    insIndustry
''    insSchool
''    insCompany
''
''    deleteNameMatch
''    insMiniProfile
''    copyProfileMiniProfile
''    insProfile
''    insPosition
''    insEducation
''
''    insMemberBadges
''    insProfileContactInfo
''    insMe
''    insMiniGroup
''    insSearchUrl
''    insSearchHitV2

End Sub

'// Tws tables

Private Sub mkAccountValue()    '// root table TWS
    sql = ""
    sql = sql & "CREATE TABLE AccountValues ("
    sql = sql & "   indexXp CHAR(255) NOT NULL"  '// F749544.TotalCashValue.USD
    sql = sql & "   , accountName CHAR(255) NOT NULL"  '// F749544A
    sql = sql & "   , [key] VARCHAR(63) NOT NULL"     '// key is reserved word!!!
    sql = sql & "   , [currency] VARCHAR(15)"
    sql = sql & "   , [value] VARCHAR(255)"
    sql = sql & "   , [numValue] FLOAT"
    sql = sql & "   , [timeStamp] VARCHAR(15) "
    sql = sql & "   );"
    cn.Execute sql
    cn.Execute "SELECT * INTO [_AccountValues] FROM AccountValues WHERE FALSE"
End Sub
Private Sub mkPorfolio()
    sql = ""
    sql = sql & "CREATE TABLE Portfolios ("
    sql = sql & "   indexXp CHAR(255) NOT NULL"  '// F749544.BND
    sql = sql & "   , accountName VARCHAR(255) NOT NULL "
    sql = sql & "   , ContractId INT NOT NULL "
    sql = sql & "   , [Position] FLOAT "          '// Position is reserved word !!!
    sql = sql & "   , marketPrice FLOAT "
    sql = sql & "   , marketValue FLOAT "
    sql = sql & "   , averageCost FLOAT "
    sql = sql & "   , unrealizedPNL FLOAT "
    sql = sql & "   , realizedPNL FLOAT "
    sql = sql & "   , [timeStamp] VARCHAR(15) "
    sql = sql & "   );"
    cn.Execute sql
    cn.Execute "SELECT * INTO [_Portfolios] FROM Portfolios WHERE FALSE"
End Sub
Private Sub mkContract()
    sql = ""
    sql = sql & "CREATE TABLE Contracts ("
    sql = sql & "   ConId INT NOT NULL CONSTRAINT PK_Contract PRIMARY KEY"
    sql = sql & "   , [Symbol] VARCHAR(255) "
    sql = sql & "   , SecType VARCHAR(255) "
    sql = sql & "   , LastTradeDateOrContractMonth  VARCHAR(15) "
    sql = sql & "   , Strike FLOAT "
    sql = sql & "   , [Right] VARCHAR(255) "
    sql = sql & "   , Multiplier VARCHAR(255) "
    sql = sql & "   , Exchange VARCHAR(255) "
    sql = sql & "   , [Currency] VARCHAR(255) "
    sql = sql & "   , LocalSymbol VARCHAR(255) "
    sql = sql & "   , PrimaryExch VARCHAR(255) "
    sql = sql & "   , TradingClass VARCHAR(255) "
    sql = sql & "   , IncludeExpired BIT"
    sql = sql & "   , SecIdType VARCHAR(255)"
    sql = sql & "   , SecId VARCHAR(255)"
    sql = sql & "   , ComboLegDescription VARCHAR(255)"
    sql = sql & "   , [timeStamp] VARCHAR(15) "
    sql = sql & "   );"
    cn.Execute sql
    cn.Execute "SELECT * INTO [_Contracts] FROM Contracts WHERE FALSE"
End Sub

Private Sub mkComboLeg()
    sql = ""
    sql = sql & "CREATE TABLE ComboLeg ("
    sql = sql & "   ComboLegId INT CONSTRAINT PK_ComboLeg PRIMARY KEY"
    sql = sql & "   , birthDateOn INT"
    sql = sql & "   , address NVARCHAR(255) WITH COMPRESSION"
    sql = sql & "   , emailAddress VARCHAR(255) NOT NULL"
    sql = sql & "   , phoneNumbers VARCHAR(255)"
    sql = sql & "   , connectedAt DECIMAL(13,0)"
    sql = sql & "   , websites NVARCHAR(255) WITH COMPRESSION"
    sql = sql & "   , epoch DECIMAL(13,0) NOT NULL"
    sql = sql & "   , CONSTRAINT FK_ProfileContactInfo_MiniProfile FOREIGN KEY (entityUrn)"
    sql = sql & "       REFERENCES MiniProfile (entityUrn) ON DELETE CASCADE"
    sql = sql & "   );"
    cn.Execute sql
    cn.Execute "SELECT * INTO [_ComboLeg] FROM ComboLeg WHERE FALSE"
End Sub

