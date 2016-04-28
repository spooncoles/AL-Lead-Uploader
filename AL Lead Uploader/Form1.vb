Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop

Public Class Form1

    'TEST COMMENT FOR GITHUB

    'Dim conn.conSQL As MySqlConnection = New MySqlConnection("Data Source=196.44.211.224;Database=AssetLife;User=structest_mysql;Password=Z35tL1f3!PWD")
    'Dim conn.comSQL As MySqlCommand = New MySqlCommand(Nothing, conn.conSQL)
    'Dim conn.daSQL As MySqlDataAdapter = New MySqlDataAdapter(Nothing, conn.conSQL)
    'Dim conn.ds As DataSet = New DataSet()

    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlWS As Excel.Worksheet

#Region "Dictionary Code"
    '  Console.WriteLine("Start: " & Now.TimeOfDay.ToString)

    '  conn.daSQL.SelectCommand = conn.comSQL
    '  conn.comSQL.CommandText = "SELECT DISTINCT(VINNUMBER) as VIN, SOURCE from asset"
    '  conn.daSQL.Fill(conn.ds, "vinTable")

    '  Console.WriteLine("datatable Filled: " & Now.TimeOfDay.ToString)


    '  Dim dict As Dictionary(Of String, List(Of String)) = (From row As DataRow In conn.ds.Tables("vinTable").Rows.Cast(Of DataRow)().AsEnumerable _
    'Group row By type = row.Field(Of String)("VIN") Into Group _
    'Select New With { _
    '.k = type, _
    '.v = (From r As DataRow In Group.Cast(Of DataRow)().AsEnumerable _
    '      Select CStr(r.Field(Of String)("VIN"))).ToList} _
    '      ).ToDictionary(Function(kvp) kvp.k, Function(kvp) kvp.v)

    '  Console.WriteLine("Dict Filled: " & Now.TimeOfDay.ToString)

    '  If dict.ContainsKey("WAUZZZ8X4GB025272") Then
    '      'Console.WriteLine(dict.Item("WAUZZZ8X4GB025272"))
    '      Console.WriteLine("Found: " & Now.TimeOfDay.ToString)
    '  End If
#End Region

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim conn As New clConn

        Console.WriteLine("Start: " & Now.TimeOfDay.ToString)
        Dim zl_prod_id As Integer = 0
        Dim insertColumns As String = ""
        Dim insertValues As String = ""

        Dim primaryKeyName As String = ""
        Dim primaryKeyValue As String = ""

        Dim secondaryKeyValue As String = ""

        Dim insuranceProuctKey As Integer = 0


        xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True

        xlWB = xlApp.Workbooks.Open("D:\Documents\Zestlife\AssetLife\Lead Allocation\Data_out.csv")
        xlWS = xlWB.Sheets(1)
        xlApp.WindowState = Excel.XlWindowState.xlMaximized

        conn.daSQL.SelectCommand = conn.comSQL

        'Array storing each MySQL table to run through
        Dim tableArray() As String = {"insuranceProduct", "asset", "client", "company", "dealership", "clientBank"}
        conn.ds.Clear()

        'Far Each to run through each MySQL table and store the column and datatype
        For Each table In tableArray
            conn.comSQL.CommandText = "SELECT column_name, data_type from information_schema.columns where table_schema = 'AssetLife' and table_name = '" & table & "'"
            conn.daSQL.Fill(conn.ds, table)
            'Adding a column in each table for the Excel column letter
            conn.ds.Tables(table).Columns.Add("columnLetter", Type.GetType("System.String"))
            Dim range As Excel.Range = Nothing
            'For loop looping through each datatable row finding and populating the column letter
            For a = 0 To conn.ds.Tables(table).Rows.Count() - 1
                range = xlWS.Rows("1:1").Find(conn.ds.Tables(table).Rows(a).Item(0), , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False)
                If Not range Is Nothing Then conn.ds.Tables(table).Rows(a).Item(2) = Split(range.Address, "$")(1)
            Next
        Next

        Dim lastrow As Integer
        lastrow = xlWS.Cells(xlWS.Rows.Count, "C").End(Excel.XlDirection.xlUp).Row
        If conn.conSQL.State = ConnectionState.Closed Then conn.conSQL.Open()

        'Looping through each line in Excel
        For i = 3000 To lastrow
            'Looping through each MySQL table
            For Each table In tableArray
                'Looping through each column in the MySQL table
                For Each columnName In conn.ds.Tables(table).Rows

                    'columnName.Item(0) = field name ; columnName.Item(1) = datatype ; columnName.Item(2) = Excel column letter
                    If Not IsDBNull(columnName.Item(2)) Then



                        'If cell is "NA" do not insert anything
                        If CStr(xlWS.Cells(i, columnName.Item(2)).Value()) <> "NA" Then
                            If insertColumns = "" Then
                                insertColumns = columnName.Item(0)
                                insertValues = "'" & xlWS.Cells(i, columnName.Item(2)).Value() & "'"
                            Else
                                insertColumns = insertColumns & ", " & columnName.Item(0)
                                'If statment to format either data or datatime columns
                                If columnName.Item(1) = "date" Then
                                    insertValues = insertValues & ", '" & Format(xlWS.Cells(i, columnName.Item(2)).Value(), "yyyy-MM-dd") & "'"
                                ElseIf columnName.Item(1) = "datetime" Then
                                    insertValues = insertValues & ", '" & Format(xlWS.Cells(i, columnName.Item(2)).Value(), "yyyy-MM-dd hh:mm:ss") & "'"
                                Else
                                    insertValues = insertValues & ", '" & Replace(xlWS.Cells(i, columnName.Item(2)).Value(), "'", "''") & "'"
                                End If
                            End If
                            'If no bank acc then
                        ElseIf table = "clientBank" And columnName.Item(0) = "CLIENTBANKACCNO" Then
                            insertColumns = insertColumns & ", " & columnName.Item(0)
                            conn.comSQL.CommandText = ("SELECT RIGHT(CLIENTBANKACCNO, LENGTH(CLIENTBANKACCNO)-4) + 1  FROM clientBank WHERE LEFT(CLIENTBANKACCNO, 4) = 'FAKE' LIMIT 1")
                            insertValues = insertValues & ", '" & CInt(conn.comSQL.ExecuteScalar()) & "'"
                        End If
                    ElseIf table = "client" And columnName.Item(0) = "ZL_PROD_ID" And insuranceProuctKey <> 0 Then
                        insertColumns = insertColumns & ", " & columnName.Item(0)
                        insertValues = insertValues & ", '" & insuranceProuctKey & "'"
                    End If

                    'Select Case to record the primary key for that table. clientBank has 2 primary keys
                    Select Case table & columnName.Item(0)
                        Case "asset" & "VINNUMBER", "client" & "CLIENTIDNUMBER", "clientbank" & "CLINTBANKNAME", "company" & "COMPANYREGISTRATIONNUMBER", "dealership" & "BRANCHCODE", "insuranceproduct" & "ZL_PROD_ID", "clientBank" & "CLIENTBANKNAME"
                            primaryKeyName = columnName.Item(0)
                            primaryKeyValue = xlWS.Cells(i, columnName.Item(2)).Value()
                        Case "clientBank" & "CLIENTBANKACCNO"
                            secondaryKeyValue = xlWS.Cells(i, columnName.Item(2)).Value()
                    End Select
                Next columnName

                'Insurance product doesn't have a primary key so can be inserted regardless and need to return auto ID
                If table <> "insuranceProduct" Then
                    'Checking duplicates
                    If Not conn.ds.Tables("dupCheck") Is Nothing Then conn.ds.Tables("dupCheck").Clear()
                    If table <> "clientBank" Then
                        conn.comSQL.CommandText = "SELECT COUNT(" & primaryKeyName & ") AS count FROM " & table & " WHERE " & primaryKeyName & " = '" & primaryKeyValue & "'"
                    Else
                        conn.comSQL.CommandText = "SELECT COUNT(CLIENTBANKNAME) AS count FROM " & table & " WHERE CLIENTBANKNAME = '" & primaryKeyValue & "' AND CLIENTBANKACCNO = '" & secondaryKeyValue & "'"
                    End If
                    conn.daSQL.Fill(conn.ds, "dupCheck")

                    'If there are no duplicates then INSERT row for that table
                    If primaryKeyValue <> "" And primaryKeyValue <> "NA" And conn.ds.Tables("dupCheck").Rows(0).Item(0) = "0" Then
                        conn.comSQL.CommandText = "INSERT INTO " & table & " (" & insertColumns & ")  VALUES (" & insertValues & ")"
                        conn.comSQL.ExecuteNonQuery()
                    End If
                Else
                    conn.comSQL.CommandText = ("INSERT INTO " & table & " (" & insertColumns & ")  VALUES (" & insertValues & "); SELECT LAST_INSERT_ID();")
                    insuranceProuctKey = CInt(conn.comSQL.ExecuteScalar())
                End If

                'Clean up after insert
                insertColumns = ""
                insertValues = ""
                primaryKeyName = ""
                primaryKeyValue = ""
                secondaryKeyValue = ""
                If table = "client" Then insuranceProuctKey = 0

            Next table
            Console.Write(" - " & i)

        Next i
        If conn.conSQL.State = ConnectionState.Open Then conn.conSQL.Close()
        Console.WriteLine("End: " & Now.TimeOfDay.ToString)

    End Sub

    Public Sub log(ByVal text As String)
        Dim FILE_NAME As String = "D:\Documents\Zestlife\AssetLife\Lead Allocation\log.txt"

        If System.IO.File.Exists(FILE_NAME) = True Then

            Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True)

            objWriter.WriteLine(Now() & " - " & text)
            objWriter.Close()
        Else
            MessageBox.Show("File Does Not Exist")
        End If
    End Sub
End Class
