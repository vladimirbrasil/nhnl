Attribute VB_Name = "NHNL_Codes"
'Option Explicit

Const CONST_MYUSER As String = "Users\Vla"
Const CONST_PATH_IMAGENS_NHNL As String = "C:\" & CONST_MYUSER & "\Pictures\NH-NL"
    
'Migração
    'Snipping Tool
        'Desativar Snipping Tool Colocar automaticamente imagem na clipboard!
        'Caso contrário, Snipping Tool limpa o texto da clipboard
    'Referências
        'fm20.dll - Clipboard
            'Download fm20.dll (Google)
            'Colocar na pasta system32'
            'Procurar arquivo e adicionar em "Referências"
        'Visual Basic For Applications
        'OLE Automation
        'Microsoft Outlook 12.0 Object Library
        'Microsoft Office 12.0 Object Library
        'Microsoft ActiveX Data Objects 2.8 Library
        'Microsoft Excel 12.0 Object Library
        'Microsoft Scrpting Runtime
        'Microsoft Visual Basic For Applications Extensibility 5.3
        'Microsoft Forms 2.0 Object Library
    'Adaptar pastas do usuário


Sub NHNL1_GetData()
    On Error Resume Next     ' Ative a rotina de tratamento de erro.

Application.ScreenUpdating = False ' turns off screen updating
Application.DisplayStatusBar = True ' makes sure that the statusbar is visible
Application.StatusBar = "Iniciando GetData..."

    Dim QuerySheet As Worksheet
    Dim DataSheet As Worksheet
    Dim StocksSheet As Worksheet
    Dim qurl As String
    Dim i As Integer
    Dim nFirstRow As Double
    Dim nLastRow As Double
    Dim sStock As String
    Dim sStockURL As String
    Dim dDate As Date
    Dim sPrimeiraCelulaInvalida As Double
    Dim nStocks As Integer
    Dim iStock As Integer
    
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
    
    Set DataSheet = Worksheets("Data")
    Set StocksSheet = Worksheets("Stocks")
    
    ActiveWorkbook.RefreshAll
    
    'Criar backup e apagar dos dados anteriores
        'Desvincular dados
    For Each QueryTable In DataSheet.QueryTables
        a = QueryTable.Name
        QueryTable.Delete
    Next
        
    'Apagar
    DataSheet.Activate
    DataSheet.Range("A1:F1000000").Select
    Selection.ClearContents
        
    nStocks = DataSheet.Range("TotalStockNumber").Value
    iStock = 2
    nFirstRow = 1
    dDate = DataSheet.Range("J2").Value + 1
    sAno = Year(dDate)
    sMes = Format(Month(dDate) - 1, "00") 'Yahoo usa (mês - 1) para definir o mês
    sDia = Day(dDate)
    While StocksSheet.Cells(iStock, 1) <> ""
        sStock = StocksSheet.Cells(iStock, 1)
        sStockURL = sStock
        
Application.StatusBar = "Obtendo dados > " & sStock & " (" & iStock - 1 & "\" & nStocks & ")" & "..."

'        If sStock <> "^BVSP" Then
'            sStockURL = sStock & ".SA"
'        End If
        
'Preços Históricos
        qurl = "http://ichart.finance.yahoo.com/table.csv?s=" + sStockURL + "&a=" & sMes & "&b=" & sDia & "&c=" & sAno & "&d=11&e=3&f=2020&g=d"
Application.StatusBar = "Obtendo dados > " & sStock & " (" & iStock - 1 & "\" & nStocks & ")" & " > Preços Históricos..."
        iStock = iStock + 1
        
QueryQuote:
        With DataSheet.QueryTables.Add(Connection:="URL;" & qurl, Destination:=DataSheet.Cells(nFirstRow, 4))
            .BackgroundQuery = True
            .TablesOnlyFromHTML = False
            .Refresh BackgroundQuery:=False
            .SaveData = True
        End With
                            
        nLastRow = DataSheet.Cells(nFirstRow, 4).CurrentRegion.Rows.Count + nFirstRow - 2
        
        'Index - pegar preços e volume
        If StocksSheet.Cells(iStock - 1, 2).Value = "Index" Then
            'Volume
            Call NHNL1_CopyCells(DataSheet, sStock, nLastRow, nFirstRow, "_1V")
            'Open
            Call NHNL1_CopyCells(DataSheet, sStock, nLastRow, nFirstRow, "_2O")
            'High
            Call NHNL1_CopyCells(DataSheet, sStock, nLastRow, nFirstRow, "_3H")
            'Low
            Call NHNL1_CopyCells(DataSheet, sStock, nLastRow, nFirstRow, "_4L")
            'Close
            Call NHNL1_CopyCells(DataSheet, sStock, nLastRow, nFirstRow, "_5C")
            nFirstRow = nLastRow + 3
        Else
            Call NHNL1_FillCells(DataSheet, sStock, nFirstRow, nLastRow, 1)
            'Formar colunas
            Call NHNL1_ToColumns(DataSheet, sStock, nLastRow, nFirstRow, "")
            nFirstRow = nLastRow + 3
        End If

'Preço atual
        If DataSheet.Range("Country").Value = "Brasil" Then
        'Somente se depois das 19h
            If Time() > CDate("23:59") Or Time() < CDate("0:00") Then '18:40 8:00
                qurl = "http://download.finance.yahoo.com/d/quotes.csv?s=" & sStock & "&f=d1ohgl1v&e=.csv"
    Application.StatusBar = "Obtendo dados > " & sStock & " (" & iStock - 2 & "\" & nStocks & ")" & " > Preço Atual..."
                    
                With DataSheet.QueryTables.Add(Connection:="URL;" & qurl, Destination:=DataSheet.Cells(nFirstRow, 4))
                    .BackgroundQuery = True
                    .TablesOnlyFromHTML = False
                    .Refresh BackgroundQuery:=False
                    .SaveData = True
                End With
                                    
                nLastRow = DataSheet.Cells(nFirstRow, 4).CurrentRegion.Rows.Count + nFirstRow - 2
                
                'Index - pegar preços e volume
                If StocksSheet.Cells(iStock - 1, 2).Value = "Index" Then
        '            'Volume - Volume aparece sempre zero em preços atuais. Preencher somente em Preços Históricos, no dia seguinte.
        '            Call NHNL1_CopyCells_TodaysData(DataSheet, sStock, nLastRow, nFirstRow, "_1V")
                    'Open
                    Call NHNL1_CopyCells_TodaysData(DataSheet, sStock, nLastRow, nFirstRow, "_2O")
                    'High
                    Call NHNL1_CopyCells_TodaysData(DataSheet, sStock, nLastRow, nFirstRow, "_3H")
                    'Low
                    Call NHNL1_CopyCells_TodaysData(DataSheet, sStock, nLastRow, nFirstRow, "_4L")
                    'Close
                    Call NHNL1_CopyCells_TodaysData(DataSheet, sStock, nLastRow, nFirstRow, "_5C")
                    nFirstRow = nLastRow + 3
                Else
                    Call NHNL1_FillCells_TodaysData(DataSheet, sStock, nFirstRow, nLastRow, 1)
                    'Formar colunas
                    Call NHNL1_ToColumns(DataSheet, sStock, nLastRow, nFirstRow, "")
                    nFirstRow = nLastRow + 3
                End If
            End If 'Depois das 19h
        End If 'Country "Brasil"

NextStock:
    Wend
        
Application.StatusBar = "Organizando dados..."
        
    DataSheet.Range("A1").Select
    DataSheet.Range("A1") = "Stock"
    DataSheet.Range("B1").Select
    DataSheet.Range("B1") = "Date"
    DataSheet.Range("C1").Select
    DataSheet.Range("C1") = "Close"
    
Application.StatusBar = "Organizando dados > Apagando Query Tables" & "..."
    
    For Each QueryTable In DataSheet.QueryTables
        a = QueryTable.Name
        QueryTable.Delete
    Next
    
    Columns("D:D").Select
    Selection.ClearContents
    
Application.StatusBar = "Organizando dados > Ordenando dados" & "..."
    
    'Ordenar
    Columns("A:C").Select
    DataSheet.Sort.SortFields.Clear
    DataSheet.Sort.SortFields.Add Key:=Range("A2:A1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    DataSheet.Sort.SortFields.Add Key:=Range("B2:B1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With DataSheet.Sort
        .SetRange Range("A1:C1000000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'    ActiveWorkbook.Save
    
Application.StatusBar = "Organizando dados > Atualizando tabela dinâmica" & "..."
    
    ActiveWorkbook.RefreshAll
    
'Apagar cotações em feriados (repetem dia anterior)
    'Falha do yahoo. Algumas ações

    '"H1" informa primeira linha com dados repetidos (atualizar)
    Cells.Select
    Cells.Calculate
'    Application.Calculation = xlCalculationAutomatic
'    Application.Calculation = xlCalculationManual

Application.StatusBar = "Organizando dados > Classificando por dados válidos" & "..."

    'Descobrir e apagar células inválidas (feriados)
    Columns("A:I").Select
    DataSheet.Sort.SortFields.Clear
    DataSheet.Sort.SortFields.Add Key:=Range("I2:I1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    DataSheet.Sort.SortFields.Add Key:=Range("B2:B1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With DataSheet.Sort
        .SetRange Range("A1:I1000000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Application.StatusBar = "Organizando dados > Calculando e apagando dados inválidos" & "..."
    
    'Atualizar célula "H1"
    Cells.Select
    Cells.Calculate
'    Application.Calculation = xlCalculationAutomatic
'    Application.Calculation = xlCalculationManual
    
    sPrimeiraCelulaInvalida = Range("H1").Value
   
    DataSheet.Activate
    Range("A" & sPrimeiraCelulaInvalida & ":C10000").Select
    Selection.ClearContents
   
Application.StatusBar = "Organizando dados > Classificando dados válidos" & "..."
   
    'Ordenar
    Columns("A:C").Select
    DataSheet.Sort.SortFields.Clear
    DataSheet.Sort.SortFields.Add Key:=Range("A2:A1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    DataSheet.Sort.SortFields.Add Key:=Range("B2:B1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With DataSheet.Sort
        .SetRange Range("A1:C1000000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
Application.StatusBar = "Finalizando" & "..."
   
    'Enquanto não capturar lastdate do banco de dados
        'Capturar também valid dates utilizando a resposta de um dos índices - tabela dinâmica
    DataSheet.Range("J2").Value = DataSheet.Range("B2").Value - 2
   
   
    'Estética
    DataSheet.Cells.EntireColumn.AutoFit
    
    'Turn calculation back on
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
   
Application.StatusBar = False ' gives control of the statusbar back to the programme
Application.ScreenUpdating = True
   
Exit Sub        ' Saia para evitar manipulador.
ErrorHandler:    ' Rotina de tratamento de erro.
    Select Case Err.Number    ' Avalie o número do erro.
        Case 1004    ' Arquivo não existe no site do Yahoo
            MsgBox ("Erro " & Err.Number & ": " & Err.Description)
            GoTo NextStock     ' Feche o arquivo aberto.
        Case Else
            MsgBox ("Erro " & Err.Number & ": " & Err.Description)
            GoTo NextStock     ' Feche o arquivo aberto.
            ' Trate outras situações aqui...
    End Select
    Resume    ' Continue a execução na mesma linha
                ' que provocou o erro.
End Sub

Sub NHNL1_FillCells(pSheet As Worksheet, sTexto As String, i1 As Double, i2 As Double, j As Integer)
   For k = i1 + 1 To i2 'i1 + 1 para ignorar cabeçalho em preços históricos
        pSheet.Cells(k, j) = sTexto
   Next k
End Sub

Sub NHNL1_FillCells_TodaysData(pSheet As Worksheet, sTexto As String, i1 As Double, i2 As Double, j As Integer)
   For k = i1 To i2
        pSheet.Cells(k, j) = sTexto
   Next k
End Sub

Sub NHNL1_CopyCells(ByRef DataSheet As Worksheet, sStock As String, ByRef nLastRow As Double, ByRef nFirstRow As Double, sPriceText As String)
     
    Call NHNL1_FillCells(DataSheet, "_" & sStock & sPriceText, nFirstRow, nLastRow, 1)
    
    'Não precisa copiar células para o último preço (5-Close)
    If Right$(sPriceText, 2) = "5C" Then
        'Formar colunas
        Call NHNL1_ToColumns(DataSheet, sStock, nLastRow, nFirstRow, sPriceText)
        Exit Sub
    End If
        
    'Copiar intervalo para baixo
    DataSheet.Cells(nLastRow, 4).Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Copy
    DataSheet.Cells(nLastRow + 3, 4).Select
    ActiveSheet.Paste
    
    'Formar colunas
    Call NHNL1_ToColumns(DataSheet, sStock, nLastRow, nFirstRow, sPriceText)
    
    nFirstRow = nLastRow + 3
    nLastRow = DataSheet.Cells(nFirstRow, 4).CurrentRegion.Rows.Count + nFirstRow - 1

End Sub

Sub NHNL1_CopyCells_TodaysData(ByRef DataSheet As Worksheet, sStock As String, ByRef nLastRow As Double, ByRef nFirstRow As Double, sPriceText As String)
     
    Call NHNL1_FillCells_TodaysData(DataSheet, "_" & sStock & sPriceText, nFirstRow, nLastRow, 1)
    
    'Não precisa copiar células para o último preço (5-Close)
    If Right$(sPriceText, 2) = "5C" Then
        'Formar colunas
        Call NHNL1_ToColumns(DataSheet, sStock, nLastRow, nFirstRow, sPriceText)
        Exit Sub
    End If
        
    'Copiar intervalo para baixo
    DataSheet.Cells(nLastRow, 4).Select
'    Range(Selection, Selection.End(xlUp)).Select
    Selection.Copy
    DataSheet.Cells(nLastRow + 3, 4).Select
    ActiveSheet.Paste
    
    'Formar colunas
    Call NHNL1_ToColumns(DataSheet, sStock, nLastRow, nFirstRow, sPriceText)
    
    nFirstRow = nLastRow + 3
    nLastRow = DataSheet.Cells(nFirstRow, 4).CurrentRegion.Rows.Count + nFirstRow - 1

End Sub

Sub NHNL1_ToColumns_NHNLNow(ByRef DataSheet As Worksheet, sStock As String, ByRef nLastRow As Double, ByRef nFirstRow As Double, sPriceText As String)
    Dim sDestRange As String
    
    sDestRange = "B" & nFirstRow
    sTableRange = "G" & nFirstRow & ":" & "G" & nLastRow
    
    'Substituir cabeçalhos de texto
    Range(sTableRange).Select
    Selection.Replace What:="Date,Open,High,Low,Close,Volume,Adj Close", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    'Substituir vírgula por ponto-e-vírgula
'    Columns("D:D").Select
    Range(sTableRange).Select
    Selection.Replace What:=",", Replacement:=";", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'Substituir ponto por vírgula
'    Columns("D:D").Select
    Range(sTableRange).Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    'Transformar texto para colunas
'    Columns("D:D").Select
    Range(sTableRange).Select
    

    Selection.TextToColumns Destination:=Range(sDestRange), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
        

End Sub

Sub NHNL1_ToColumns(ByRef DataSheet As Worksheet, sStock As String, ByRef nLastRow As Double, ByRef nFirstRow As Double, sPriceText As String)
    Dim sDestRange As String
    
    sDestRange = "B" & nFirstRow
    sTableRange = "D" & nFirstRow & ":" & "D" & nLastRow
    
    'Substituir cabeçalhos de texto
    Range(sTableRange).Select
    Selection.Replace What:="Date,Open,High,Low,Close,Volume,Adj Close", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    'Substituir vírgula por ponto-e-vírgula
'    Columns("D:D").Select
    Range(sTableRange).Select
    Selection.Replace What:=",", Replacement:=";", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'Substituir ponto por vírgula
'    Columns("D:D").Select
    Range(sTableRange).Select
    Selection.Replace What:=".", Replacement:=",", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
    'Transformar texto para colunas
'    Columns("D:D").Select
    Range(sTableRange).Select
    
    Select Case Right$(sPriceText, 2)
        Case "1V"
            'Incluir outros preços e volume
            Selection.TextToColumns Destination:=Range(sDestRange), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                :=Array(Array(1, 5), Array(2, 9), Array(3, 9), Array(4, 9), Array(5, 9), Array(6, 1), _
                Array(7, 9)), TrailingMinusNumbers:=True
        Case "2O"
            'Incluir outros preços e volume
            Selection.TextToColumns Destination:=Range(sDestRange), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                :=Array(Array(1, 5), Array(2, 1), Array(3, 9), Array(4, 9), Array(5, 9), Array(6, 9), _
                Array(7, 9)), TrailingMinusNumbers:=True
        Case "3H"
            'Incluir outros preços e volume
            Selection.TextToColumns Destination:=Range(sDestRange), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                :=Array(Array(1, 5), Array(2, 9), Array(3, 1), Array(4, 9), Array(5, 9), Array(6, 9), _
                Array(7, 9)), TrailingMinusNumbers:=True
        Case "4L"
            'Incluir outros preços e volume
            Selection.TextToColumns Destination:=Range(sDestRange), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                :=Array(Array(1, 5), Array(2, 9), Array(3, 9), Array(4, 1), Array(5, 9), Array(6, 9), _
                Array(7, 9)), TrailingMinusNumbers:=True
'        Case "5C"
'            'Incluir outros preços e volume
'            Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                :=Array(Array(1, 5), Array(2, 9), Array(3, 9), Array(4, 9), Array(5, 1), Array(6, 9), _
                Array(7, 9)), TrailingMinusNumbers:=True
        Case Else
            Selection.TextToColumns Destination:=Range(sDestRange), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
                :=Array(Array(1, 5), Array(2, 9), Array(3, 9), Array(4, 9), Array(5, 1), Array(6, 9), _
                Array(7, 9)), TrailingMinusNumbers:=True
        
    End Select
    
End Sub

Sub NHNL1_RefreshDB()


    'Conectar ao db
    Dim sourcePath As String
    Dim Conn As ADODB.Connection
    Dim strConn As String
    Dim strSQL As String
    
    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ActiveWorkbook.FullName & ";" & "Extended Properties=Excel 8.0;"
    Set Conn = New ADODB.Connection
    
    'strAccessConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sourcePath
    
    
    
    
' REQUIREMENTS: Access 2000 / 2002 / 2003
'               VBA Reference to the Microsoft ActiveX Data Objects Library

    sourceFileName = Worksheets("Data").Range("dbName") '"Stocks-2-Quotes.accdb"
    sourcePath = ActiveWorkbook.Path & "\" & sourceFileName

Dim g_ADOBackEndConn As ADODB.Connection
Dim g_BackEndADOConnectStr As String
Const g_BackendPassword = ""

' For Microsoft.ACE.OLEDB.12.0,you need Microsoft Office 12.0 Access Database Engine to be installed.
Const g_BackendADOProvider = "Microsoft.ACE.OLEDB.12.0"
    g_BackEndADOConnectStr = "Provider=" & g_BackendADOProvider & ";" & _
                                        "Data Source=" & sourcePath & ";" & _
                                        "Jet OLEDB:Database Password=" & g_BackendPassword
    
    Conn.Open g_BackEndADOConnectStr
        
'    On Error Resume Next
    sSQLDropTable1 = "DROP TABLE Excel_HistData"
    sSQLDropTable2 = "DROP TABLE Excel_HistData_DistinctData"
    sSQLDropTable3 = "DROP TABLE Excel_HistData_NewData"
    sSQLCreateTable = "SELECT * INTO Excel_HistData" & _
           " FROM rngStockData IN '' 'Excel 8.0;database=" & ActiveWorkbook.FullName & "';"
    
    Conn.Execute sSQLDropTable1
    Conn.Execute sSQLDropTable2
    Conn.Execute sSQLDropTable3
    Conn.Execute sSQLCreateTable
    
    Dim objCommand As ADODB.Command
    Set objCommand = New ADODB.Command
    With objCommand
        Set .ActiveConnection = Conn ' Reference to a Connection object.
        .CommandType = adCmdStoredProc
        .CommandText = "Excel_HistData_CreateTableExcelDistinctData"
        objCommand.Execute
        .CommandText = "Excel_HistData_CreateTableExcelNewData"
        objCommand.Execute
        .CommandText = "Excel_HistData_AddNewDataToQuotes"
        objCommand.Execute
    End With

    Set Conn = Nothing
    
    
'***    'Funciona no Access! Somente falta configurar com tabela do excel!
'    sSQLCreateTable = "SELECT UpdTbl.Stock, UpdTbl.Date, UpdTbl.Close INTO NewTable FROM UpdTbl LEFT JOIN Quotes ON (UpdTbl.Stock = Quotes.Stock) AND (UpdTbl.Date = Quotes.QuoteDate) WHERE (((UpdTbl.Stock)<>"") AND ((Quotes.Quote) Is Null));"
'***    'Tentativas de solução em uma única execução!!! Adaptar SQL acima com tabela do Excel.
'    sSQLCreateTable = "SELECT UpdTbl.Stock, UpdTbl.Date, UpdTbl.Close INTO NewTable FROM UpdTbl LEFT JOIN Quotes ON (UpdTbl.Stock = Quotes.Stock) AND (UpdTbl.Date = Quotes.QuoteDate) WHERE (((UpdTbl.Stock)<>"") AND ((Quotes.Quote) Is Null));"
'    sSQLCreateTable = "SELECT [Stock], [Date], [Close] INTO UpdTbl FROM rngStockData IN '' 'Excel 8.0;database=" & ActiveWorkbook.FullName & "' LEFT JOIN Quotes ON ([Stock] = Quotes.Stock) AND ([Date] = Quotes.QuoteDate) WHERE ((([Stock])<>'') AND ((Quotes.Quote) Is Null));"
'    sSQLCreateTable = "SELECT [Stock], [Date], [Close] INTO UpdTbl FROM rngStockData IN '' 'Excel 8.0;database=" & ActiveWorkbook.FullName & "' LEFT JOIN Quotes ON ([Close] = Quotes.Quote) WHERE (((Quotes.Quote) Is Null));"

        
'    sSQLupdatequotes = "INSERT INTO Quotes " & _
            "( Stock, QuoteDate, Quote ) " & _
            "SELECT " & _
            "colStockData.Stock, " & _
            "colDateData.Date, " & _
            "colCloseData2.Close " & _
            "FROM colStockData IN ''" & _
            " 'Excel 8.0;database=" & ActiveWorkbook.FullName & "'," & _
            " colDateData IN ''" & _
            " 'Excel 8.0;database=" & ActiveWorkbook.FullName & "'," & _
            " colCloseData2 IN ''" & _
            " 'Excel 8.0;database=" & ActiveWorkbook.FullName & "';"
'            "FROM UpdTbl " & _
'            "WHERE (((UpdTbl.Close)<>0));"
        
'    sSQLupdatequotes = "INSERT INTO Quotes ( Stock, QuoteDate, Quote ) " & _
            "(SELECT Stock FROM colStockData IN '' 'Excel 8.0;database=" & ActiveWorkbook.FullName & "')," & _
            "(SELECT Date FROM colDateData IN '' 'Excel 8.0;database=" & ActiveWorkbook.FullName & "')," & _
            "(SELECT Close FROM colCloseData2 IN '' 'Excel 8.0;database=" & ActiveWorkbook.FullName & "');"

        
'        sSQL = "INSERT INTO Quotes " & _
               "( Stock, QuoteDate, Quote )" & _
               "SELECT " & _
               "RecentHistoricalData.Stock, " & _
               "RecentHistoricalData.Date, " & _
               "RecentHistoricalData.Close" & _
               "FROM RecentHistoricalData" & _
               "WHERE (((RecentHistoricalData.Close)<>0));"
        
'INSERT INTO table1 ( column1, column2, 8, 'some string etc.' )
'SELECT  table2.column1, table2.column2
'FROM table2
'WHERE table2.ID = 7
        
'    sSQLCreateTable = "SELECT * INTO UpdTbl" & _
           " FROM [Data$] IN ''" & _
           " 'Excel 8.0;database=" & ActiveWorkbook.FullName & "';"
        
'    sSQL = "UPDATE C1R0 " & _
           "SET " & _
           "col1 = ( SELECT col1 FROM [C1R0$] IN 'Excel 8.0;database=c:\excel\UpdateFinal1.xls'; " & _
           "WHERE [C1R0$].PK = C1R0.PK ), " & _
           "col2 = ( SELECT col2 FROM [C1R0$] IN 'Excel 8.0;database=c:\excel\UpdateFinal1.xls'; " & _
           "WHERE [C1R0$].PK = C1R0.PK ),"
        
'    sExcelSQL = "SELECT * FROM rngStockData;"
'    Set objConn = NHNL1_GetExcelConnection(sourcePath)
'    objRS.Open sExcelSQL, objConn
    
End Sub

Private Function NHNL1_GetExcelConnection(ByVal Path As String, _
    Optional ByVal Headers As Boolean = True) As Connection
    Dim strConn As String
    Dim objConn As ADODB.Connection
    Set objConn = New ADODB.Connection
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=" & Path & ";" & _
              "Extended Properties=""Excel 8.0;HDR=" & _
              IIf(Headers, "Yes", "No") & """"
    objConn.Open strConn
    Set GetExcelConnection = objConn
End Function

'Codes for NHNL-3-Charts Files

Sub NHNL3_ConstruirGraficoSemanal()
'
' ConstruirGraficoSemanal Macro
'

'
    Dim dDateLastPlot As Date
    Dim sDate As String

    ActiveWorkbook.RefreshAll
    
    Application.Calculation = xlCalculationManual

    Dim chDailyChart As Chart
    Dim chDailyChartZoom As Chart
    Dim chWeeklyChart As Chart
    Dim chWeeklyChartZoom As Chart
   
    Set chDailyChart = Sheets("Daily Chart").ChartObjects("Daily Chart").Chart
    Set chDailyChartZoom = Sheets("Daily Chart").ChartObjects("Daily Chart Zoom").Chart
    Set chWeeklyChart = Sheets("Weekly Chart").ChartObjects("Weekly Chart").Chart
    Set chWeeklyChartZoom = Sheets("Weekly Chart").ChartObjects("Weekly Chart Zoom").Chart

    Sheets("Daily Chart").Select
    Range("WeeklyData").Select
    Selection.Copy
    Sheets("Weekly Chart").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Algumas falhas ao colar dados semanais. Repetir para ter certeza
'    Range("A1").Select
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Application.CutCopyMode = False
    
    'Substituir "#VALOR! e #N/D"
    ActiveSheet.Cells.Select
    Cells.Replace What:="#VALOR!", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    'Ordenar
    ActiveWorkbook.Worksheets("Weekly Chart").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Weekly Chart").Sort.SortFields.Add Key:=Range _
        ("A2:A2000"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Weekly Chart").Sort
        .SetRange Range("A1:H2000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'    sUltimaCelula = Range("I2").Value - 1

'    chWeeklyChart.Activate
'    With chWeeklyChart
'        .SeriesCollection(1).Values = "='Weekly Chart'!$C$2:$C$" & .Range("IndexLimit_WeeklyMainChart").Value
'        .SeriesCollection(2).Values = "='Weekly Chart'!$B$2:$B$" & sUltimaCelula
'        .SeriesCollection(2).XValues = "='Weekly Chart'!$A$2:$A$" & sUltimaCelula
'    End With
'
'    chWeeklyChartZoom.Activate
'    With chWeeklyChart
'        .SeriesCollection(1).Values = "='Weekly Chart'!$B$2:$B$" & sUltimaCelula
'        .SeriesCollection(2).Values = "='Weekly Chart'!$A$2:$A$" & sUltimaCelula
'        .SeriesCollection(2).XValues = "='Weekly Chart'!$C$2:$C$" & sUltimaCelula
'    End With
    
    Cells.Select
    Calculate
    
    'Atualizar eixos dos gráficos
    Sheets("Daily Chart").Activate
    With ActiveSheet
        Call NHNL3_UpdateChartAxis(chDailyChart, _
                             .Range("IndexLimit_DailyMainChart").Value, _
                             .Range("NHNLLimit_DailyMainChart").Value, _
                             .Range("InitialDate_DailyMainChart").Value, _
                             .Range("UltimaCelula_Daily").Value _
                             )
        Call NHNL3_UpdateChartAxis(chDailyChartZoom, _
                             .Range("IndexLimit_DailyZoomChart").Value, _
                             .Range("NHNLLimit_DailyZoomChart").Value, _
                             .Range("InitialDate_DailyZoomChart").Value, _
                             .Range("UltimaCelula_DailyZoom").Value _
                             )
    End With
    Sheets("Weekly Chart").Activate
    With ActiveSheet
        Call NHNL3_UpdateChartAxis(chWeeklyChart, _
                             .Range("IndexLimit_WeeklyMainChart").Value, _
                             .Range("NHNLLimit_WeeklyMainChart").Value, _
                             .Range("InitialDate_WeeklyMainChart").Value, _
                             .Range("UltimaCelula_Weekly").Value _
                             )
        Call NHNL3_UpdateChartAxis(chWeeklyChartZoom, _
                             .Range("IndexLimit_WeeklyZoomChart").Value, _
                             .Range("NHNLLimit_WeeklyZoomChart").Value, _
                             .Range("InitialDate_WeeklyZoomChart").Value, _
                             .Range("UltimaCelula_WeeklyZoom").Value _
                             )
    End With
    
    Call NHNL3_ColocarClipboard("Weekly")
        
    'Inserir data no título do gráfico
    Call AtualizarTituloGraficos
    
    'Diário
    Sheets("Daily Chart").Select
'    ActiveSheet.ChartObjects("Daily Chart").Activate
'    sDate = fForceDateToEnglish(Format$(CDate(Range("TillDate_Daily").Value), "d mmm"))
'    ActiveChart.ChartTitle.Text = Range("CountryName").Value & "              " & sDate & Chr(13) & "-----------------------------------"
    Range("A1").Select
    
    'Semanal
    Sheets("Weekly Chart").Select
'    ActiveSheet.ChartObjects("Weekly Chart").Activate
'    sDate = fForceDateToEnglish(Format$(CDate(Range("TillDate_Weekly").Value), "d mmm"))
'    ActiveChart.ChartTitle.Text = Range("CountryName").Value & "              " & sDate & Chr(13) & "-----------------------------------"
    Range("A1").Select
        
    Application.Calculation = xlCalculationAutomatic

End Sub

Sub AtualizarTituloGraficos()
    'Inserir data no título do gráfico
    'Diário
    Sheets("Daily Chart").Select
    ActiveSheet.ChartObjects("Daily Chart").Activate
    sDate = fForceDateToEnglish(Format$(CDate(Range("TillDate_Daily").Value), "d mmm"))
    ActiveChart.ChartTitle.Text = Range("CountryName").Value & "              " & sDate & Chr(13) & ".----------------------------------"
    Range("A1").Select
    
    'Semanal
    Sheets("Weekly Chart").Select
    ActiveSheet.ChartObjects("Weekly Chart").Activate
    sDate = fForceDateToEnglish(Format$(CDate(Range("TillDate_Weekly").Value), "d mmm"))
    ActiveChart.ChartTitle.Text = Range("CountryName").Value & "              " & sDate & Chr(13) & ".----------------------------------"
    Range("A1").Select
        
End Sub

Sub NHNL3_FileNaClipboard_Daily()
    Call NHNL3_ColocarClipboard("Daily")

End Sub

Sub NHNL3_FileNaClipboard_Weekly()
    Call NHNL3_ColocarClipboard("Weekly")

End Sub

Function fGerarNomeArquivoOld(sChartDailyOrWeekly As String)
    Dim myCountry, sChartDW, MyPath, MyIndexName As String
    Dim sFileDate, myMiddleName, sPathCompleto As String
    Dim dDate As Date
    '\ImagensOld\
    
    myCountry = Sheets("Daily Chart").Range("CountryRIP").Value
    sChartDW = Left$(sChartDailyOrWeekly, 1)
    MyPath = CONST_PATH_IMAGENS_NHNL & "\ImagensOld\"
    MyIndexName = Sheets("Daily Chart").Range("IndexName").Value
    dDate = Sheets("Daily Chart").Range("TillDate_DailyCharts").Value
    sFileDate = Format(dDate, "yyyy") & Format(dDate, "MM") & Format(dDate, "dd")
    
    myMiddleName = myCountry & "-" & sChartDW & "-VladimirDietrichkeitRightsReserved"

    sPathCompleto = MyPath & myMiddleName & ".jpg"
    
    fGerarNomeArquivoOld = sPathCompleto
    
End Function

Function fGerarNomeArquivo(sChartDailyOrWeeklyZoomOrNot As String)
    Dim myCountry, sChartDW, MyPath, MyIndexName As String
    Dim sFileDate, myMiddleName, sPathCompleto As String
    Dim dDate As Date
    '\Imagens\
    
    If Right$(sChartDailyOrWeeklyZoomOrNot, 4) = "Zoom" Then sZoom = "-Zoom"
    
    myCountry = Sheets("Daily Chart").Range("CountryRIP").Value
    sChartDW = Left$(sChartDailyOrWeeklyZoomOrNot, 1)
    MyPath = CONST_PATH_IMAGENS_NHNL & "\Imagens\" '"C:\Users\Vla\Pictures\NH-NL\Imagens\"
'    MyIndexName = Sheets("Daily Chart").Range("IndexName").Value
'    dDate = Sheets("Daily Chart").Range("TillDate_DailyCharts").Value
'    sFileDate = Format(dDate, "yyyy") & Format(dDate, "MM") & Format(dDate, "dd")
    
    myMiddleName = myCountry & "-" & sChartDW & sZoom & "-VladimirDietrichkeitRightsReserved"

    sPathCompleto = MyPath & myMiddleName & ".gif"
    
    fGerarNomeArquivo = sPathCompleto
    
End Function


Sub NHNL3_ColocarClipboard(sChartDailyOrWeekly As String)
    Dim MyData As DataObject
    
    sPathArquivo = fGerarNomeArquivoOld(sChartDailyOrWeekly)
    
    Set MyData = New DataObject
    MyData.SetText sPathArquivo
        
    'Desativar Snipping Tool Colocar imagem na clipboard automaticamente!
    'Caso contrário, Snipping Tool limpa o texto da clipboard
    MyData.PutInClipboard


End Sub

Sub NHNL3_UpdateChartAxis(chChart As Chart, nChartLimitAxis_Index, nChartLimitAxis_NHNL, nChartLimitAxis_InitialDate, sUltimaCelula)
    Dim sDayOrWeek As String
    Dim sMainOrZoom As String
    
    If InStr(1, chChart.Parent.Name, "Weekly") > 0 Then
        sDayOrWeek = "WeeklyChart"
    Else
        sDayOrWeek = "DailyChart"
    End If
    
    If InStr(1, chChart.Parent.Name, "Zoom") > 0 Then
        sMainOrZoom = "Zoom"
    Else
        sMainOrZoom = ""
    End If
    
'    ActiveWorkbook.Names("DailyChart_DateInterval").RefersToRange = ""
'        .Name = "CountryName"
'        .RefersToR1C1 = "='NH-NL'!R2C1"
'        .Comment = ""
'    End With
    
    
    With chChart
            'Last Value - X Values
            .SeriesCollection(1).XValues = Range(sDayOrWeek & sMainOrZoom & "_DateInterval").Value
            .SeriesCollection(2).XValues = Range(sDayOrWeek & sMainOrZoom & "_DateInterval").Value
            .SeriesCollection(3).XValues = Range(sDayOrWeek & sMainOrZoom & "_DateInterval").Value
            .SeriesCollection(4).XValues = Range(sDayOrWeek & sMainOrZoom & "_DateInterval").Value
            .SeriesCollection(5).XValues = Range(sDayOrWeek & sMainOrZoom & "_DateInterval").Value
            .SeriesCollection(6).XValues = Range(sDayOrWeek & sMainOrZoom & "_DateInterval").Value
            .SeriesCollection(7).XValues = Range(sDayOrWeek & sMainOrZoom & "_DateInterval").Value
            .SeriesCollection(8).XValues = Range(sDayOrWeek & sMainOrZoom & "_DateInterval").Value
                    
            'Last Value - Y Values
            .SeriesCollection(1).Values = Range(sDayOrWeek & sMainOrZoom & "_EMAInterval").Value
            .SeriesCollection(2).Values = Range(sDayOrWeek & sMainOrZoom & "_OInterval").Value
            .SeriesCollection(3).Values = Range(sDayOrWeek & sMainOrZoom & "_HInterval").Value
            .SeriesCollection(4).Values = Range(sDayOrWeek & sMainOrZoom & "_LInterval").Value
            .SeriesCollection(5).Values = Range(sDayOrWeek & sMainOrZoom & "_CInterval").Value
            .SeriesCollection(6).Values = Range(sDayOrWeek & sMainOrZoom & "_NHNLInterval").Value
            .SeriesCollection(7).Values = Range(sDayOrWeek & sMainOrZoom & "_NHInterval").Value
            .SeriesCollection(8).Values = Range(sDayOrWeek & sMainOrZoom & "_NLInterval").Value

'        .Axes(xlValue).MinimumScale = -Abs(nChartLimitAxis_Index)
'        .Axes(xlValue).MaximumScale = Abs(nChartLimitAxis_Index)
'        .Axes(xlValue, xlSecondary).MinimumScale = -Abs(nChartLimitAxis_NHNL)
'        .Axes(xlValue, xlSecondary).MaximumScale = Abs(nChartLimitAxis_NHNL)
''        .HasAxis(xlTimeScale) = True
'        .Axes(xlCategory).MinimumScale = nChartLimitAxis_InitialDate
'        nChartLimitAxis_FinalDate = Now() + 0.1 * (Now() - .Axes(xlCategory).MinimumScale)
'        .Axes(xlCategory).MaximumScale = nChartLimitAxis_FinalDate
    
'        .Axes(xlCategory).MajorUnitScale = xlMonths
'        .Axes(xlCategory).MajorUnit = 1
'        .HasAxis(xlValue) = True
    
    
    End With
End Sub

Sub NHNL3_ExportPictures()
    Dim dDate As Date
    Dim MyPath As String
    Dim myMiddleName As String
    Dim sFileDate As String
    Dim sPathCompleto As String
    Dim shWeeklySheet As Worksheet
    
'    Set shWeeklySheet = Sheets("Weekly Chart")
'    Set shDailySheet = Sheets("Daily Chart")
    
'    myCountryName = Sheets("NH-NL").Range("CountryName").Value
'    MyPath = CONST_PATH_IMAGENS_NHNL & "\" & myCountryName & "\" '"C:\Users\Vla\Pictures\NH-NL\" & myCountryName & "\"
'    MyIndexName = Sheets("Daily Chart").Range("IndexName").Value
'    myMiddleName = "- Country " & myCountryName & " NH-NL and " & MyIndexName & " Index- Vladimir Dietrichkeit- Till "
    
    'Daily Charts
'    dDate = Sheets("Daily Chart").Range("TillDate_DailyCharts").Value
    For Each Chart In Sheets("Daily Chart").ChartObjects
        sPathCompleto = fGerarNomeArquivo(Chart.Name)
        
        
'        sFileDate = Format(dDate, "yyyy") & Format(dDate, "MM") & Format(dDate, "dd")
'        sName = Chart.Name
'        sPathCompleto = MyPath & Chart.Name & myMiddleName & sFileDate & ".gif"
        
        Sheets("Daily Chart").ChartObjects(Chart.Name).Chart.Export sPathCompleto
    Next
    
    'Weekly Charts
'    dDate = Sheets("Weekly Chart").Range("TillDate_WeeklyCharts").Value
    For Each Chart In Sheets("Weekly Chart").ChartObjects
        sPathCompleto = fGerarNomeArquivo(Chart.Name)
        
        
        
'        sFileDate = Format(dDate, "yyyy") & Format(dDate, "MM") & Format(dDate, "dd")
'        sName = Chart.Name
'        sPathCompleto = MyPath & Chart.Name & myMiddleName & sFileDate & ".gif"
        
        Sheets("Weekly Chart").ChartObjects(Chart.Name).Chart.Export sPathCompleto
    Next
    
End Sub

Sub NHNL_ChangeCountry()

    'Definir "CountryName" de acordo com o nome do arquivo
    Sheets("NH-NL").Range("CountryName").Value = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, "- ") - 1)
    
    MyPath = ActiveWorkbook.Path & "\" '"C:\Users\Vla\Pictures\NH-NL\" & MyCountryName & "\"
    myCountryName = Sheets("NH-NL").Range("CountryName").Value
    
    'Conexão
    Sheets("Tabela").Select
    Range("A4").Select
    ActiveWorkbook.Connections.Add _
                   myCountryName & "- NHNL-2-Quotes", _
                   "", Array("OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;" & _
                   "Data Source=" & MyPath & myCountryName & "- NHNL-2-Quotes.accdb;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False"), _
                   Array("QuotesReduzida"), 3
    ActiveSheet.PivotTables("Tabela dinâmica1").ChangeConnection ActiveWorkbook. _
                Connections(myCountryName & "- NHNL-2-Quotes")
    
    'Títulos dos gráficos
        'Daily Chart
    MyIndexName = Sheets("Daily Chart").Range("IndexName").Value
    Sheets("Daily Chart").Activate
    Range("A1").Select
    Sheets("Daily Chart").ChartObjects("Daily Chart").Activate
    ActiveChart.ChartTitle.Text = "Daily " & MyIndexName & " and NH-NL              " & myCountryName
        'Weekly Chart
    Sheets("Weekly Chart").Activate
    Range("A1").Select
    Sheets("Weekly Chart").ChartObjects("Weekly Chart").Activate
    ActiveChart.ChartTitle.Text = "Weekly " & MyIndexName & " and NH-NL              " & myCountryName
        
   'Intervalo de cálculo NH-NL
    StrTemp = Sheets("Tabela").PivotTables("Tabela dinâmica1").ColumnRange.Address
        'Encontrar última coluna da tabela dinâmica
    myLastTableColumn = Left$(Right$(StrTemp, Len(StrTemp) - (InStr(StrTemp, ":$")) - 1), 2)
    myLastNHNLColumn = Sheets("NH-NL").Range("LastNHNLColumn").Value
    
    Sheets("NH-NL").Select
    'GC (number) em Sheets("NH-NL").Range("LastNHNLColumn").Value
    Range(Cells(1, myLastNHNLColumn), Cells(800, myLastNHNLColumn)).Select
'    Range("GC1:GC800").Select
    Selection.Copy
    If Range(myLastTableColumn & "1").Column >= myLastNHNLColumn Then
        'Copiar fórmulas nas células que faltam
'        Range("GC1:" & myLastTableColumn & "800").Select
        Range(Cells(1, myLastNHNLColumn), Cells(800, Range(myLastTableColumn & "1").Column)).Select
        ActiveSheet.Paste
    Else
        'Apagar fórmulas nas células que sobram
        Range(Cells(1, Range(myLastTableColumn & "1").Column + 1), Cells(800, myLastNHNLColumn)).Select
        Selection.ClearContents
    End If
'    Range("A1").Select

End Sub

Sub NHNL4_GetNHNLNow()
    On Error Resume Next     ' Ative a rotina de tratamento de erro.

Application.ScreenUpdating = False ' turns off screen updating
Application.DisplayStatusBar = True ' makes sure that the statusbar is visible
Application.StatusBar = "Iniciando GetData..."

    Dim QuerySheet As Worksheet
    Dim DataSheetOld As Worksheet
    Dim DataSheet As Worksheet
    Dim StocksSheet As Worksheet
    Dim qurl As String
    Dim i As Integer
    Dim nFirstRow As Double
    Dim nLastRow As Double
    Dim sStock As String
    Dim sStockURL As String
    Dim dDate As Date
    Dim sPrimeiraCelulaInvalida As Double
    Dim nStocks As Integer
    Dim iStock As Integer
    
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
    
    Set DataSheet = Worksheets("NHNLNow")
    Set DataSheetOld = Worksheets("Data")
    Set StocksSheet = Worksheets("Stocks")
    
'    ActiveWorkbook.RefreshAll
    
    'Criar backup e apagar dos dados anteriores
        'Desvincular dados
    For Each QueryTable In DataSheet.QueryTables
        a = QueryTable.Name
        QueryTable.Delete
    Next
        
    'Apagar
    DataSheet.Activate
    DataSheet.Range("A1:G1000000").Select
    Selection.ClearContents
        
    nStocks = DataSheetOld.Range("TotalStockNumber").Value
    iStock = 2
    nFirstRow = 1
'    dDate = DataSheetOld.Range("J2").Value + 1
'    sAno = Year(dDate)
'    sMes = Format(Month(dDate) - 1, "00") 'Yahoo usa (mês - 1) para definir o mês
'    sDia = Day(dDate)
    While StocksSheet.Cells(iStock, 1) <> ""
        sStock = StocksSheet.Cells(iStock, 1)
        sStockURL = sStock
        
Application.StatusBar = "Obtendo dados > " & sStock & " (" & iStock - 1 & "\" & nStocks & ")" & "..."

'Preço atual
    'Somente se depois das 19h
        qurl = "http://download.finance.yahoo.com/d/quotes.csv?s=" & sStock & "&f=l1jk&e=.csv"
Application.StatusBar = "Obtendo dados > " & sStock & " (" & iStock - 2 & "\" & nStocks & ")" & " > Preço Atual..."
        iStock = iStock + 1
            
QueryQuote:
            
        With DataSheet.QueryTables.Add(Connection:="URL;" & qurl, Destination:=DataSheet.Cells(nFirstRow, 7))
            .BackgroundQuery = True
            .TablesOnlyFromHTML = False
            .Refresh BackgroundQuery:=False
            .SaveData = True
        End With
                            
        nLastRow = DataSheet.Cells(nFirstRow, 7).CurrentRegion.Rows.Count + nFirstRow - 2
        
        'Index - pegar preços e volume
        If StocksSheet.Cells(iStock - 1, 2).Value = "Index" Then
'            'Volume - Volume aparece sempre zero em preços atuais. Preencher somente em Preços Históricos, no dia seguinte.
'            Call NHNL1_CopyCells_TodaysData(DataSheet, sStock, nLastRow, nFirstRow, "_1V")
'            'Open
'            Call NHNL1_CopyCells_TodaysData(DataSheet, sStock, nLastRow, nFirstRow, "_2O")
'            'High
'            Call NHNL1_CopyCells_TodaysData(DataSheet, sStock, nLastRow, nFirstRow, "_3H")
'            'Low
'            Call NHNL1_CopyCells_TodaysData(DataSheet, sStock, nLastRow, nFirstRow, "_4L")
'            'Close
'            Call NHNL1_CopyCells_TodaysData(DataSheet, sStock, nLastRow, nFirstRow, "_5C")
            nFirstRow = nLastRow + 3
        Else
            Call NHNL1_FillCells_TodaysData(DataSheet, sStock, nFirstRow, nLastRow, 1)
            'Formar colunas
            Call NHNL1_ToColumns_NHNLNow(DataSheet, sStock, nLastRow, nFirstRow, "")
            nFirstRow = nLastRow + 3
        End If

NextStock:
    Wend
        
Application.StatusBar = "Organizando dados..."
        
    DataSheet.Range("A1").Select
    DataSheet.Range("A1") = "Stock"
    DataSheet.Range("B1").Select
    DataSheet.Range("B1") = "Date"
    DataSheet.Range("C1").Select
    DataSheet.Range("C1") = "Time"
    DataSheet.Range("D1").Select
    DataSheet.Range("D1") = "Close"
    DataSheet.Range("E1").Select
    DataSheet.Range("E1") = "52Low"
    DataSheet.Range("F1").Select
    DataSheet.Range("F1") = "52High"
    
Application.StatusBar = "Organizando dados > Apagando Query Tables" & "..."
    
    For Each QueryTable In DataSheet.QueryTables
        a = QueryTable.Name
        QueryTable.Delete
    Next
    
    Columns("G:G").Select
    Selection.ClearContents
    
Application.StatusBar = "Organizando dados > Ordenando dados" & "..."
    
    'Ordenar
    Columns("A:F").Select
    DataSheet.Sort.SortFields.Clear
    DataSheet.Sort.SortFields.Add Key:=Range("A2:A1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    DataSheet.Sort.SortFields.Add Key:=Range("B2:B1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With DataSheet.Sort
        .SetRange Range("A1:F1000000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'    ActiveWorkbook.Save
    
Application.StatusBar = "Organizando dados > Atualizando tabela dinâmica" & "..."
    
    Cells.Select
    Cells.Calculate
    
    ActiveWorkbook.RefreshAll
    
    'Estética
    DataSheet.Cells.EntireColumn.AutoFit
    
    'Turn calculation back on
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
   
Application.StatusBar = False ' gives control of the statusbar back to the programme
Application.ScreenUpdating = True
    
Exit Sub

'Apagar cotações em feriados (repetem dia anterior)
    'Falha do yahoo. Algumas ações

    '"H1" informa primeira linha com dados repetidos (atualizar)
    Cells.Select
    Cells.Calculate
'    Application.Calculation = xlCalculationAutomatic
'    Application.Calculation = xlCalculationManual

Application.StatusBar = "Organizando dados > Classificando por dados válidos" & "..."

    'Descobrir e apagar células inválidas (feriados)
    Columns("A:I").Select
    DataSheet.Sort.SortFields.Clear
    DataSheet.Sort.SortFields.Add Key:=Range("I2:I1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    DataSheet.Sort.SortFields.Add Key:=Range("B2:B1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With DataSheet.Sort
        .SetRange Range("A1:I1000000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
Application.StatusBar = "Organizando dados > Calculando e apagando dados inválidos" & "..."
    
    'Atualizar célula "H1"
    Cells.Select
    Cells.Calculate
'    Application.Calculation = xlCalculationAutomatic
'    Application.Calculation = xlCalculationManual
    
    sPrimeiraCelulaInvalida = Range("H1").Value
   
    DataSheet.Activate
    Range("A" & sPrimeiraCelulaInvalida & ":C10000").Select
    Selection.ClearContents
   
Application.StatusBar = "Organizando dados > Classificando dados válidos" & "..."
   
    'Ordenar
    Columns("A:C").Select
    DataSheet.Sort.SortFields.Clear
    DataSheet.Sort.SortFields.Add Key:=Range("A2:A1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    DataSheet.Sort.SortFields.Add Key:=Range("B2:B1000000" _
        ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With DataSheet.Sort
        .SetRange Range("A1:C1000000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
Application.StatusBar = "Finalizando" & "..."
   
    'Enquanto não capturar lastdate do banco de dados
        'Capturar também valid dates utilizando a resposta de um dos índices - tabela dinâmica
    DataSheet.Range("J2").Value = DataSheet.Range("B2").Value - 2
   
   
Exit Sub        ' Saia para evitar manipulador.
ErrorHandler:    ' Rotina de tratamento de erro.
    Select Case Err.Number    ' Avalie o número do erro.
        Case 1004    ' Arquivo não existe no site do Yahoo
            MsgBox ("Erro " & Err.Number & ": " & Err.Description)
            GoTo NextStock     ' Feche o arquivo aberto.
        Case Else
            MsgBox ("Erro " & Err.Number & ": " & Err.Description)
            GoTo NextStock     ' Feche o arquivo aberto.
            ' Trate outras situações aqui...
    End Select
    Resume    ' Continue a execução na mesma linha
                ' que provocou o erro.
End Sub


Function fForceDateToEnglish(sDate As String)
    sDate = Replace(sDate, " fev ", " feb ")
    sDate = Replace(sDate, " abr ", " apr ")
    sDate = Replace(sDate, " mai ", " may ")
    sDate = Replace(sDate, " ago ", " aug ")
    sDate = Replace(sDate, " set ", " sep ")
    sDate = Replace(sDate, " out ", " oct ")
    sDate = Replace(sDate, " dez ", " dec ")
    fForceDateToEnglish = sDate
End Function