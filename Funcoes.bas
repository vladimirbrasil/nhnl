Attribute VB_Name = "Funcoes"
Sub PasteUnformatted()
Attribute PasteUnformatted.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' PasteUnformatted Macro
' Pastes unformatted text
'
' Keyboard Shortcut: Ctrl+q
'
    Application.CutCopyMode = False
    ActiveSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:=False
End Sub

Function Extenso(ByVal numero) ''Escreve numero por extenso
    Dim Reais, Centavos, Temp
    Dim PontoDecimal, Contar
    Dim strExtenso As String
    ReDim lugar(9) As String
    Dim dNum As Single
    lugar(2) = " Mil "
    lugar(3) = " Milhões "
    lugar(4) = " Bilhões"
    lugar(5) = " Trilhões"
    
    dNum = numero
    dNum = Round(dNum + 0.0000001, 2)
    numero = Trim(Str(dNum))
    
    ''Posição da casa decimal se 0 numero inteiro
    PontoDecimal = InStr(numero, ".")
    ''Converter centavos
    If PontoDecimal > 0 Then
        Centavos = GetDez(Left(Mid(numero, PontoDecimal + 1) & "00", 2))
        numero = Trim(Left(numero, PontoDecimal - 1))
    End If
    Contar = 1
    Do While numero <> ""
        Temp = GetCem(Right(numero, 3))
        If Temp <> "" Then Reais = Temp & lugar(Contar) & Reais
            If Len(numero) > 3 Then
                numero = Left(numero, Len(numero) - 3)
        Else
          numero = ""
        End If
        Contar = Contar + 1
    Loop
    Select Case Reais
        Case ""
            Reais = ""
        Case " Um"
            Reais = " Um Real"
        Case Else
            Reais = Reais & " Reais"
        End Select
    Select Case Centavos
        Case ""
            Centavos = ""
        Case " Um"
            Centavos = "Um centavo"
        Case Else
            Centavos = Centavos & " Centavos"
    End Select
    If Reais <> "" And Centavos <> "" Then
        strExtenso = Reais & " e " & Centavos
    ElseIf Reais <> "" Then
        strExtenso = Reais
    Else
        strExtenso = Centavos
    End If
    strExtenso = Replace$(strExtenso, "  ", " ")
    strExtenso = Replace$(strExtenso, " e Centavos", " Centavos")
    strExtenso = Replace$(strExtenso, "Mil ", "Mil, ")
    strExtenso = Replace$(strExtenso, "Mil, Reais", "Mil Reais")
    strExtenso = Replace$(strExtenso, "os e Reais", "os Reais")
    strExtenso = Replace$(strExtenso, " e Reais", " Reais")
    strExtenso = Replace$(strExtenso, "Mil, ", "Mil e ")
    Extenso = UCase$(Trim$(strExtenso))

End Function
    
'' Converter um numero entre 100 e 999 em texto
Function GetCem(ByVal numero)
    Dim resultado As String
    If Val(numero) = 0 Then Exit Function
    numero = Right("000" & numero, 3)
    If Mid(numero, 1, 1) <> "0" Then
        resultado = GetDigit(Mid(numero, 1, 1)) '' ALTERAR ESTÁ FUNÇÃO SE 1=CEM ; 2 = DUZENTOS
        Select Case resultado
            Case " Um": resultado = " Cento "
            Case " Dois": resultado = " Duzentos "
            Case " Três": resultado = " Trezentos "
            Case " Quatro": resultado = " Quatrocentos "
            Case " Cinco": resultado = " Quinhentos "
            Case " Seis": resultado = " Seiscentos "
            Case " Sete": resultado = " Setecentos "
            Case " Oito": resultado = " Oitocentos "
            Case " Nove": resultado = " Novecentos "
        End Select
        
    
    End If
    '' Converte um numero entre 01 e 10 em texto
    If Mid(numero, 2, 1) <> "0" Then
        If resultado = "" Then
            resultado = resultado & GetDez(Mid(numero, 2))
        Else
            resultado = resultado & " e " & GetDez(Mid(numero, 2))
        End If
    Else
        If resultado = "" Then
            resultado = resultado & GetDigit(Mid(numero, 3))
        Else
            resultado = resultado & " e " & GetDigit(Mid(numero, 3))
        End If
    End If
GetCem = resultado
End Function
    
'' Converte um numero de 10 a 99 em texto
Function GetDez(DezTXT)
    Dim result As String
    result = "" ''Nulo
    If Val(Left(DezTXT, 1)) = 1 Then ''Se valor entre 10-19
        Select Case Val(DezTXT)
            Case 10: result = "Dez"
            Case 11: result = "Onze"
            Case 12: result = "Doze"
            Case 13: result = "Treze"
            Case 14: result = "Quatorze"
            Case 15: result = "Quinze"
            Case 16: result = "Dezesseis"
            Case 17: result = "Dezessete"
            Case 18: result = "Dezoito"
            Case 19: result = "Dezenove"
            Case Else
        End Select
    Else '' Valores entre 20-99
        Select Case Val(Left(DezTXT, 1))
            Case 2: result = " Vinte"
            Case 3: result = " Trinta"
            Case 4: result = " Quarenta"
            Case 5: result = " Cinquenta"
            Case 6: result = " Sessenta"
            Case 7: result = " Setenta"
            Case 8: result = " Oitenta"
            Case 9: result = " Noventa"
            Case Else
        End Select
        If result = "" Then
            result = result & GetDigit(Right(DezTXT, 1)) '' retorna um unico valor
        Else
            result = result & " e " & GetDigit(Right(DezTXT, 1)) '' retorna um unico valor
        End If
    End If
    GetDez = result
End Function

''Converte numeros entre 1 e 9 em texto
Function GetDigit(Digit)
    Select Case Val(Digit)
        Case 1: GetDigit = " Um"
        Case 2: GetDigit = " Dois"
        Case 3: GetDigit = " Três"
        Case 4: GetDigit = " Quatro"
        Case 5: GetDigit = " Cinco"
        Case 6: GetDigit = " Seis"
        Case 7: GetDigit = " Sete"
        Case 8: GetDigit = " Oito"
        Case 9: GetDigit = " Nove"
        Case Else: GetDigit = ""
    End Select
End Function
