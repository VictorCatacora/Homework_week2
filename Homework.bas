Attribute VB_Name = "Module1"
Sub Homework3()
    
    Dim ticker_arreglo() As String
    Dim contador As Integer
    Dim temporal As String
    Dim init, vol As Double
    Dim valor1, valor2, valor3, valor4, pre_volumen As Double
    Dim ticker1, ticker2, pre_volumen_ticke As String

    
    contador = 0
    temporal = ""
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastrow
        If Cells(i, 1).Value <> temporal Then
            temporal = Cells(i, 1)
            contador = contador + 1
        End If
    Next i
    
    ReDim ticker_arreglo(contador)
    
    temporal = ""
    contador = 0
    For i = 2 To lastrow
        If Cells(i, 1).Value <> temporal Then
            temporal = Cells(i, 1).Value
            ticker_arreglo(contador) = temporal
            contador = contador + 1
        End If
    Next i
    
    
    lastcol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Definiendo la primera fila
    Cells(1, lastcol + 2).Value = "Ticker"
    Cells(1, lastcol + 3).Value = "Yearly Change"
    Columns(lastcol + 3).AutoFit
    Cells(1, lastcol + 4).Value = "Percent Change"
    Columns(lastcol + 4).AutoFit
    Cells(1, lastcol + 5).Value = "Total Stock Volume"
    Columns(lastcol + 5).AutoFit
    
    valor3 = 0
   valor4 = 0
    pre_volumen = 0
    

    init = 2
    For i = 0 To (contador - 1)
        vol = 0
        valor1 = 0
        For j = init To lastrow
            If Cells(j, 1).Value = ticker_arreglo(i) Then
                vol = vol + Cells(j, lastcol)
                If valor1 = 0 Then
                    valor1 = Cells(j, 3).Value
                End If
                valor2 = Cells(j, 6).Value
            Else
                init = j
                Exit For
            End If
        Next j
        Cells(i + 2, lastcol + 2).Value = ticker_arreglo(i)
        Cells(i + 2, lastcol + 3).Value = valor2 - valor1
        Cells(i + 2, lastcol + 3).NumberFormat = "0.###############"
        If valor1 <> 0 Then
            Cells(i + 2, lastcol + 4).Value = FormatPercent((Cells(i + 2, lastcol + 3).Value / valor1), 2)
        Else
            Cells(i + 2, lastcol + 4).Value = FormatPercent((Cells(i + 2, lastcol + 3).Value), 2)
        End If
        If Cells(i + 2, lastcol + 3).Value < 0 Then
            Cells(i + 2, lastcol + 3).Interior.Color = vbRed
            If Cells(i + 2, lastcol + 4) < valor3 Then
                valor3 = Cells(i + 2, lastcol + 4)
                ticker1 = ticker_arreglo(i)
            End If
        Else
            Cells(i + 2, lastcol + 3).Interior.Color = vbGreen
            If Cells(i + 2, lastcol + 4) > valor4 Then
               valor4 = Cells(i + 2, lastcol + 4)
                ticker2 = ticker_arreglo(i)
            End If
        End If
        Cells(i + 2, lastcol + 5).Value = vol
        If vol > pre_volumen Then
            pre_volumen = vol
            pre_volumen_ticke = ticker_arreglo(i)
        End If
    Next i
    
    'Insertando el ultimo recuadro
    lastcol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    Cells(1, lastcol + 4).Value = "Ticker"
    Cells(1, lastcol + 5).Value = "Value"
    
    Cells(2, lastcol + 3).Value = "Greatest % Increase"
    Cells(3, lastcol + 3).Value = "Greatest % Decrease"
    Cells(4, lastcol + 3).Value = "Greatest Total Volume"
    Columns(lastcol + 3).AutoFit
    
    Cells(2, lastcol + 4).Value = ticker2
    Cells(2, lastcol + 5).Value = FormatPercent(prev_pos_val, 2)
    Cells(3, lastcol + 4).Value = ticker1
    Cells(3, lastcol + 5).Value = FormatPercent(valor3, 2)
    Cells(4, lastcol + 4).Value = pre_volumen_ticke
    Cells(4, lastcol + 5).Value = pre_volumen
    Columns(lastcol + 5).AutoFit




End Sub

