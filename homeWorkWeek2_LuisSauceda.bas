Attribute VB_Name = "Module1"
Option Explicit
Public Sub Stock_Market_Analysis_Easy()
'variables
Dim lRow As Long ' guarda la ultima fila con valor en la hoja
Dim lCol As Integer ' guarda la ultima columne con valor en la hoja
Dim StockRange As Range ' guarda el rango con valores
Dim StockRow As Object 'variable para navegar por las filas del rango
Dim StockTotalVol As Currency

Dim stockValues As Object
Dim thiker As String
Dim volume As Long
Dim key As Variant
Dim cont As Integer


'*******************************************************************


Dim W_Sheet As Worksheet

For Each W_Sheet In ActiveWorkbook.Worksheets

    'Set W_Sheet = Workbooks("alphabtical_testing").Worksheets("A")
    
    W_Sheet.Activate
        
    lRow = Cells.Find(What:="*", _
                        After:=Range("A1"), _
                        LookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
                        
    lCol = Cells.Find(What:="*", _
                        After:=Range("A1"), _
                        LookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
                        
    
    Set stockValues = CreateObject("Scripting.Dictionary")
    
    
    With W_Sheet
        Set StockRange = .Range(.Cells(2, 1), .Cells(lRow, lCol))
        'Debug.Print StockRange.Cells(1, 1)
    
        For Each StockRow In StockRange.Rows
       
            thiker = StockRow.Cells(1, 1)
            If IsNumeric(StockRow.Cells(1, lCol)) Then
                volume = StockRow.Cells(1, lCol)
            End If
            
            If Not stockValues.exists(thiker) Then
                stockValues.Add thiker, volume
            Else
                StockTotalVol = stockValues(thiker) + volume
                stockValues(thiker) = StockTotalVol
                StockTotalVol = 0
            End If
        Next
    End With
    
    cont = 2
    
    Cells(cont, 10).Activate
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Total Stock Volume"
    For Each key In stockValues.keys
        Cells(cont, 10).Value = key
        Cells(cont, 11).Value = stockValues(key)
        cont = cont + 1
    'Debug.Print key, stockValues(key)
    Next
    Range("K:K").NumberFormat = "General"
Next

End Sub
Public Sub Stock_Market_Analysis_Hard()
'variables
Dim lRow As Long ' guarda la ultima fila con valor en la hoja
Dim lCol As Integer ' guarda la ultima columne con valor en la hoja
Dim StockRange As Range ' guarda el rango con valores
Dim StockRow As Object 'variable para navegar por las filas del rango
Dim StockTotalVol As Currency
Dim StockFactors As String
Dim StockOpenPrice As String
Dim StockClosePrice As String
Dim StockDate As String
Dim yearlyChange As Double
Dim percentageChange As Double
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As Currency
Dim greatestIncThiker As String
Dim greatestDecThiker As String
Dim greatestVolumeThiker As String
'Dim stockOb As StockObject

Dim stockValues As Object
Dim thiker As String
Dim key As Variant
Dim cont As Integer


'*******************************************************************


Dim W_Sheet As Worksheet

For Each W_Sheet In ActiveWorkbook.Worksheets

    'Set W_Sheet = Workbooks("alphabtical_testing").Worksheets("A")
    
    W_Sheet.Activate
        
    lRow = Cells.Find(What:="*", _
                        After:=Range("A1"), _
                        LookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Row
                        
    lCol = Cells.Find(What:="*", _
                        After:=Range("A1"), _
                        LookAt:=xlPart, _
                        LookIn:=xlFormulas, _
                        SearchOrder:=xlByColumns, _
                        SearchDirection:=xlPrevious, _
                        MatchCase:=False).Column
                        
    
    Set stockValues = CreateObject("Scripting.Dictionary")

    
    
    With W_Sheet
        Set StockRange = .Range(.Cells(2, 1), .Cells(lRow, lCol))
        StockRange.Sort key1:=Range("B1"), Order1:=xlAscending ', Header:=xlYes
        
    
        For Each StockRow In StockRange.Rows
            
            thiker = StockRow.Cells(1, 1)
                        
            If IsNumeric(StockRow.Cells(1, lCol)) Then
                StockTotalVol = StockRow.Cells(1, lCol)
            End If
             
           
            If Not stockValues.exists(thiker) Then
                
                'build a string with all the stock values we need
                'totalVolume | open price | close price | date
                StockFactors = CStr(StockRow.Cells(1, lCol)) & "|" & CStr(StockRow.Cells(1, 3)) & "|" & CStr(StockRow.Cells(1, 6)) & "|" & CStr(StockRow.Cells(1, 2))
                stockValues.Add thiker, StockFactors
                StockTotalVol = 0
                StockFactors = ""
            Else
                'guardamos el volumen total
'                Dim text As String
'                text = StockFactorsFunction(stockValues(thiker), 1)
'                If IsNumeric(text) Then
'                    StockTotalVol = CCur(text)
'                End If
'If thiker = "A" Then
'Debug.Print thiker & "|" & stockValues(thiker)
'End If
                StockTotalVol = CCur(StockFactorsFunction(stockValues(thiker), 1)) + StockTotalVol
                
                'validamos la fecha para el precio de cierre
                
                If CLng(StockFactorsFunction(stockValues(thiker), 4)) < CLng(StockRow.Cells(1, 2).Value) Then
                    StockClosePrice = CStr(StockRow.Cells(1, 6).Value)
                    'StockDate = StockFactorsFunction(stockValues(thiker), 4)
                    StockDate = CStr(StockRow.Cells(1, 2).Value)
                End If
                'armamos el string de factores con los nuevos datos
                StockOpenPrice = StockFactorsFunction(stockValues(thiker), 2)
                
                'asignamos el string de factores al thiker
                stockValues(thiker) = StockTotalVol & "|" & StockOpenPrice & "|" & StockClosePrice & "|" & StockDate
                StockTotalVol = 0
                StockClosePrice = ""
                'StockDate = ""

                'StockTotalVol = stockValues(thiker) + volume
                'stockValues(thiker) = StockTotalVol
                'StockTotalVol = 0
            End If
        Next
    End With
    
    cont = 2
    
    Cells(cont, 10).Activate
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percentage Change"
    Range("M1").Value = "Total Stock Volume"
    
    greatestIncrease = -9999999
    greatestDecrease = 9999999
    greatestVolume = 0
    greatestIncThiker = ""
    greatestDecThiker = ""
    greatestVolumeThiker = ""

    For Each key In stockValues.keys
        
    
        Cells(cont, 10).Value = key
        StockOpenPrice = StockFactorsFunction(stockValues(key), 2)
        StockClosePrice = StockFactorsFunction(stockValues(key), 3)
        yearlyChange = CDbl(StockClosePrice) - CDbl(StockOpenPrice)
        If StockOpenPrice <> 0 Then
            percentageChange = (CDbl(StockClosePrice) / CDbl(StockOpenPrice)) - 1
        Else
            percentageChange = 0
        End If
        If percentageChange > greatestIncrease Then
            greatestIncrease = percentageChange
            greatestIncThiker = key
        End If
        If percentageChange < greatestDecrease Then
            greatestDecrease = percentageChange
            greatestDecThiker = key
        End If
        If greatestVolume < CCur(StockFactorsFunction(stockValues(key), 1)) Then
            greatestVolume = CCur(StockFactorsFunction(stockValues(key), 1))
            greatestVolumeThiker = key
        End If
        
        Cells(cont, 11).Value = yearlyChange 'stockValues(key)(2)
        If yearlyChange > 0 Then
            Cells(cont, 11).Interior.ColorIndex = 4
        ElseIf yearlyChange < 0 Then
            Cells(cont, 11).Interior.ColorIndex = 3
        End If
        Cells(cont, 12).Value = percentageChange
        Cells(cont, 13).Value = StockFactorsFunction(stockValues(key), 1)
        
        cont = cont + 1
        yearlyChange = 0
        percentageChange = 0
    'Debug.Print key, stockValues(key)
    Next
    'Hard version
    Range("P1").Value = "Thiker"
    Range("Q1").Value = "Value"

    Range("K:K").NumberFormat = "General"
    Range("L:L").NumberFormat = "0.00%"
    
    Range("O2").Value = "Greatest % decrease"
    Range("P2").Value = greatestDecThiker
    Range("Q2").Value = greatestDecrease
    Range("Q2").NumberFormat = "0.00%"
    
    Range("O3").Value = "Greatest % Increase"
    Range("P3").Value = greatestIncThiker
    Range("Q3").Value = greatestIncrease
    Range("Q3").NumberFormat = "0.00%"
    
    Range("O4").Value = "Greatest Total Volume"
    Range("P4").Value = greatestVolumeThiker
    Range("Q4").Value = greatestVolume
    
Next

End Sub

Public Function StockFactorsFunction(factor As String, position As Integer) As String


    Dim factors() As String
    'Dim StockFactors As String
    factors() = Split(factor, "|")
    If position = 1 Then ' regresa volumen
        StockFactorsFunction = factors(0)
    ElseIf position = 2 Then 'regresa open price
        StockFactorsFunction = factors(1)
    ElseIf position = 3 Then 'regresa open price
        StockFactorsFunction = factors(2)
    ElseIf position = 4 Then 'regresa la fecha
        StockFactorsFunction = factors(3)
    End If

    'Return
End Function

