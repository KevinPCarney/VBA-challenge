VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub stock_ticker():

 Dim ws As Worksheet
 
 For Each ws In ThisWorkbook.Sheets
 
 
 

    
    Dim i As Double ' row number
    Dim k As Integer ' leaderboard row
        
        Dim st_open As Double ' contents of column C
        Dim first_open As Double ' contents of column C
        Dim st_high As Double ' contents of column D
        Dim last_close As Double ' contents of column D
    
    Dim ticker As String ' what is going to go in column A
    Dim st_date As Date ' what is going to go in column B
    
        Dim st_low As Double ' contents of column E
        Dim st_close As Double ' contents of column F
        Dim vol As Double ' contents of column G
    
    'Leaderboard
    
        Dim tot_stock As Double
        Dim quart_change As Double
        Dim perc_change As Double
    
    
    tot_stock = 0
    k = 2
    
    'total number of rows = 93001.... AAF = 1 - 63,  AAR = 64 - 125
    
    For i = 2 To 93001:
        vol = ws.Cells(i, 7).Value 'stock volume
        ticker = ws.Cells(i, 1).Value
        
      '----------------------------------------
      
      ' resolves columns I and L - Start
        
        If (ws.Cells(i + 1, 1).Value <> ticker) Then ' checks to see if the next i is a different ticker
            
            tot_stock = tot_stock + vol
            
            ws.Cells(k, 9).Value = ticker
            ws.Cells(k, 12).Value = tot_stock
            
            ' reset
            
        Else
            ' we just add to the total
            tot_stock = tot_stock + vol
        End If
        
        'resolves columns I and L - end
        
        '--------------------------------------
        
    ' resolves column J & K - Start
      
      If ((ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value) And ((ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value))) Then
            
            
            first_open = Round(ws.Cells(i, 3).Value, 2)
         
         End If
            
            
        If ((ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) And ((ws.Cells(i - 1, 1).Value = ws.Cells(i, 1).Value))) Then
        
            last_close = Round(ws.Cells(i, 6).Value, 2)
            
            
            quart_change = Round(last_close - first_open, 2)
            perc_change = quart_change / first_open
            
            
            ws.Cells(k, 10).Value = quart_change ' column J

            If (ws.Cells(k, 10).Value > 0) Then ' establish the color coding based on value for row J
            ws.Cells(k, 10).Interior.ColorIndex = 4
            
        ElseIf (ws.Cells(k, 10).Value < 0) Then
            ws.Cells(k, 10).Interior.ColorIndex = 3
            
        Else
            ws.Cells(k, 10).Interior.ColorIndex = 2
            
        End If

            ws.Cells(k, 11).Value = perc_change ' column K
            
           
            
            ' reset
            tot_stock = 0
            k = k + 1
        
        End If
        
        'resolves column J & K - end
        '-----------------------------------------
        
        'resolves for column N - Start
            ' rows 2 through 1501
                
    Next i
    
 Dim j As Long ' define loop
 
  
 Dim max_per As Double
 Dim min_per As Double
 Dim max_tot As Double
   
    
max_per = Application.WorksheetFunction.Max(ws.range("K:K")) ' finds max of Percent change
   
  Dim max_ticker As String
  Dim max_tot_tick As String
  
  
min_per = Application.WorksheetFunction.Min(ws.range("K:K")) ' finds min of percent change
    
  Dim min_ticker As String
  
max_tot = Application.WorksheetFunction.Max(ws.range("L:L")) ' finds max of total stock volume
  
  ws.Cells(2, 17).Value = max_per
  ws.Cells(3, 17).Value = min_per
  ws.Cells(4, 17).Value = max_tot
  
  For j = 2 To 1501 ' defining range of loop
  
  ' checks to see if j,11 values equals the "Greatest % increase" entry
   
   If ws.Cells(j, 11).Value = ws.Cells(2, 17).Value Then
    
   max_ticker = ws.Cells(j, 9).Value
    
   ws.Cells(2, 16).Value = max_ticker ' if value matches, populate the associated ticker
    
   End If
    
   
   ' checks to see if j,11 values equals the "Greatest % decrease" entry
   
   
   If ws.Cells(j, 11).Value = ws.Cells(3, 17).Value Then
    
   min_ticker = ws.Cells(j, 9).Value
    
   ws.Cells(3, 16).Value = min_ticker
   
   End If
   
   ' checks to see if j,12 values equals the "Greatest total volume" entry
   
   If ws.Cells(j, 12).Value = ws.Cells(4, 17).Value Then
   
   max_tot_tick = ws.Cells(j, 9).Value
   
   ws.Cells(4, 16).Value = max_tot_tick
    
   End If
       
   
   Next j
   
   Next ws
  
  
  
   
    
End Sub


