Attribute VB_Name = "ROUTINES_MIC_WL_VS_TIME"
'AUTHOR: Joshua William Adams
'REV HISTORY:
'REV: A DESC.: Issued for Review                    DATE: 27/04/2017
'REV: 0 DESC.: Issued for Use                       DATE: 27/04/2017
'DESCRIPTION: Module for containing all code to generate a Microbial Induced Corrosion (MIC) wall loss vs time graph.
Option Explicit

'DESCRIPTION: Master routine for generating graph.
Sub generate_mic_wall_loss_vs_time_graph()
    
    'General parameters declarations and constructors
    Dim wkbk As Workbook:                               Set wkbk = ThisWorkbook
    Dim sht As Worksheet:                               Set sht = wkbk.Worksheets("MIC_Graph")
    Dim clear_range As Range:                           Set clear_range = sht.Range("U2:X1000")
    Dim insert_cell As Range:                           Set insert_cell = sht.Range("U2")
    Dim xmin As Long
    Dim xmax As Long
    
    'Declaring and defining array of MIC ACR band data
    Dim acr_bands_array_text As String:                 acr_bands_array_text = sht.Range("C7").Value
    
    'Declaring and defining CML parameters
    Dim last_inspection_date As Date:                   last_inspection_date = sht.Range("C8").Value
    Dim last_inspection_date_wall_loss As Double:       last_inspection_date_wall_loss = sht.Range("C9").Value
    Dim nominal_wall_thickness As Double:               nominal_wall_thickness = sht.Range("C10").Value
    Dim minimum_allowable_wall_thickness As Double:     minimum_allowable_wall_thickness = sht.Range("C15").Value
    Dim current_acr As Double:                          current_acr = sht.Range("C12").Value
    Dim actual_cr As Double:                            actual_cr = sht.Range("C13").Value
    Dim current_rl As Double:                           current_rl = sht.Range("C11").Value
    Dim current_end_of_life As Date:                    current_end_of_life = sht.Range("C16").Value
    Dim actual_cr_rl As Double:                         actual_cr_rl = sht.Range("C14").Value
    
    'Declaring and defining parameters to store MIC ACR calculation outputs
    Dim return_value As Variant
    Dim recommended_rl As Double
    Dim recommended_end_of_life As Date
    Dim recommended_acr As Double
    
    'Unprotect sheet to allow editing
    sht.Unprotect ("Rhino1234")
    
    'Calling MIC ACR calculation function and storing return values
    return_value = FUNCTIONS_MIC_WL_VS_TIME.calculate_acr_bands_data("database", last_inspection_date, last_inspection_date_wall_loss, _
                        acr_bands_array_text, nominal_wall_thickness, minimum_allowable_wall_thickness, current_acr, actual_cr, current_rl, current_end_of_life, actual_cr_rl)
    recommended_rl = FUNCTIONS_MIC_WL_VS_TIME.calculate_acr_bands_data("recommended_rl", last_inspection_date, last_inspection_date_wall_loss, _
                        acr_bands_array_text, nominal_wall_thickness, minimum_allowable_wall_thickness, current_acr, actual_cr, current_rl, current_end_of_life, actual_cr_rl)
    recommended_acr = FUNCTIONS_MIC_WL_VS_TIME.calculate_acr_bands_data("recommended_acr", last_inspection_date, last_inspection_date_wall_loss, _
                        acr_bands_array_text, nominal_wall_thickness, minimum_allowable_wall_thickness, current_acr, actual_cr, current_rl, current_end_of_life, actual_cr_rl)
    recommended_end_of_life = FUNCTIONS_MIC_WL_VS_TIME.calculate_acr_bands_data("recommended_end_of_life", last_inspection_date, last_inspection_date_wall_loss, _
                        acr_bands_array_text, nominal_wall_thickness, minimum_allowable_wall_thickness, current_acr, actual_cr, current_rl, current_end_of_life, actual_cr_rl)

    'Populate outputs table
    sht.Range("C" & 38).Value = current_acr
    sht.Range("C" & 39).Value = actual_cr
    sht.Range("C" & 40).Value = recommended_acr
    sht.Range("C" & 41).Value = current_rl
    sht.Range("C" & 42).Value = actual_cr_rl
    sht.Range("C" & 43).Value = recommended_rl
    
    'Output graph data to work sheet
    Call output_graph_data(sht, clear_range, insert_cell, return_value)
    
    'Defining minimum and maximum values for chart x-axis
    xmin = CLng(last_inspection_date)
    xmax = IIf(CLng(recommended_end_of_life) > CLng(current_end_of_life), CLng(recommended_end_of_life) + 500, CLng(current_end_of_life) + 500)
    
    'Create graph using all output graph data and passed parameters
    Call configure_graph(sht, "U1:X1000", xmin, xmax)
    
    'Protect sheet
    sht.Protect ("Rhino1234")
    
End Sub

'DESCRIPTION: Outputs all data from the passed array to a specific location on a worksheet then sorts
Function output_graph_data(sht As Worksheet, clear_range As Range, insert_cell As Range, output_array As Variant)
    
    Dim x As Integer:   x = insert_cell.Row
    Dim y As Integer:   y = insert_cell.Column
    Dim i As Integer
    
    'Remove all existing data from sheet
    clear_range.Clear
    
    'Loop through array and output to location
    For i = 0 To UBound(output_array)
    
        sht.Cells(x + i, y).Value = output_array(i).graph_name
        sht.Cells(x + i, y + 1).Value = output_array(i).date_value
        sht.Cells(x + i, y + 2).Value = output_array(i).wall_loss
        sht.Cells(x + i, y + 3).Value = output_array(i).acr
    
    Next i
    
    'Sort data by name then by wall_loss
    Range(Cells(x, y), Cells(x + UBound(output_array), y + 3)).Sort _
        key1:=Range(Cells(x, y), Cells(x + UBound(output_array), y)), _
        key2:=Range(Cells(x, y + 2), Cells(x + UBound(output_array), y + 2)), _
        order1:=xlAscending, Header:=xlNo

End Function

'DESCRIPTION: Creates MIC wall loss vs time graph
Function configure_graph(sht As Worksheet, search_range As String, x_axis_min As Long, x_axis_max As Long)
    
    Dim n As Integer
    Dim xrng As Range
    Dim yrng As Range
    Dim first_row As Integer
    Dim last_row As Integer
    Dim ChartObj As ChartObject
    Dim chrt As Chart
    Dim oSeries As Series
    Dim p As Point
    Dim i As Integer
    Dim vx As Variant
    Dim vy As Variant
    
    'Delete existing chart if exists
    On Error Resume Next
    If sht.ChartObjects(1).Chart Is Nothing Then
    
    Else
    
        sht.ChartObjects(1).Delete
    
    End If
    On Error GoTo 0
    
    'Create new chart
    With sht
    
        Set ChartObj = .ChartObjects.Add(Left:=Round(sht.Range("F1").Left, 0), Top:=sht.Range("E2").Top, Width:=600, Height:=600)
        Set chrt = ChartObj.Chart
        
        With chrt
            
            'Generic chart settings
            .ChartType = xlXYScatterLinesNoMarkers
            .ChartStyle = 242
            .HasTitle = True
            .ChartTitle.Text = "Microbial Induced Corrosion (MIC)"
            
            'Adding series to chart
            .SeriesCollection.NewSeries.Name = "Actual CR"
            .SeriesCollection.NewSeries.Name = "Actual RL"
            .SeriesCollection.NewSeries.Name = "Band Data Points"
            .SeriesCollection.NewSeries.Name = "Current ACR"
            .SeriesCollection.NewSeries.Name = "Current RL"
            .SeriesCollection.NewSeries.Name = "Fail FFS"
            .SeriesCollection.NewSeries.Name = "Nominal Wt"
            .SeriesCollection.NewSeries.Name = "Recommended ACR"
            .SeriesCollection.NewSeries.Name = "Recommended RL"
            .SeriesCollection.NewSeries.Name = "Today"
            
            'Configure x-axis
            With .Axes(xlCategory)
                .TickLabels.Orientation = xlTickLabelOrientationUpward
                .HasTitle = True
                .AxisTitle.Characters.Text = "Date"
                .MinimumScale = CDbl(x_axis_min)
                .MaximumScale = CDbl(x_axis_max)
            End With
            
            'Configure y-axis
            With .Axes(xlValue)
                .HasTitle = True
                .AxisTitle.Characters.Text = "Wall Loss (mm)"
            End With
            
            'Change size of plotted graph
            .PlotArea.Select
            
            With Selection
                .Width = 600
                .Height = 480
                .Left = 25
                .Top = 20
            End With
        
        End With
    
    End With
    
    'Loop through all data series, set data ranges and set general configuration
    For n = 1 To chrt.SeriesCollection.Count
        
        Set oSeries = chrt.SeriesCollection(n)
        
        'Find location of series data
        first_row = sht.Range(search_range).Find(oSeries.Name, searchorder:=xlByRows, SearchDirection:=xlNext).Row
        last_row = sht.Range(search_range).Find(oSeries.Name, searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
        
        'Define data ranges
        Set xrng = sht.Range("V" & first_row & ":V" & last_row)
        Set yrng = sht.Range("W" & first_row & ":W" & last_row)
        
        'Sets new range for each series
        oSeries.XValues = xrng
        oSeries.Values = yrng
        
        'Define Formatting for Each Series
        If oSeries.Name = "Actual CR" Then
        
            oSeries.Format.Line.ForeColor.RGB = RGB(51, 153, 255)
            oSeries.Format.Line.DashStyle = msoLineDash
        
        ElseIf oSeries.Name = "Actual RL" Then
        
            oSeries.Format.Line.ForeColor.RGB = RGB(51, 153, 255)
            oSeries.Format.Line.DashStyle = msoLineSysDot
            
        ElseIf oSeries.Name = "Band Data Points" Then
        
            oSeries.Format.Line.ForeColor.RGB = RGB(0, 0, 255)
            oSeries.MarkerStyle = xlMarkerStyleCircle
            
            vx = oSeries.XValues
            vy = oSeries.Values
                
            For i = 1 To oSeries.Points.Count
                
                Set p = oSeries.Points(i)
                
                p.HasDataLabel = True
                
                With p.DataLabel
                
                    .Text = CDate(vx(i)) & "," & Round(vy(i), 2)
                    
                End With
                
                p.MarkerStyle = xlMarkerStyleCircle
                p.MarkerSize = 8
                
            Next i
            
        ElseIf oSeries.Name = "Current ACR" Then
        
            oSeries.Format.Line.ForeColor.RGB = RGB(153, 153, 255)
            oSeries.Format.Line.DashStyle = msoLineDash
        
        ElseIf oSeries.Name = "Current RL" Then
        
            oSeries.Format.Line.ForeColor.RGB = RGB(153, 153, 255)
            oSeries.Format.Line.DashStyle = msoLineSysDot
        
        ElseIf oSeries.Name = "Fail FFS" Then
        
            oSeries.Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        
        ElseIf oSeries.Name = "Nominal Wt" Then
            
            oSeries.Format.Line.ForeColor.RGB = RGB(255, 153, 153)
        
        ElseIf oSeries.Name = "Recommended ACR" Then
        
            oSeries.Format.Line.ForeColor.RGB = RGB(255, 178, 102)
            oSeries.Format.Line.DashStyle = msoLineDash
            
        ElseIf oSeries.Name = "Recommended RL" Then
        
            oSeries.Format.Line.ForeColor.RGB = RGB(0, 0, 255)
            oSeries.Format.Line.DashStyle = msoLineSysDot

        ElseIf oSeries.Name = "Today" Then
        
            oSeries.Format.Line.ForeColor.RGB = RGB(0, 255, 0)
        
        End If
        
    Next n
    
End Function
