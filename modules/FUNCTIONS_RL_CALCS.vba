Attribute VB_Name = "FUNCTIONS_RL_CALCS"
Option Explicit

'Description: Specific sub routine call to generate a wall loss vs time graph for the parameters specific in the 'Wall_Loss_Vs_Time_Graph' worksheet
Sub generate_graph()
    
    Dim wkbk As Workbook:                           Set wkbk = ThisWorkbook
    Dim sht As Worksheet:                           Set sht = wkbk.Worksheets("Wall_Loss_Vs_Time_Graph")
    
    Dim last_inspection_date As Date:               last_inspection_date = sht.Range("I5").Value
    Dim last_inspection_date_wall_loss As Double:   last_inspection_date_wall_loss = sht.Range("I6").Value
    Dim acr_bands_array As Range:                   Set acr_bands_array = sht.Range(sht.Range("I7").Value)
    Dim nominal_wall_thickness As Double:           nominal_wall_thickness = sht.Range("I8").Value
    
    Call calculate_acr_bands_data(last_inspection_date, last_inspection_date_wall_loss, acr_bands_array, nominal_wall_thickness)
    
    Call clear_graph_data
    
    Call output_graph_data
    
    Call configure_graph
    
End Sub

Function calculate_acr_bands_data(last_inspection_date As Date, last_inspection_date_wall_loss As Double, _
    acr_array As Range, nominal_wall_thickness As Double, return_parameter As String) As Variant
    
    Dim n As Integer: n = 0
    Dim i As Integer: i = 0
    Dim acr_bands_array As Range: Set acr_bands_array = acr_array
    Dim graph_coordinates() As Class1
    Dim current_band_position As Integer
    
    'Update band array with nominal wall thickness of current cml
    acr_bands_array(acr_bands_array.Rows.Count, 1) = nominal_wall_thickness
    
    'Determine current band
    For n = 1 To acr_bands_array.Rows.Count - 1
        
        If last_inspection_date_wall_loss < acr_bands_array(n + 1, 1) Then
        
            current_band_position = n
            Exit For
        
        End If
    
    Next n
    
    i = i + 1
    graph_coordinates = push_to_bands_array(graph_coordinates, last_inspection_date_wall_loss, Format(last_inspection_date, "Short Date"), acr_bands_array(n, 2))
    
    'For n = 1 To i
    
        'Debug.Print "wall_loss: " & graph_coordinates(i).wall_loss
        'Debug.Print "date_value: " & graph_coordinates(i).date_value
        'Debug.Print "acr: " & graph_coordinates(i).acr
    
    'Next n
    
    'Plot all corrosion rate band points
    'Still remaining wall thickness in current CML
    If n <> 0 Then
    
        For n = current_band_position To acr_bands_array.Rows.Count - 1
            
            i = i + 1
            
            'Debug.Print graph_coordinates(i - 1).date_value
            'Debug.Print acr_bands_array(n + 1, 1)
            
            graph_coordinates = push_to_bands_array(graph_coordinates, _
                acr_bands_array(n + 1, 1), _
                Format(DateAdd("d", (acr_bands_array(n + 1, 1) - graph_coordinates(i - 1).wall_loss) / acr_bands_array(n, 2) * 365, graph_coordinates(i - 1).date_value), "Short Date"), _
                acr_bands_array(n, 2))
            
        Next n
    
    End If
    
    'Plot point for current wall loss
    For n = 1 To i - 1
    
        If Now() < graph_coordinates(n + 1).date_value Then
        
            i = i + 1
            'Debug.Print DateDiff("yyyy", graph_coordinates(n).date_value, Now())
            
            graph_coordinates = push_to_bands_array(graph_coordinates, _
                graph_coordinates(n).acr * DateDiff("d", graph_coordinates(n).date_value, Now()) / 365 + graph_coordinates(n).wall_loss, _
                Format(Now(), "Short Date"), _
                0)
            
            Exit For
        
        End If
    
    Next n
    
    'Return requested parameter to user
    If return_parameter = "database" Then
    
        calculate_acr_bands_data = graph_coordinates
    
    ElseIf return_parameter = "recommended_acr" Then
    
        calculate_acr_bands_data = graph_coordinates
        
    ElseIf return_parameter = "forecast_wall_loss" Then
        
        calculate_acr_bands_data = graph_coordinates
        
    ElseIf return_parameter = "recommended_rl" Then
    
        calculate_acr_bands_data = graph_coordinates
    
    Else
    
        calculate_acr_bands_data = Null
    
    End
    
End Function

'Pass array and values to push
Function push_to_bands_array(bands_array As Variant, wall_loss As Double, date_value As Date, acr As Double) As Variant
    
    Dim size As Integer: size = UBound(bands_array) + 1
    
    ReDim Preserve graph_coordinates(size)
    
    Set graph_coordinates(size) = New Class1
    
    'Plot point 1
    With graph_coordinates(size)
        .wall_loss = wall_loss
        .date_value = date_value
        .acr = acr
    End With

    push_to_bands_array = bands_array
    
End Function

Function clear_graph_data()



End Function

Function output_graph_data()
    
    Dim wkbk As Workbook: Set wkbk = ThisWorkbook
    Dim sht As Worksheet: Set sht = wkbk.Worksheets("Wall_Loss_Bands")
    
    Dim x As Integer
    Dim y As Integer
    
    x = 1 'spreadsheet data insert anchor point for x axis
    y = 1 'spreadsheet data insert anchor point for y axis
    
'Ouput All Data to Sheet
    'Todays date
    Worksheets("wall_loss_bands").Cells(x + 14, y).Value = DateValue(Format(Now(), "Short Date"))
    Worksheets("wall_loss_bands").Cells(x + 15, y).Value = DateValue(Format(Now(), "Short Date"))
    Worksheets("wall_loss_bands").Cells(x + 14, y + 1).Value = 0
    Worksheets("wall_loss_bands").Cells(x + 15, y + 1).Value = nominal_wall_thickness
    
    'Remaining Life
    Worksheets("wall_loss_bands").Cells(x + 19, y).Value = DateValue(Format(graph_coordinates(i - 1).date_value, "Short Date"))
    Dim RL_Recommended_Acr As Date: RL_Recommended_Acr = DateValue(Format(graph_coordinates(i - 1).date_value, "Short Date"))
    Worksheets("wall_loss_bands").Cells(x + 20, y).Value = RL_Recommended_Acr
    Worksheets("wall_loss_bands").Cells(x + 19, y + 1).Value = 0
    Worksheets("wall_loss_bands").Cells(x + 20, y + 1).Value = nominal_wall_thickness
    
    'Recommended ACR
    Worksheets("wall_loss_bands").Cells(x + 29, y).Value = DateValue(Format(graph_coordinates(i).date_value, "Short Date"))
    Worksheets("wall_loss_bands").Cells(x + 30, y).Value = RL_Recommended_Acr
    Worksheets("wall_loss_bands").Cells(x + 29, y + 1).Value = graph_coordinates(i).wall_loss
    Worksheets("wall_loss_bands").Cells(x + 29, y + 2).Value = (nominal_wall_thickness - graph_coordinates(i).wall_loss) / (DateDiff("d", graph_coordinates(i).date_value, graph_coordinates(i - 1).date_value) / 365)
    Worksheets("wall_loss_bands").Cells(x + 30, y + 1).Value = nominal_wall_thickness
    
    'Current Remaining Life
    Worksheets("wall_loss_bands").Cells(x + 34, y).Value = DateAdd("d", sht.Cells(5, 15) * 365, last_inspection_date)
    Dim RL_Current_Acr As Date: RL_Current_Acr = DateAdd("d", sht.Cells(5, 15) * 365, last_inspection_date)
    Worksheets("wall_loss_bands").Cells(x + 35, y).Value = RL_Current_Acr
    Worksheets("wall_loss_bands").Cells(x + 34, y + 1).Value = 0
    Worksheets("wall_loss_bands").Cells(x + 35, y + 1).Value = nominal_wall_thickness
    
    'Current ACR
    'HOLD ACR VALUE TO BE ADDED
    Worksheets("wall_loss_bands").Cells(x + 39, y).Value = DateValue(Format(last_inspection_date, "Short Date"))
    Worksheets("wall_loss_bands").Cells(x + 40, y).Value = DateAdd("d", sht.Cells(5, 15) * 365, last_inspection_date)
    Worksheets("wall_loss_bands").Cells(x + 39, y + 1).Value = last_inspection_date_wall_loss
    Worksheets("wall_loss_bands").Cells(x + 40, y + 1).Value = nominal_wall_thickness
    
    'Actual Corrosion Rate
    
    'HOLD ACR VALUE TO BE ADDED
    Worksheets("wall_loss_bands").Cells(x + 44, y).Value = DateValue(Format(last_inspection_date, "Short Date"))
    Worksheets("wall_loss_bands").Cells(x + 45, y).Value = DateAdd("d", sht.Cells(8, 15) * 365, last_inspection_date)
    Worksheets("wall_loss_bands").Cells(x + 44, y + 1).Value = last_inspection_date_wall_loss
    Worksheets("wall_loss_bands").Cells(x + 45, y + 1).Value = nominal_wall_thickness
    
    'Actual Corrosion Rate Remaining Life
    Worksheets("wall_loss_bands").Cells(x + 49, y).Value = DateAdd("d", sht.Cells(8, 15), last_inspection_date)
    Worksheets("wall_loss_bands").Cells(x + 50, y).Value = DateAdd("d", sht.Cells(8, 15), last_inspection_date)
    Worksheets("wall_loss_bands").Cells(x + 49, y + 1).Value = nominal_wall_thickness
    Worksheets("wall_loss_bands").Cells(x + 50, y + 1).Value = nominal_wall_thickness
    
    'Maximum Allowable Wall Loss
    Worksheets("wall_loss_bands").Cells(x + 54, y).Value = DateValue(Format(last_inspection_date, "Short Date"))
    Worksheets("wall_loss_bands").Cells(x + 55, y).Value = IIf(RL_Recommended_Acr > RL_Current_Acr, RL_Recommended_Acr, RL_Current_Acr)
    Worksheets("wall_loss_bands").Cells(x + 54, y + 1).Value = nominal_wall_thickness - Cells(9, 15).Value
    Worksheets("wall_loss_bands").Cells(x + 55, y + 1).Value = nominal_wall_thickness - Cells(9, 15).Value
    
    'Nominal Wall Thickness
    Worksheets("wall_loss_bands").Cells(x + 24, y).Value = DateValue(Format(graph_coordinates(1).date_value, "Short Date"))
    Worksheets("wall_loss_bands").Cells(x + 25, y).Value = IIf(RL_Recommended_Acr > RL_Current_Acr, RL_Recommended_Acr, RL_Current_Acr)
    Worksheets("wall_loss_bands").Cells(x + 24, y + 1).Value = nominal_wall_thickness
    Worksheets("wall_loss_bands").Cells(x + 25, y + 1).Value = nominal_wall_thickness
    
    'Actual Corrosion Rate Remaining Life
    'To be created
    
    'Plotted Data
    For n = 1 To i
        
        Worksheets("wall_loss_bands").Cells(x + 57, y) = "date_value"
        Worksheets("wall_loss_bands").Cells(x + 57, y + 1) = "wall_loss"
        Worksheets("wall_loss_bands").Cells(x + 57, y + 2) = "acr"
        
        Worksheets("wall_loss_bands").Cells(x + 57 + n, y) = graph_coordinates(n).date_value
        Worksheets("wall_loss_bands").Cells(x + 57 + n, y + 1) = graph_coordinates(n).wall_loss
        Worksheets("wall_loss_bands").Cells(x + 57 + n, y + 2) = graph_coordinates(n).acr
    
    Next n
    
    'Sort Plotted Data by wall loss
    Range(Cells(x + 57 + 1, y), Cells(x + 57 + i, y + 2)).Sort key1:=Range(Cells(x + 57 + 1, y + 1), Cells(x + 57 + i, y + 1)), _
        order1:=xlAscending, Header:=xlNo

End Function

Function configure_graph()

    With sht.ChartObjects(1).Chart.Axes(xlCategory)
        .MinimumScale = CDbl(last_inspection_date)
        .MaximumScale = CDbl(IIf(RL_Recommended_Acr > RL_Current_Acr, RL_Recommended_Acr + 100, RL_Current_Acr + 100))
    End With

End Function
