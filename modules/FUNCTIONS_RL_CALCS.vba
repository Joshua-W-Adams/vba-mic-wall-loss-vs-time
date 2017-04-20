Attribute VB_Name = "FUNCTIONS_RL_CALCS"
Option Explicit

'Description: Specific sub routine call to generate a wall loss vs time graph for the parameters specific in the 'Wall_Loss_Vs_Time_Graph' worksheet
Sub generate_graph()
    
    Dim wkbk As Workbook:                           Set wkbk = ThisWorkbook
    Dim sht As Worksheet:                           Set sht = wkbk.Worksheets("Wall_Loss_Vs_Time_Graph")
    Dim clear_range As Range:                       Set clear_range = sht.Range("U2:X1000")
    Dim insert_cell As Range:                       Set insert_cell = sht.Range("U2")
    Dim acr_bands_array As Range:                   Set acr_bands_array = sht.Range(sht.Range("C7").Value)
    Dim last_inspection_date As Date:               last_inspection_date = sht.Range("C8").Value
    Dim last_inspection_date_wall_loss As Double:   last_inspection_date_wall_loss = sht.Range("C9").Value
    Dim nominal_wall_thickness As Double:           nominal_wall_thickness = sht.Range("C10").Value
    Dim minimum_allowable_wall_thickness As Double: minimum_allowable_wall_thickness = sht.Range("C15").Value
    Dim current_acr As Double:                      current_acr = sht.Range("C12").Value
    Dim actual_cr As Double:                        actual_cr = sht.Range("C13").Value
    Dim current_rl As Double:                       current_rl = sht.Range("C11").Value
    Dim current_end_of_life As Date:                current_end_of_life = sht.Range("C16").Value
    Dim recommended_rl As Double
    Dim recommended_end_of_life As Date
    Dim return_value As Variant
    
    return_value = calculate_acr_bands_data("database", last_inspection_date, last_inspection_date_wall_loss, _
    acr_bands_array, nominal_wall_thickness, minimum_allowable_wall_thickness, current_acr, actual_cr, current_rl, current_end_of_life)
    
    recommended_rl = calculate_acr_bands_data("recommended_rl", last_inspection_date, last_inspection_date_wall_loss, _
    acr_bands_array, nominal_wall_thickness, minimum_allowable_wall_thickness, current_acr, actual_cr, current_rl, current_end_of_life)
    
    recommended_end_of_life = DateValue(Format(DateAdd("d", recommended_rl * 365, Now()), "Short Date"))
    Call output_graph_data(sht, clear_range, insert_cell, return_value)
    Call configure_graph(sht, CLng(last_inspection_date), IIf(CLng(recommended_end_of_life) > CLng(current_end_of_life), CLng(recommended_end_of_life) + 100, CLng(current_end_of_life) + 100))
    
End Sub

'Description: Calculate database of wall loss vs time information so relevant parameters can be returned
Function calculate_acr_bands_data(return_parameter As String, last_inspection_date As Date, last_inspection_date_wall_loss As Double, _
    passed_acr_bands_array As Range, nominal_wall_thickness As Double, _
    minimum_allowable_wall_thickness As Double, current_acr As Double, actual_cr As Double, current_rl As Double, current_end_of_life As Date) As Variant
    
    'Create copy of data range so nominal wall thickness value is not overriden
    Dim acr_bands_array As Range:                   Set acr_bands_array = passed_acr_bands_array
    Dim n As Integer:                               n = 0
    Dim acr_bands_array_size As Integer:            acr_bands_array_size = acr_bands_array.Rows.Count
    Dim return_data_array() As BAND_ARRAY_CLASS
    Dim current_band_position As Integer
    Dim recommended_acr As Double
    Dim recommended_rl As Double
    Dim forecast_wall_loss As Double
    Dim recommended_end_of_life As Date
    
    'Update band array with fail FFS wall thickness of current cml
    acr_bands_array(acr_bands_array.Rows.Count, 1) = nominal_wall_thickness - minimum_allowable_wall_thickness
    
    'Determine current band that the last inspection date wall loss falls into
    For n = 1 To acr_bands_array_size - 1
        
        If last_inspection_date_wall_loss < acr_bands_array(n + 1, 1) Then
        
            current_band_position = n
            Exit For
        
        End If
    
    Next n
    
    'Push first row to array (last inspection date details)
    return_data_array = push_to_array(return_data_array, "Todays Date Vertical Line", Format(last_inspection_date, "Short Date"), last_inspection_date_wall_loss, acr_bands_array(current_band_position, 2))
    
    'Push all corrosion rate band data
    If n <> 0 Then 'If a hole out is not detected
    
        For n = current_band_position To acr_bands_array_size - 1
            
            return_data_array = push_to_array(return_data_array, _
                "Corrosion Rate Bands Lines", _
                Format(DateAdd("d", (acr_bands_array(n + 1, 1) - return_data_array(UBound(return_data_array)).wall_loss) / acr_bands_array(n, 2) * 365, return_data_array(UBound(return_data_array)).date_value), "Short Date"), _
                acr_bands_array(n + 1, 1), _
                acr_bands_array(n, 2))
            
        Next n
    
    End If
    
    'Push data for current wall loss as of today
    For n = 1 To UBound(return_data_array) - 1
    
        If Now() < return_data_array(n + 1).date_value Then
            
            return_data_array = push_to_array(return_data_array, _
                "Corrosion Rate Bands Lines", _
                Format(Now(), "Short Date"), _
                return_data_array(n).acr * DateDiff("d", return_data_array(n).date_value, Now()) / 365 + return_data_array(n).wall_loss, _
                return_data_array(n).acr)
                
                Dim rsize As Integer
                rsize = UBound(return_data_array)
                Debug.Print return_data_array(rsize).graph_name
                Debug.Print return_data_array(rsize).date_value
                Debug.Print return_data_array(rsize).wall_loss
                Debug.Print return_data_array(rsize).acr
            
            Exit For
        
        End If
    
    Next n
    
    recommended_acr = (return_data_array(UBound(return_data_array) - 1).wall_loss - return_data_array(UBound(return_data_array)).wall_loss) _
                        / (return_data_array(UBound(return_data_array) - 1).date_value - return_data_array(UBound(return_data_array)).date_value)
    forecast_wall_loss = return_data_array(UBound(return_data_array)).wall_loss
    recommended_end_of_life = return_data_array(UBound(return_data_array) - 1).date_value
    recommended_rl = DateDiff("d", return_data_array(UBound(return_data_array) - 1).date_value, Now()) / 365
    
    'Debug.Print recommended_rl
    
    'Add additional graph data for reference
    return_data_array = push_to_array(return_data_array, "Todays Date Vertical Line", DateValue(Format(Now(), "Short Date")), 0, 0)
    return_data_array = push_to_array(return_data_array, "Todays Date Vertical Line", DateValue(Format(Now(), "Short Date")), nominal_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Recommended RL Vertical Line", DateValue(Format(recommended_end_of_life, "Short Date")), 0, 0)
    return_data_array = push_to_array(return_data_array, "Recommended RL Vertical Line", DateValue(Format(recommended_end_of_life, "Short Date")), nominal_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Recommended ACR Line", DateValue(Format(return_data_array(UBound(return_data_array)).date_value, "Short Date")), return_data_array(UBound(return_data_array)).wall_loss, recommended_acr)
    return_data_array = push_to_array(return_data_array, "Recommended ACR Line", DateValue(Format(recommended_end_of_life, "Short Date")), nominal_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Current RL Vertical Line", current_end_of_life, 0, 0)
    return_data_array = push_to_array(return_data_array, "Current RL Vertical Line", current_end_of_life, nominal_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Current ACR Line", DateValue(Format(last_inspection_date, "Short Date")), last_inspection_date_wall_loss, current_acr)
    return_data_array = push_to_array(return_data_array, "Current ACR Line", DateValue(Format(current_end_of_life, "Short Date")), nominal_wall_thickness, current_acr)
      
    return_data_array = push_to_array(return_data_array, "Actual CR Line", DateValue(Format(last_inspection_date, "Short Date")), last_inspection_date_wall_loss, actual_cr)
    return_data_array = push_to_array(return_data_array, "Actual CR Line", DateAdd("d", actual_cr * 365, last_inspection_date), nominal_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Actual CR RL Line", DateAdd("d", actual_cr * 365, last_inspection_date), 0, 0)
    return_data_array = push_to_array(return_data_array, "Actual CR RL Line", DateAdd("d", actual_cr * 365, last_inspection_date), nominal_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Fail FFS Line", DateValue(Format(last_inspection_date, "Short Date")), nominal_wall_thickness - minimum_allowable_wall_thickness, 0)
    return_data_array = push_to_array(return_data_array, "Fail FFS Line", IIf(recommended_end_of_life > DateAdd("d", current_acr * 365, last_inspection_date), recommended_end_of_life, DateAdd("d", current_acr * 365, last_inspection_date)), nominal_wall_thickness - minimum_allowable_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Nominal Wall Thickness Line", last_inspection_date, nominal_wall_thickness, 0)
    return_data_array = push_to_array(return_data_array, "Nominal Wall Thickness Line", IIf(recommended_end_of_life > current_end_of_life, recommended_end_of_life, current_end_of_life), nominal_wall_thickness, 0)
    
    'Return requested parameter to user
    If return_parameter = "database" Then
    
        calculate_acr_bands_data = return_data_array
    
    ElseIf return_parameter = "recommended_acr" Then
    
        calculate_acr_bands_data = recommended_acr
        
    ElseIf return_parameter = "forecast_wall_loss" Then
        
        calculate_acr_bands_data = forecast_wall_loss
        
    ElseIf return_parameter = "recommended_rl" Then
    
        calculate_acr_bands_data = recommended_rl
    
    Else
    
        calculate_acr_bands_data = Null
    
    End If
    
End Function

'Description: Pass array and values to push
Function push_to_array(bands_array As Variant, graph_name As String, date_value As Date, wall_loss As Double, acr As Double) As Variant
    
    Dim size As Integer: size = UBound(bands_array) + 1
    
    ReDim Preserve bands_array(size)
    
    Set bands_array(size) = New BAND_ARRAY_CLASS
    
    With bands_array(size)
        .graph_name = graph_name
        .wall_loss = wall_loss
        .date_value = date_value
        .acr = acr
    End With

    push_to_array = bands_array
    
End Function

'Description: Outputs all data from the passed array to a specific location on a worksheet then sorts
Function output_graph_data(sht As Worksheet, clear_range As Range, insert_cell As Range, output_array As Variant)
    
    Dim x As Integer:   x = insert_cell.Row
    Dim y As Integer:   y = insert_cell.Column
    Dim i As Integer
    
    'Remove all existing data from sheet
    clear_range.Clear
    
    'Loop through array and output to location
    For i = 0 To UBound(output_array) - 1
    
        sht.Cells(x + i, y).Value = output_array(i + 1).graph_name
        sht.Cells(x + i, y + 1).Value = output_array(i + 1).date_value
        sht.Cells(x + i, y + 2).Value = output_array(i + 1).wall_loss
        sht.Cells(x + i, y + 3).Value = output_array(i + 1).acr
    
    Next i
    
    'Sort data by name then by wall_loss
    Range(Cells(x, y), Cells(x + UBound(output_array) - 1, y + 3)).Sort _
        key1:=Range(Cells(x, y), Cells(x + UBound(output_array) - 1, y)), _
        key2:=Range(Cells(x, y + 2), Cells(x + UBound(output_array) - 1, y + 2)), _
        order1:=xlAscending, Header:=xlNo

End Function

'Description: Corrects data ranges in chart to correct locations as per the output dataset
Function configure_graph(sht As Worksheet, x_axis_min As Long, x_axis_max As Long)

    With sht.ChartObjects(1).Chart.Axes(xlCategory)
        .MinimumScale = CDbl(x_axis_min)
        .MaximumScale = CDbl(x_axis_max)
    End With

End Function
