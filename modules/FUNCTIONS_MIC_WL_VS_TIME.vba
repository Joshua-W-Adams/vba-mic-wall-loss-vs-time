Attribute VB_Name = "FUNCTIONS_MIC_WL_VS_TIME"
'AUTHOR: Joshua William Adams
'REV HISTORY:
'REV: A DESC.: Issued for Review                    DATE: 27/04/2017
'REV: 0 DESC.: Issued for Use                       DATE: 27/04/2017
'DESCRIPTION: Module to store MIC wall loss vs time iterative linear equation calculation and supporting functions.
Option Explicit

'DESCRIPTION: MIC wall loss vs time iterative linear equation calculation.
Function calculate_acr_bands_data(return_parameter As String, last_inspection_date As Date, last_inspection_date_wall_loss As Double, _
    acr_bands_array_text As String, nominal_wall_thickness As Double, minimum_allowable_wall_thickness As Double, _
    current_acr As Double, actual_cr As Double, current_rl As Double, current_end_of_life As Date, actual_cr_rl As Double) As Variant
    
    'General parameter definitions
    Dim rng As Range:                               Set rng = ThisWorkbook.Worksheets("MIC_Graph").Range(acr_bands_array_text)
    Dim acr_bands_array_size As Integer:            acr_bands_array_size = rng.Rows.Count
    Dim current_band_position As Variant:           current_band_position = Null
    Dim n As Integer
    Dim acr_bands_array() As BAND_ARRAY_CLASS
    
    'Defining return parameters for function
    Dim return_data_array() As BAND_ARRAY_CLASS
    Dim recommended_acr As Double
    Dim recommended_rl As Double
    Dim forecast_wall_loss As Double
    Dim recommended_end_of_life As Date
    
    'Store range reference data in virtual array to allow modification
    For n = 1 To acr_bands_array_size
        
        If n = acr_bands_array_size Then
        
            acr_bands_array = push_to_array(acr_bands_array, "acr bands dataset", Now(), nominal_wall_thickness - minimum_allowable_wall_thickness, rng(n, 2).Value)
        
        Else
        
            acr_bands_array = push_to_array(acr_bands_array, "acr bands dataset", Now(), rng(n, 1).Value, rng(n, 2).Value)
    
        End If
            
    Next n
    
    'Redefine size of array based on new virtual array data structure
    acr_bands_array_size = UBound(acr_bands_array)

    'Determine current band that the last inspection date wall loss falls into
    For n = 0 To acr_bands_array_size - 1
        
        If last_inspection_date_wall_loss < acr_bands_array(n + 1).wall_loss Then
        
            current_band_position = n
            Exit For
        
        End If
    
    Next n
    
    'Push first point to array and accoutn for hole out case (no band position found)
    If Not IsNull(current_band_position) Then
    
        return_data_array = push_to_array(return_data_array, "Band Data Points", Format(last_inspection_date, "Short Date"), last_inspection_date_wall_loss, acr_bands_array(current_band_position).acr)
    
    Else
    
        return_data_array = push_to_array(return_data_array, "Band Data Points", Format(last_inspection_date, "Short Date"), last_inspection_date_wall_loss, 0)
    
    End If
    
    'Push all corrosion rate band data
    If Not IsNull(current_band_position) Then 'If a hole out is not detected
    
        For n = current_band_position To acr_bands_array_size - 1
            
            return_data_array = push_to_array(return_data_array, _
                "Band Data Points", _
                Format(DateAdd("d", (acr_bands_array(n + 1).wall_loss - return_data_array(UBound(return_data_array)).wall_loss) / acr_bands_array(n).acr * 365, return_data_array(UBound(return_data_array)).date_value), "Short Date"), _
                acr_bands_array(n + 1).wall_loss, _
                acr_bands_array(n + 1).acr)
            
        Next n
    
    End If
    
    'Push data for current wall loss as of today
    For n = 0 To UBound(return_data_array) - 1
    
        If Now() < return_data_array(n + 1).date_value Then
            
            return_data_array = push_to_array(return_data_array, _
                "Band Data Points", _
                Format(Now(), "Short Date"), _
                return_data_array(n).acr * DateDiff("d", return_data_array(n).date_value, Now()) / 365 + return_data_array(n).wall_loss, _
                return_data_array(n).acr)
                
            Exit For
        
        End If
    
    Next n
    
    forecast_wall_loss = return_data_array(UBound(return_data_array)).wall_loss
    
    'Hole (100% WL) case
    If Not IsNull(current_band_position) Then
        
        recommended_acr = (return_data_array(UBound(return_data_array) - 1).wall_loss - return_data_array(UBound(return_data_array)).wall_loss) _
                            / ((return_data_array(UBound(return_data_array) - 1).date_value - return_data_array(UBound(return_data_array)).date_value) / 365)
                            
        recommended_end_of_life = return_data_array(UBound(return_data_array) - 1).date_value
        
        recommended_rl = DateDiff("d", Now(), recommended_end_of_life) / 365
        
    Else
    
        recommended_acr = 0
        recommended_end_of_life = Now()
        recommended_rl = 0
        
    End If
    
    'Add additional graph data for reference
    return_data_array = push_to_array(return_data_array, "Today", DateValue(Format(Now(), "Short Date")), 0, 0)
    return_data_array = push_to_array(return_data_array, "Today", DateValue(Format(Now(), "Short Date")), nominal_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Recommended RL", DateValue(Format(recommended_end_of_life, "Short Date")), 0, 0)
    return_data_array = push_to_array(return_data_array, "Recommended RL", DateValue(Format(recommended_end_of_life, "Short Date")), nominal_wall_thickness - minimum_allowable_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Recommended ACR", DateValue(Format(Now(), "Short Date")), forecast_wall_loss, recommended_acr)
    return_data_array = push_to_array(return_data_array, "Recommended ACR", DateValue(Format(recommended_end_of_life, "Short Date")), nominal_wall_thickness - minimum_allowable_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Current RL", current_end_of_life, 0, 0)
    return_data_array = push_to_array(return_data_array, "Current RL", current_end_of_life, nominal_wall_thickness - minimum_allowable_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Current ACR", DateValue(Format(last_inspection_date, "Short Date")), last_inspection_date_wall_loss, current_acr)
    return_data_array = push_to_array(return_data_array, "Current ACR", DateValue(Format(current_end_of_life, "Short Date")), nominal_wall_thickness - minimum_allowable_wall_thickness, current_acr)
      
    return_data_array = push_to_array(return_data_array, "Actual CR", DateValue(Format(last_inspection_date, "Short Date")), last_inspection_date_wall_loss, actual_cr)
    return_data_array = push_to_array(return_data_array, "Actual CR", DateAdd("d", actual_cr_rl * 365, last_inspection_date), nominal_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Actual RL", DateAdd("d", actual_cr_rl * 365, last_inspection_date), 0, 0)
    return_data_array = push_to_array(return_data_array, "Actual RL", DateAdd("d", actual_cr_rl * 365, last_inspection_date), nominal_wall_thickness - minimum_allowable_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Fail FFS", DateValue(Format(last_inspection_date, "Short Date")), nominal_wall_thickness - minimum_allowable_wall_thickness, 0)
    return_data_array = push_to_array(return_data_array, "Fail FFS", IIf(recommended_end_of_life > current_end_of_life, recommended_end_of_life, current_end_of_life), nominal_wall_thickness - minimum_allowable_wall_thickness, 0)
    
    return_data_array = push_to_array(return_data_array, "Nominal Wt", last_inspection_date, nominal_wall_thickness, 0)
    return_data_array = push_to_array(return_data_array, "Nominal Wt", IIf(recommended_end_of_life > current_end_of_life, recommended_end_of_life, current_end_of_life), nominal_wall_thickness, 0)
    
    'Return requested parameter to user
    If return_parameter = "database" Then
    
        calculate_acr_bands_data = return_data_array
    
    ElseIf return_parameter = "recommended_acr" Then
    
        calculate_acr_bands_data = Round(recommended_acr, 2)
        
    ElseIf return_parameter = "forecast_wall_loss" Then
        
        calculate_acr_bands_data = Round(forecast_wall_loss, 2)
        
    ElseIf return_parameter = "recommended_rl" Then
    
        calculate_acr_bands_data = Round(recommended_rl, 2)

    ElseIf return_parameter = "recommended_end_of_life" Then
        
        calculate_acr_bands_data = DateValue(Format(DateAdd("d", recommended_rl * 365, Now()), "Short Date"))
        
    Else
    
        calculate_acr_bands_data = Null
    
    End If
    
End Function

'DESCRIPTION: Pass array and values to push
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

