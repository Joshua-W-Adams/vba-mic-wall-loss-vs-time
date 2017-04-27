VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BAND_ARRAY_CLASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'AUTHOR: Joshua William Adams
'REV HISTORY:
'REV: A DESC.: Issued for Review                    DATE: 27/04/2017
'REV: 0 DESC.: Issued for Use                       DATE: 27/04/2017
'DESCRIPTION: Class to store all MIC wall loss vs time graph data
    
    Public graph_name As String
    Public date_value As Date
    Public wall_loss As Double
    Public acr As Double
