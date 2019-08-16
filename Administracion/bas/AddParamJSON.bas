Attribute VB_Name = "AddParamJSON"
Public Sub AddParamJSON(sBuff As String, sNombreCampo As String, sValorCampo As String, Optional Cerrar As Boolean)
    If Len(sBuff) = 0 Then sBuff = "{"
    sBuff = sBuff & Chr$(34) & sNombreCampo & Chr$(34) & ":" & Chr$(34) & sValorCampo & Chr$(34)
    If Cerrar Then sBuff = sBuff & "}" Else sBuff = sBuff & ","
End Sub

