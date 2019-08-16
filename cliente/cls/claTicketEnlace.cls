VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claTicketEnlace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proiIncidentId As Long
Public provchDesc1 As String
Public prochAssignedTo As String
Public provchParameterDesc As String
Public prodtInsertDate As Date
Public provchSerialNumber As String
Public proiIncidentTypeId As Long
Public proiIncidentCategory As Long
Public proiStatusId As Long
Public prosEstado As String
Private TieneAsociaciones As Long
'Conexion
Public proConexion As ADODB.Connection
Public Property Get proTieneAsociaciones(proiIncidentId As Long) As Boolean
Dim varResultados As ADODB.Recordset
Set varResultados = New ADODB.Recordset
On Error GoTo ErrManager
varResultados.Open " SELECT     ISNULL(COUNT(iIncidentId), 0) AS Cuantos " _
                                & " From Ct_TicketxDetalleDatosProducto " _
                                & " Where (iIncidentId = " & proiIncidentId & ") ", Me.proConexion
If Not varResultados.EOF Then
    proTieneAsociaciones = (varResultados.Fields("Cuantos") > 0)
Else
     proTieneAsociaciones = 0
End If
    Exit Property
ErrManager:
    SubGMuestraError
End Property

Private Sub Class_Initialize()
    On Error GoTo ErrorManager
    Exit Sub
ErrorManager:
    SubGMuestraError
End Sub
