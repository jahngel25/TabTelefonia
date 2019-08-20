VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claIncidentType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colIncidentType"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public proParentId As String
Public proParameterId  As String
Public provchParameterDesc As String
Public proConexion As ADODB.Connection
Private mvarcolIncidentType As colIncidentType
Private Sub Class_Terminate()
    Set mvarcolIncidentType = Nothing
End Sub

