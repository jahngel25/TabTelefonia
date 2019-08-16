VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claProductMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proProductNumber As String
Public proSiteId As String
Public proAlternateLangId As String
Public proSecondaryId As String
Public proHierarchyId As String
Public proDescription As String
Public proUser1 As String
Public proUser2 As String
Public proUser3 As String
Public proUser4 As String
Public proUser5 As String
Public proUser6 As String
Public proUser7 As String
Public proUser8 As String
Public proUser9 As String
Public proUser10 As String
Public proStockLevel As String
Public proListPrice As String
Public proCost As String
Public proSupportType As String
Public proWarrantyCalls As String
Public proWarrantyDays As String
Public proWarrantyExpDate As String
Public proIncidentProductNumber As String
Public proOrderable As String
Public dtFirstShipDate As String
Public proObsoleteDate As String
Public proStatusId As String
Public proInsertBy As String
Public proInsertDate As String
Public proUpdateBy As String
Public proUpdateDate As String
Public proRecordStatus As String
Public proModifiedDate As String

'Conexion
Public proConexion As ADODB.Connection
