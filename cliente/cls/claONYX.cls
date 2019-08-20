VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claONYX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public ContactName As String
Public ContactID As String
Public AlternateID As String
Public DetailID As String
Public UserLogin As String
Public UserPassword As String
Public UserSite As String
Public ServerName As String
Public DatabaseName As String

Public proConexion As ADODB.Connection
Private Function FunFIniciaConexion() As Boolean
'*****************************************************************
' OBJETIVO: Abrir la conexión con el servidor
'******************************************************************
' AUTOR: Raúl Cruz
' Fecha: 26/12/2000
'******************************************************************
On Error GoTo ErrorManager

    'Inicia la Conexión
    Set proConexion = New ADODB.Connection
    Dim ApplicationName As String
        
    ApplicationName = "Tab de Telefonia Vers. 3.7.3"
    
    'Cadena de conexión
    proConexion.ConnectionString = "driver={SQL Server};" & _
      "server=" & ServerName & ";" & _
      "uid=" & UserLogin & ";" & _
      "pwd=" & UserPassword & ";" & _
      "database=" & DatabaseName & ";" & _
      "App= " & ApplicationName
      
    'Abrir la conexión
    proConexion.ConnectionTimeout = 30
    proConexion.Open
    FunFIniciaConexion = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function
Public Sub Delete()
    '-------------------------------------------------------------------------
    ' Funcionalidad de Eliminación
    '-------------------------------------------------------------------------
    'Inicio de la Conexión
    If FunFIniciaConexion = False Then
        MsgBox "No fue posible conectarse al servidor especificado.", vbInformation, App.Title
        Exit Sub
    End If
    
    'Set frmEdificios.proConexion = Me.proConexion
    'frmEdificios.proEdificioId = Me.DetailID
    'frmEdificios.proCompanyId = Me.ContactID
    'frmEdificios.proCompanyName = Me.ContactName
   
    'frmSplash.Show vbModal
End Sub

Public Sub Initiate()
    '-------------------------------------------------------------------------
    ' Place your initiate (edit/create) functionality here.
    '-------------------------------------------------------------------------
    
    'Inicio de la conexión
    If FunFIniciaConexion = False Then
        MsgBox "No fue posible conectarse al servidor especificado.", vbInformation, App.Title
        Exit Sub
    End If
   
    
    'MsgBox "       Initiate" & Chr(13) & Chr(13) & _
                "Detail:                  " & Me.DetailID & Chr(13) & _
                "Contact Id:            " & Me.ContactID & Chr(13) & _
                "Contact Name:       " & Me.ContactName & Chr(13) & _
                "Alternate Id:          " & Me.AlternateID & Chr(13) & _
                "Database Name:     " & Me.DatabaseName & Chr(13) & _
                "Server Name:         " & Me.ServerName & Chr(13) & _
                "User Login:             " & Me.UserLogin & Chr(13) & _
                "User Password:       " & Me.UserPassword & Chr(13) & _
                "User Site:               " & Me.UserSite
                
    Set frmVoz.proConexion = Me.proConexion
    
    frmVoz.proDatosProductoId = Me.DetailID
    frmVoz.proCompanyId = Me.ContactID
    frmVoz.proCompanyName = Me.ContactName
    
    Set frmVoz.proOnyx = Me
    
    'indica que la forma que empieza es Selección de la facturacion
    frmSplash.Show vbModal
End Sub

Public Sub Load()
    '-------------------------------------------------------------------------
    ' Place your load (view) functionality here.
    '-------------------------------------------------------------------------
    'Inicio de la conexión
    If FunFIniciaConexion = False Then
        MsgBox "No fue posible conectarse al servidor especificado.", vbInformation, App.Title
        Exit Sub
    End If
    
    'MsgBox "       Load" & Chr(13) & Chr(13) & _
                "Detail:                  " & Me.DetailID & Chr(13) & _
                "Contact Id:            " & Me.ContactID & Chr(13) & _
                "Contact Name:       " & Me.ContactName & Chr(13) & _
                "Alternate Id:          " & Me.AlternateID & Chr(13) & _
                "Database Name:     " & Me.DatabaseName & Chr(13) & _
                "Server Name:         " & Me.ServerName & Chr(13) & _
                "User Login:             " & Me.UserLogin & Chr(13) & _
                "User Password:       " & Me.UserPassword & Chr(13) & _
                "User Site:               " & Me.UserSite
    
    
    Set frmVoz.proConexion = Me.proConexion
    
    frmVoz.proDatosProductoId = Me.DetailID
    frmVoz.proCompanyId = Me.ContactID
    frmVoz.proCompanyName = Me.ContactName
    
    Set frmVoz.proOnyx = Me
            
    'indica que la forma que empieza es Selección de la facturacion
    frmSplash.Show vbModal
End Sub

Public Sub Terminate()
    Unload frmVoz
    Me.proConexion.Close
End Sub

