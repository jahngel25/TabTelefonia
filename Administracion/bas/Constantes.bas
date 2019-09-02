Attribute VB_Name = "Constantes"
Option Explicit

Global Const AplicacionID = 20
Global Const ConstDolares = 2

Global varGAdminAplicacion As Boolean
Global varGAdminTelefonia As Boolean
Global UserName As String
Global IncidentId As String
Global ClientId As String
Global TramaRequest As String
Global TramaResponse As String
Global EventoLog As String
Global TipoLog As String
Global CodigoLog As String
Global NombreMaquinaLog As String

Global varResultados As ADODB.Recordset
varResultados = New ADODB.Recordset
Script = "SELECT vchMetododAtributo " & _
         "FROM AtributosSoapWebService " & _
         "WHERE vchMetodo = 'NetCracker'"
varResultados.Open Script, Me.proConexion





