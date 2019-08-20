VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************
'       DESCRIPCION: Clase con id y descripcion de las categorias, venta
'       atencion u ot
'       MODIFICADO POR:       TOPGROUP S.A.
'       DESCRIPCION CAMBIO:   Desarrollo inicial de la clase
'       VERSION:       1.0.000
'       REQUERIMIENTO: 3488
'       FECHA:       2009/08/05
'*******************************************************************
Option Explicit

Public proCategoriaID As String
Public proDescripcion As String

'Conexion
Public proConexion As ADODB.Connection

