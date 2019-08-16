VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claRegla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public proReglaId As String
Public proDescripcion As String
Public proCantidadDigitos As String
Public proRepeticiones As String
Public proPosicionDigitos As String
Public proConsecutivoDigitos As String
Public proRecordStatus As String

Public proConexion As adodb.Connection

Public Function FunGGuardar() As Boolean
'***********************************************************************
' OBJETIVO: Toma la decisión de Insertar o Modificar
'***********************************************************************
'¨PARAMETROS:
'       Ninguno.
'************************************************************************
' AUTOR: Raúl Cruz
' FECHA: 02/03/2001
'************************************************************************
On Error GoTo ErrorManager

    If Len(Trim(Me.proReglaId)) = 0 Then
        FunGGuardar = Me.FunGInsertar
    Else
        FunGGuardar = Me.FunGModificar
    End If
    If FunGGuardar = True Then
       'genera los numeros de la regla si es una modificacion borra los numero generadoa
       ' y los vuelve a generar
        Dim varCommand As adodb.Command
        Set varCommand = New adodb.Command
        varCommand.CommandType = adCmdStoredProc
        varCommand.ActiveConnection = Me.proConexion
            
        varCommand.CommandText = "CP_GenerarReglaConsulta"
        varCommand.Parameters(1) = Me.proReglaId
        'Espera máximo 10 minutos
        varCommand.CommandTimeout = 6000
        varCommand.Execute
    End If
    
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGInsertar() As Boolean

Dim varResultados As adodb.Recordset
On Error GoTo ErrorManager

     Me.proRecordStatus = 1
    'Ejecución de Insert
    Me.proConexion.Execute " INSERT INTO CT_Reglas (" & _
      " vchDescripcionRegla,iCantidadDigitos, " & _
      " iRepeticiones,chPosicionDigitos, " & _
      " chConsecutivoDigitos, tiRecordStatus)" & _
      "VALUES (" & _
     " '" & Me.proDescripcion & "', '" & Me.proCantidadDigitos & "', " & _
     " '" & Me.proRepeticiones & "', '" & Me.proPosicionDigitos & "', " & _
     " '" & Me.proConsecutivoDigitos & "','" & Me.proRecordStatus & "'  )"
    
    If Me.proConexion.Errors.Count = 0 Then 'Si no hubo error al ingresar los datos
        'Consulta del código que acaba de insertar
        Set varResultados = New adodb.Recordset
        varResultados.Open "SELECT proReglaId = max(iReglaId) FROM CT_Reglas", Me.proConexion
                       
        If Not (varResultados.EOF) Then
                Me.proReglaId = varResultados("proReglaId").Value
        End If
    End If

    FunGInsertar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

Public Function FunGModificar() As Boolean

On Error GoTo ErrorManager

    'Ejecución de update
    Me.proConexion.Execute "UPDATE CT_Reglas set " & _
     " vchDescripcionRegla ='" & Me.proDescripcion & "'," & _
     " iCantidadDigitos ='" & Me.proCantidadDigitos & "', " & _
     " iRepeticiones= '" & Me.proRepeticiones & "', " & _
     " chPosicionDigitos= '" & Me.proPosicionDigitos & "', " & _
     " chConsecutivoDigitos ='" & Me.proConsecutivoDigitos & "'," & _
     " tiRecordStatus= '" & Me.proRecordStatus & "' " & _
     "WHERE iReglaId = " & Me.proReglaId
    
    If Me.proConexion.Errors.Count Then Exit Function
    FunGModificar = True
    Exit Function
    
ErrorManager:
    SubGMuestraError
End Function

