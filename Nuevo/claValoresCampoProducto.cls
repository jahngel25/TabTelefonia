VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claValoresCampoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public proProductNumber As String
Public proCampo As String
Public proValorId As String
Public proValorDesc As String
Public proValorIdPadre As String
Public proConexion As ADODB.Connection
Public Function MetExistenRelaciones() As Boolean

Dim varResultados As ADODB.Recordset
Dim varValordatos As claValordatos
Dim varPosicion As Integer
On Error GoTo ErrorManager

        MetExistenRelaciones = False
        'Trae los datos de la base
        Set varResultados = New ADODB.Recordset

        varResultados.Open "select isnull(count(*),0) as cuantos " _
                            & " from         ct_detalledatosproducto inner join " _
                            & " ct_datosproducto on ct_detalledatosproducto.idatosproductoid = ct_datosproducto.idatosproductoid " _
                            & " group by ct_datosproducto.chproductnumber,ct_detalledatosproducto." & Trim(Me.proCampo) _
                            & " having      (ct_datosproducto.chproductnumber = '" & Me.proProductNumber & "') and (ct_detalledatosproducto." & Trim(Me.proCampo) & " <>'' AND ct_detalledatosproducto." & Trim(Me.proCampo) & " is not null) ", Me.proConexion
        If Not varResultados.EOF Then
            If varResultados.Fields("cuantos") > 0 Then
                MetExistenRelaciones = True
            End If
        End If
    varResultados.Close
    Exit Function
ErrorManager:
    SubGMuestraError
End Function

Public Function MetInsertar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    varComando = "INSERT INTO CT_ValoresCampoXProducto (    chProductNumber,    " & _
                                                " chCampo,            " & _
                                                "  iValorId ,iValorIdPadre   )       " & _
                                                "VALUES ('" & Me.proProductNumber & "', '" & _
                                                Me.proCampo & "', " & _
                                                Me.proValorId & "," & IIf(Me.proValorIdPadre = "" Or Me.proValorIdPadre = "0", Me.proValorId, Me.proValorIdPadre) & ")"
    
    Me.proConexion.Execute varComando
    
    MetInsertar = True
    
    Exit Function
ErrManager:
    SubGMuestraError
    Resume
End Function

Public Function MetEliminar() As Boolean
    Dim varComando As String
    On Error GoTo ErrManager
    
    varComando = "DELETE FROM CT_ValoresCampoXProducto  " & _
                 "WHERE chProductNumber = '" & Me.proProductNumber & "' " & _
                 "AND   chCampo         = '" & Me.proCampo & "' " & _
                 "AND   iValorId        =  " & Me.proValorId & _
                 "AND   iValorIdPadre      =  " & Me.proValorIdPadre
    Me.proConexion.Execute varComando
    
    MetEliminar = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function MetValidarExistencia()
    Dim varComando As String
    Dim varResultado As ADODB.Recordset
    On Error GoTo ErrManager
    
    varComando = " Select    1                           " & _
                 " From      CT_ValoresCampoxproducto    " & _
                 " Where     chProductNumber = '" & Me.proProductNumber & "' " & _
                 " And       chCampo         = '" & Me.proCampo & "' " & _
                 " And       iValorId        =  " & Me.proValorId & _
                 " And       iValorIdPadre        =  " & Me.proValorIdPadre
                 
    Set varResultado = New ADODB.Recordset
    
    varResultado.Open varComando, Me.proConexion
    
    If Not varResultado.EOF Then
        If IsNull(varResultado.Fields(0)) Then
            MetValidarExistencia = False
        Else
            If Trim(varResultado.Fields(0)) = 1 Then
                MetValidarExistencia = True
            Else
                MetValidarExistencia = False
            End If
        End If
    End If
    
    Set varResultado = Nothing

    Exit Function
ErrManager:
    SubGMuestraError
End Function
