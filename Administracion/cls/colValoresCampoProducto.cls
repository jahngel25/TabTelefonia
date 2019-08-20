VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colValoresCampoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"claValoresCampoProducto"
Attribute VB_Ext_KEY = "Member0" ,"claValoresCampoProducto"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'***********************************************************************
'* Descripcion:
'*
'*
'*
'*
'*
'* Parametros:
'*
'*
'*
'*
'*
'*
'*
'**************************************************************************
'**********************************************************************
' MODIFICADO POR :      CARLOS ALBERTO BARRERA
' DESCRIPCION CAMBIO:   Se pasa como parametro la propiedad del id del cliente
' VERSION: 1.0.100
' FECHA: SEPTIEMBRE 7 /2009
'**********************************************************************

Option Explicit

Public proProductNumber As String
Public proCampo As String
Public proValidarRepetidos As String
Public proConexion As ADODB.Connection
Public proValorIdPadre As Long

'local variable to hold collection
Private mCol As Collection

Public Function Add(proConexion As ADODB.Connection, proValorDesc As String, _
                    proValorId As String, proValorIdPadre As String, proCampo As String, _
                    proProductNumber As String, proMinimo As Integer, proMaximo As Integer, _
                    proUsual As Integer, Optional sKey As String) As claValoresCampoProducto
    'create a new object
    Dim objNewMember As claValoresCampoProducto
    Set objNewMember = New claValoresCampoProducto


    'set the properties passed into the method
    If IsObject(proConexion) Then
        Set objNewMember.proConexion = proConexion
    Else
        objNewMember.proConexion = proConexion
    End If
    
    objNewMember.proValorDesc = proValorDesc
    objNewMember.proValorId = proValorId
    objNewMember.proValorIdPadre = proValorIdPadre
    objNewMember.proCampo = proCampo
    objNewMember.proProductNumber = proProductNumber
    objNewMember.proMinimo = proMinimo
    objNewMember.proMaximo = proMaximo
    objNewMember.proUsual = proUsual
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As claValoresCampoProducto
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


Public Function MetConsultarValoresxProducto()

   On Error GoTo ErrManager:
    Dim varResultado As New ADODB.Recordset
    Dim varCommand As New ADODB.Command
    Dim varValorCampoProducto As New claValoresCampoProducto
    
    Set mCol = Nothing
    Set mCol = New Collection
    
    'Llamar al procedimiento de consulta
    varCommand.CommandType = adCmdStoredProc
    varCommand.ActiveConnection = Me.proConexion
    varCommand.CommandText = "CP_CONSULTARVALORESPORPRODUCTO"
    varCommand.Parameters.Append varCommand.CreateParameter("@P_IVALORIDPADRE", adInteger, adParamInput, 0, proValorIdPadre)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_CHPRODUCTNUMBER", adVarChar, adParamInput, 20, proProductNumber)
    varCommand.Parameters.Append varCommand.CreateParameter("@P_CHCAMPO", adVarChar, adParamInput, 50, proCampo)
    varResultado.Open varCommand
    
    While Not varResultado.EOF
        varValorCampoProducto.proCampo = IIf(IsNull(proCampo), "", proCampo) 'Campo
        varValorCampoProducto.proValorId = IIf(IsNull(varResultado!IVALORID), "0", varResultado!IVALORID) 'Codigo
        varValorCampoProducto.proValorIdPadre = IIf(IsNull(proValorIdPadre), "0", proValorIdPadre)
        varValorCampoProducto.proValorDesc = IIf(IsNull(varResultado!VCHVALORDESC), "", varResultado!VCHVALORDESC) 'Descripcion
        varValorCampoProducto.proProductNumber = IIf(IsNull(proProductNumber), "", proProductNumber) 'Producto
        varValorCampoProducto.proMinimo = varResultado!IMINIMO
        varValorCampoProducto.proMaximo = varResultado!IMAXIMO
        varValorCampoProducto.proUsual = varResultado!IUSUAL
        Me.Add Me.proConexion, _
                varValorCampoProducto.proValorDesc, _
                varValorCampoProducto.proValorId, _
                varValorCampoProducto.proValorIdPadre, _
                varValorCampoProducto.proCampo, _
                varValorCampoProducto.proProductNumber, _
                varValorCampoProducto.proMinimo, _
                varValorCampoProducto.proMaximo, _
                varValorCampoProducto.proUsual
        varResultado.MoveNext
    Wend
    varResultado.Close
    Set varResultado = Nothing
    MetConsultarValoresxProducto = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function



'"--------------------------------------------------------------------------------------------
'"<summary> Método que consulta los valores de acuerdo con el id de un cliente
'"</summary>
'"<remark>MetConsultarxCampoProducto(Optional pariClienteId As Long = 0)<para><version>1.0.100</version><cambio>Se adiciono para pasar como parametro el id del cliente</cambio><fecha>30/06/2009</fecha></remark>
'"</remark>
'"<param name = "pariClienteId">Id del cliente</param>
'"<returns> Booleano
'"<example> <code>if (MetConsultarxCampoProducto(pariClienteId)) then </code></example>
'"<history>
'"</history>
'"----------------------------------------------------------------------------------------------

Public Function MetConsultarxCampoProducto(Optional pariClienteId As Long = 0) As Boolean '1.0.100 Se pasa como parametro el id del cliente
    Dim varComando As ADODB.Command
    Dim varResultado As ADODB.Recordset
    Dim varValorCampoProducto As claValoresCampoProducto
    On Error GoTo ErrManager
        
    Set varComando = New ADODB.Command
    Set varComando.ActiveConnection = Me.proConexion
    varComando.CommandType = adCmdStoredProc
    varComando.CommandText = "CP_ConsultarValoresCampoProducto"
    
    varComando.Parameters.Item("@chProductNumber").Value = Me.proProductNumber
    varComando.Parameters.Item("@chCampo").Value = Me.proCampo
    varComando.Parameters.Item("@cValidarRepetidos").Value = Me.proValidarRepetidos
    varComando.Parameters.Item("@iValorIdPadre").Value = Me.proValorIdPadre
    
    varComando.Parameters.Item("@iCompanyId").Value = pariClienteId '1.0.100 Se pasa como al procedimiento como parametro el ide del Cliente

    Set varResultado = New ADODB.Recordset
    
    Set varResultado = varComando.Execute
    
    Set varValorCampoProducto = New claValoresCampoProducto
    
    While Not varResultado.EOF
        'Producto
        If IsNull(varResultado.Fields("chProductNumber")) Then
            varValorCampoProducto.proProductNumber = 0
        Else
            varValorCampoProducto.proProductNumber = varResultado.Fields("chProductNumber")
        End If
        
        'Campo
        If IsNull(varResultado.Fields("chCampo")) Then
            varValorCampoProducto.proCampo = 0
        Else
            varValorCampoProducto.proCampo = varResultado.Fields("chCampo")
        End If
        
        'Codigo
        If IsNull(varResultado.Fields("iValorId")) Then
            varValorCampoProducto.proValorId = 0
        Else
            varValorCampoProducto.proValorId = varResultado.Fields("iValorId")
        End If
        
                'Codigo
        If IsNull(varResultado.Fields("iValorIdPadre")) Then
            varValorCampoProducto.proValorIdPadre = 0
        Else
            varValorCampoProducto.proValorIdPadre = varResultado.Fields("iValorIdPadre")
        End If
        
        'Descripcion
        If IsNull(varResultado.Fields("vchValorDesc")) Then
            varValorCampoProducto.proValorDesc = 0
        Else
            varValorCampoProducto.proValorDesc = varResultado.Fields("vchValorDesc")
        End If
        
        Me.Add Me.proConexion, _
                varValorCampoProducto.proValorDesc, _
                varValorCampoProducto.proValorId, _
                varValorCampoProducto.proValorIdPadre, _
                varValorCampoProducto.proCampo, _
                varValorCampoProducto.proProductNumber, _
                varValorCampoProducto.proMinimo, _
                varValorCampoProducto.proMaximo, _
                varValorCampoProducto.proUsual
                
        varResultado.MoveNext
    Wend
    
    Set varComando = Nothing
    Set varValorCampoProducto = Nothing
    Set varResultado = Nothing
    
    MetConsultarxCampoProducto = True
    Exit Function
ErrManager:
    SubGMuestraError
End Function

Public Function BuscarIndiceProValorId(varTipoLinea As String) As Long
    'Retorna el índice de el elemento con el proValorId dado, o -1 si no existe en la colección
    Dim varIndice As Long
    On Error GoTo ErrManager
        BuscarIndiceProValorId = -1
    For varIndice = 1 To mCol.Count
        If mCol(varIndice).proValorId = varTipoLinea Then
            BuscarIndiceProValorId = varIndice
            Exit For
        End If
    Next
    Exit Function
ErrManager:
    SubGMuestraError
End Function


