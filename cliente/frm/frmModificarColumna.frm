VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmModificarColumna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Columna"
   ClientHeight    =   2625
   ClientLeft      =   4425
   ClientTop       =   2970
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4920
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   2880
      TabIndex        =   0
      Top             =   2220
      Width           =   1545
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   345
      Left            =   270
      TabIndex        =   9
      Top             =   2220
      Width           =   1545
   End
   Begin VB.CheckBox chkValor 
      Caption         =   "Check1"
      Height          =   195
      Left            =   1140
      TabIndex        =   8
      Top             =   930
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cboCodigovalor 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   870
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboNombreValor 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   870
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.TextBox txtValor 
      Height          =   315
      Left            =   1140
      TabIndex        =   7
      Top             =   870
      Visible         =   0   'False
      Width           =   3285
   End
   Begin MSComCtl2.DTPicker dtValor 
      Height          =   315
      Left            =   1140
      TabIndex        =   6
      Top             =   870
      Visible         =   0   'False
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16842753
      CurrentDate     =   38183
   End
   Begin VB.ComboBox cboCodigoColumna 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   300
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.ComboBox cboNombreColumna 
      Height          =   315
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   330
      Width           =   3285
   End
   Begin VB.Frame fraTituloProducto 
      BackColor       =   &H00C09258&
      Caption         =   "  Información  del Producto  "
      ForeColor       =   &H00FFFFFF&
      Height          =   1395
      Left            =   30
      TabIndex        =   10
      Top             =   60
      Width           =   4665
      Begin VB.Label lblValor 
         Caption         =   "Valor:"
         Height          =   165
         Left            =   450
         TabIndex        =   12
         Top             =   930
         Width           =   465
      End
      Begin VB.Label lblColumna 
         AutoSize        =   -1  'True
         Caption         =   "Columna:"
         Height          =   195
         Left            =   330
         TabIndex        =   11
         Top             =   390
         Width           =   660
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C09258&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Si existen valores hijos de esta propiedad, serán eliminados en el momento de actualizar la prodiedad padre."
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   90
      TabIndex        =   3
      Top             =   1530
      Width           =   4365
   End
End
Attribute VB_Name = "frmModificarColumna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' FECHA: SEPTIEMBRE 7/2009
'****************************************************************

Option Explicit

Public proOnyx As EDCVoz.claONYX
Public proConexion As ADODB.Connection
Public proDatosProducto As claDatosProducto
Public proOrigen As String      'A      Actuales
                                'I      Insertados
                                
'VARIABLE QUE ME DA EL ID DEL CLIENTE
Public proiClienteId As Long '1.0.000

Public proNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
Public proDetalleDatosProducto As claDetalleDatosProducto

Private varProceso As claProceso

Private Sub cboNombreColumna_Click()
    Dim varColumna As Integer
    Dim varContador As Integer
    Dim varContadorAux As Integer
    Dim varValorPadre As String
    Dim varValorPadreAnterior As String
    Dim varEncontro As Boolean
    Dim varNumeroRegistro As Integer
    On Error GoTo ErrManager
    
    
    Me.cboCodigoColumna.ListIndex = Me.cboNombreColumna.ListIndex
    
    If Me.cboNombreColumna.ListIndex = -1 Then
        Exit Sub
    End If
    
    For varColumna = 1 To Me.proDatosProducto.proParametrosProducto.Count
        If Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampo) = Trim(Me.cboCodigoColumna.Text) Then
            Exit For
        End If
    Next varColumna
            
    'varColumna = Me.cboCodigoColumna.ListIndex + 1
    Select Case Me.proDatosProducto.proParametrosProducto.Item(varColumna).proTipo
        Case "L"
            
            Me.txtValor.Visible = False
            Me.chkValor.Visible = False
            Me.dtValor.Visible = False
            Me.cboNombreValor.Visible = False
            
            'Validar si tiene valor padre
            If Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampoPadre) <> "" Then
                
                'Validar que todos los registros seleccionados tengan el mismo valor en el padre
                If Me.proOrigen = "A" Then
                    
                    'Si solo se encuentra un registro seleccionado
                    If Me.proDatosProducto.proDetalleDatosProducto.proSeleccionados = 1 Then
                        
                        For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = 1 Then
                                Exit For
                            End If
                        Next varContador
                        
                        'Validar si el registro seleccionado ya se encuentra en la parte de modificacion
                        varEncontro = False
                        For varContadorAux = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId = _
                               Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContadorAux).proDetalleDatosProductoId Then
                               varEncontro = True
                               Exit For
                            End If
                        Next varContadorAux
                        
                        'Buscar cual es el padre seleccionado y cargar los valores hijos
                        Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampoPadre)
                            Case "vchUser1"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser1
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser1
                                End If
                            Case "vchUser2"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser2
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser2
                                End If
                            Case "vchUser3"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser3
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser3
                                End If
                            Case "vchUser4"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser4
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser4
                                End If
                            Case "vchUser5"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser5
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser5
                                End If
                            Case "vchUser6"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser6
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser6
                                End If
                            Case "vchUser7"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser7
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser7
                                End If
                            Case "vchUser8"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser8
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser8
                                End If
                            Case "vchUser9"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser9
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser9
                                End If
                            Case "vchUser10"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser10
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser10
                                End If
                            Case "vchUser11"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser11
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser11
                                End If
                            Case "vchUser12"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser12
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser12
                                End If
                            Case "vchUser13"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser13
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser13
                                End If
                            Case "vchUser14"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser14
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser14
                                End If
                            Case "vchUser15"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser15
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser15
                                End If
                            Case "vchUser16"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser16
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser16
                                End If
                            Case "vchUser17"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser17
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser17
                                End If
                            Case "vchUser18"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser18
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser18
                                End If
                            Case "vchUser19"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser19
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser19
                                End If
                            Case "vchUser20"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser20
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser20
                                End If
                            Case "vchUser21"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser21
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser21
                                End If
                            Case "vchUser22"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser22
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser22
                                End If
                            Case "vchUser23"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser23
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser23
                                End If
                            Case "vchUser24"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser24
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser24
                                End If
                            Case "vchUser25"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser25
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser25
                                End If
                            Case "vchUser26"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser26
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser26
                                End If
                            Case "vchUser27"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser27
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser27
                                End If
                            Case "vchUser28"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser28
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser28
                                End If
                            Case "vchUser29"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser29
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser29
                                End If
                            Case "vchUser30"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser30
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser30
                                End If
                            Case "vchUser31"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser31
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser31
                                End If
                            Case "vchUser32"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser32
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser32
                                End If
                            Case "vchUser33"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser33
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser33
                                End If
                            Case "vchUser34"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser34
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser34
                                End If
                            Case "vchUser35"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser35
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser35
                                End If
                            Case "vchUser36"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser36
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser36
                                End If
                            Case "vchUser37"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser37
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser37
                                End If
                            Case "vchUser38"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser38
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser38
                                End If
                            Case "vchUser39"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser39
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser39
                                End If
                            Case "vchUser40"
                                If varEncontro Then
                                    varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser40
                                Else
                                    varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser40
                                End If
                        End Select
                        
                        'Validar que el campo padre contenga algun valor
                        If Trim(varValorPadre) = "" Or varValorPadre = "0" Then
                            MsgBox "La propiedad [" + Me.proDatosProducto.proParametrosProducto.Item(varColumna).proEtiqueta + "] no se puede modificar porque su campo padre no posee un valor seleccionado.", vbInformation, App.Title
                            Me.cboNombreColumna.ListIndex = -1
                            Exit Sub
                        End If
                    Else 'Si Existe mas de un registro seleccionado
                            'Si es el primer registro, debe asignar el valor anterior.
                            'Si no lo es, debe comparar el valor anterior con el nuevo
                            'Si el valor es diferente no debe cargar los valores en el combo hijo
                            'y debe mostrar un mensaje de advertencia
                            
                        varNumeroRegistro = 0
                            
                        For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
                                
                            'Si el registro se encontraba seleccionado
                            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = 1 Then
                            
                                varNumeroRegistro = varNumeroRegistro + 1
                                
                                'Validar si el registro seleccionado ya se encuentra en la parte de modificacion
                                varEncontro = False
                                For varContadorAux = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
                                    If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId = _
                                       Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContadorAux).proDetalleDatosProductoId Then
                                       varEncontro = True
                                       Exit For
                                    End If
                                Next varContadorAux
                        
                                If varNumeroRegistro = 1 Then

                                    'Buscar cual es el padre seleccionado y cargar los valores hijos
                                    Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampoPadre)
                                        Case "vchUser1"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser1
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser1
                                            End If
                                        Case "vchUser2"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser2
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser2
                                            End If
                                        Case "vchUser3"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser3
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser3
                                            End If
                                        Case "vchUser4"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser4
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser4
                                            End If
                                        Case "vchUser5"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser5
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser5
                                            End If
                                        Case "vchUser6"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser6
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser6
                                            End If
                                        Case "vchUser7"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser7
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser7
                                            End If
                                        Case "vchUser8"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser8
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser8
                                            End If
                                        Case "vchUser9"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser9
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser9
                                            End If
                                        Case "vchUser10"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser10
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser10
                                            End If
                                        Case "vchUser11"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser11
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser11
                                            End If
                                        Case "vchUser12"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser12
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser12
                                            End If
                                        Case "vchUser13"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser13
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser13
                                            End If
                                        Case "vchUser14"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser14
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser14
                                            End If
                                        Case "vchUser15"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser15
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser15
                                            End If
                                        Case "vchUser16"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser16
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser16
                                            End If
                                        Case "vchUser17"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser17
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser17
                                            End If
                                        Case "vchUser18"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser18
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser18
                                            End If
                                        Case "vchUser19"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser19
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser19
                                            End If
                                        Case "vchUser20"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser20
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser20
                                            End If
                                        Case "vchUser21"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser21
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser21
                                            End If
                                        Case "vchUser22"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser22
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser22
                                            End If
                                        Case "vchUser23"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser23
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser23
                                            End If
                                        Case "vchUser24"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser24
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser24
                                            End If
                                        Case "vchUser25"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser25
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser25
                                            End If
                                        Case "vchUser26"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser26
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser26
                                            End If
                                        Case "vchUser27"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser27
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser27
                                            End If
                                        Case "vchUser28"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser28
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser28
                                            End If
                                        Case "vchUser29"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser29
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser29
                                            End If
                                        Case "vchUser30"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser30
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser30
                                            End If
                                        Case "vchUser31"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser31
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser31
                                            End If
                                        Case "vchUser32"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser32
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser32
                                            End If
                                        Case "vchUser33"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser33
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser33
                                            End If
                                        Case "vchUser34"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser34
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser34
                                            End If
                                        Case "vchUser35"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser35
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser35
                                            End If
                                        Case "vchUser36"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser36
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser36
                                            End If
                                        Case "vchUser37"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser37
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser37
                                            End If
                                        Case "vchUser38"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser38
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser38
                                            End If
                                        Case "vchUser39"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser39
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser39
                                            End If
                                        Case "vchUser40"
                                            If varEncontro Then
                                                varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser40
                                            Else
                                                varValorPadreAnterior = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser40
                                            End If
                                    End Select
                                Else ' Si no es el primer registro
                                    
                                    'Buscar cual es el padre seleccionado y cargar los valores hijos
                                    Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampoPadre)
                                        Case "vchUser1"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser1
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser1
                                            End If
                                        Case "vchUser2"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser2
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser2
                                            End If
                                        Case "vchUser3"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser3
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser3
                                            End If
                                        Case "vchUser4"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser4
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser4
                                            End If
                                        Case "vchUser5"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser5
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser5
                                            End If
                                        Case "vchUser6"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser6
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser6
                                            End If
                                        Case "vchUser7"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser7
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser7
                                            End If
                                        Case "vchUser8"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser8
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser8
                                            End If
                                        Case "vchUser9"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser9
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser9
                                            End If
                                        Case "vchUser10"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser10
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser10
                                            End If
                                        Case "vchUser11"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser11
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser11
                                            End If
                                        Case "vchUser12"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser12
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser12
                                            End If
                                        Case "vchUser13"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser13
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser13
                                            End If
                                        Case "vchUser14"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser14
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser14
                                            End If
                                        Case "vchUser15"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser15
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser15
                                            End If
                                        Case "vchUser16"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser16
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser16
                                            End If
                                        Case "vchUser17"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser17
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser17
                                            End If
                                        Case "vchUser18"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser18
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser18
                                            End If
                                        Case "vchUser19"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser19
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser19
                                            End If
                                        Case "vchUser20"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser20
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser20
                                            End If
                                        Case "vchUser21"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser21
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser21
                                            End If
                                        Case "vchUser22"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser22
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser22
                                            End If
                                        Case "vchUser23"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser23
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser23
                                            End If
                                        Case "vchUser24"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser24
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser24
                                            End If
                                        Case "vchUser25"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser25
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser25
                                            End If
                                        Case "vchUser26"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser26
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser26
                                            End If
                                        Case "vchUser27"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser27
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser27
                                            End If
                                        Case "vchUser28"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser28
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser28
                                            End If
                                        Case "vchUser29"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser29
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser29
                                            End If
                                        Case "vchUser30"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser30
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser30
                                            End If
                                        Case "vchUser31"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser31
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser31
                                            End If
                                        Case "vchUser32"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser32
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser32
                                            End If
                                        Case "vchUser33"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser33
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser33
                                            End If
                                        Case "vchUser34"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser34
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser34
                                            End If
                                        Case "vchUser35"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser35
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser35
                                            End If
                                        Case "vchUser36"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser36
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser36
                                            End If
                                        Case "vchUser37"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser37
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser37
                                            End If
                                        Case "vchUser38"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser38
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser38
                                            End If
                                        Case "vchUser39"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser39
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser39
                                            End If
                                        Case "vchUser40"
                                            If varEncontro Then
                                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser40
                                            Else
                                                varValorPadre = Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proUser40
                                            End If
                                    End Select
                                            
                                    'Validar que todos los padres contengan el mismo valor
                                    If varValorPadreAnterior <> varValorPadre Then
                                        MsgBox "Este campo no puede ser modificado, porque los registros seleccionados no tienen el mismo valor para el padre.", vbInformation, App.Title
                                        Exit Sub
                                    End If
                                End If 'Fin de si es el primer registro
                                
                            End If 'Fin de si se encontraba seleccionado el registro
                            
                        Next varContador 'Fin del for para los registros seleccionados
                    
                        'Validar que el campo padre tenga un valor seleccionado
                        If Trim(varValorPadre) = "" Or varValorPadre = "0" Then
                            MsgBox "La propiedad [" + Me.proDatosProducto.proParametrosProducto.Item(varColumna).proEtiqueta + "] no se puede modificar porque su campo padre no posee un valor seleccionado.", vbInformation, App.Title
                            Me.cboNombreColumna.ListIndex = -1
                            Exit Sub
                        End If
                    End If 'Fin de mas de un registro seleccionado
                    
                    'Hacer consulta
                    Me.proDatosProducto.proParametrosProducto.Item(varColumna).proValorIdPadre = varValorPadre
                    
                      ''* 1.0.100 Inicio Se pasa la propiedad del id del cliente
                    If Not Me.proDatosProducto.proParametrosProducto.Item(varColumna).MetConsultarValores(Me.proiClienteId) Then
                    '* 1.0.100 Fin
                        MsgBox "Error al consultar los valores", vbCritical, App.Title
                        Exit Sub
                    End If
                    
                Else 'Si el Origen no es A
                       
                    'Si solo se encuentra un registro seleccionado
                    If Me.proDatosProducto.proNovedadDetalleDatosProducto.proSeleccionados = 1 Then
                        
                        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
                            If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = 1 Then
                                Exit For
                            End If
                        Next varContador
                                                
                        'Buscar cual es el padre seleccionado y cargar los valores hijos
                        Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampoPadre)
                            Case "vchUser1"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser1
                            Case "vchUser2"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser2
                            Case "vchUser3"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser3
                            Case "vchUser4"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser4
                            Case "vchUser5"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser5
                            Case "vchUser6"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser6
                            Case "vchUser7"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser7
                            Case "vchUser8"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser8
                            Case "vchUser9"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser9
                            Case "vchUser10"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser10
                            Case "vchUser11"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser11
                            Case "vchUser12"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser12
                            Case "vchUser13"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser13
                            Case "vchUser14"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser14
                            Case "vchUser15"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser15
                            Case "vchUser16"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser16
                            Case "vchUser17"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser17
                            Case "vchUser18"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser18
                            Case "vchUser19"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser19
                            Case "vchUser20"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser20
                            Case "vchUser21"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser21
                            Case "vchUser22"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser22
                            Case "vchUser23"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser23
                            Case "vchUser24"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser24
                            Case "vchUser25"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser25
                            Case "vchUser26"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser26
                            Case "vchUser27"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser27
                            Case "vchUser28"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser28
                            Case "vchUser29"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser29
                            Case "vchUser30"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser30
                            Case "vchUser31"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser31
                            Case "vchUser32"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser32
                            Case "vchUser33"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser33
                            Case "vchUser34"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser34
                            Case "vchUser35"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser35
                            Case "vchUser36"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser36
                            Case "vchUser37"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser37
                            Case "vchUser38"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser38
                            Case "vchUser39"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser39
                            Case "vchUser40"
                                varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser40
                        End Select
                            
                        'Validar que el campo padre tenga un valor seleccionado
                        If Trim(varValorPadre) = "" Or varValorPadre = "0" Then
                            MsgBox "La propiedad [" + Me.proDatosProducto.proParametrosProducto.Item(varColumna).proEtiqueta + "] no se puede modificar porque su campo padre no posee un valor seleccionado.", vbInformation, App.Title
                            Me.cboNombreColumna.ListIndex = -1
                            Exit Sub
                        End If
                    Else 'Si se encuentra más de un registro seleccionado
                            'Si es el primer registro, debe asignar el valor anterior.
                            'Si no lo es, debe comparar el valor anterior con el nuevo
                            'Si el valor es diferente no debe cargar los valores en el combo hijo
                            'y debe mostrar un mensaje de advertencia
                            
                        varNumeroRegistro = 0
                        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
                            
                            If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = 1 Then
                                    
                                varNumeroRegistro = varNumeroRegistro + 1
                                    
                                'Si es el primer registro seleccionado
                                If varNumeroRegistro = 1 Then
                                                                
                                    Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampoPadre)
                                        Case "vchUser1"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser1
                                        Case "vchUser2"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser2
                                        Case "vchUser3"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser3
                                        Case "vchUser4"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser4
                                        Case "vchUser5"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser5
                                        Case "vchUser6"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser6
                                        Case "vchUser7"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser7
                                        Case "vchUser8"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser8
                                        Case "vchUser9"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser9
                                        Case "vchUser10"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser10
                                        Case "vchUser11"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser11
                                        Case "vchUser12"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser12
                                        Case "vchUser13"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser13
                                        Case "vchUser14"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser14
                                        Case "vchUser15"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser15
                                        Case "vchUser16"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser16
                                        Case "vchUser17"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser17
                                        Case "vchUser18"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser18
                                        Case "vchUser19"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser19
                                        Case "vchUser20"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser20
                                        Case "vchUser21"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser21
                                        Case "vchUser22"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser22
                                        Case "vchUser23"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser23
                                        Case "vchUser24"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser24
                                        Case "vchUser25"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser25
                                        Case "vchUser26"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser26
                                        Case "vchUser27"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser27
                                        Case "vchUser28"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser28
                                        Case "vchUser29"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser29
                                        Case "vchUser30"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser30
                                        Case "vchUser31"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser31
                                        Case "vchUser32"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser32
                                        Case "vchUser33"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser33
                                        Case "vchUser34"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser34
                                        Case "vchUser35"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser35
                                        Case "vchUser36"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser36
                                        Case "vchUser37"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser37
                                        Case "vchUser38"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser38
                                        Case "vchUser39"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser39
                                        Case "vchUser40"
                                            varValorPadreAnterior = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser40
                                    End Select
                                Else ' Si no es el primer registro
                                    Select Case Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampoPadre)
                                        Case "vchUser1"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser1
                                        Case "vchUser2"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser2
                                        Case "vchUser3"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser3
                                        Case "vchUser4"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser4
                                        Case "vchUser5"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser5
                                        Case "vchUser6"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser6
                                        Case "vchUser7"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser7
                                        Case "vchUser8"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser8
                                        Case "vchUser9"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser9
                                        Case "vchUser10"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser10
                                        Case "vchUser11"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser11
                                        Case "vchUser12"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser12
                                        Case "vchUser13"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser13
                                        Case "vchUser14"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser14
                                        Case "vchUser15"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser15
                                        Case "vchUser16"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser16
                                        Case "vchUser17"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser17
                                        Case "vchUser18"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser18
                                        Case "vchUser19"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser19
                                        Case "vchUser20"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser20
                                        Case "vchUser21"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser21
                                        Case "vchUser22"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser22
                                        Case "vchUser23"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser23
                                        Case "vchUser24"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser24
                                        Case "vchUser25"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser25
                                        Case "vchUser26"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser26
                                        Case "vchUser27"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser27
                                        Case "vchUser28"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser28
                                        Case "vchUser29"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser29
                                        Case "vchUser30"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser30
                                        Case "vchUser31"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser31
                                        Case "vchUser32"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser32
                                        Case "vchUser33"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser33
                                        Case "vchUser34"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser34
                                        Case "vchUser35"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser35
                                        Case "vchUser36"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser36
                                        Case "vchUser37"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser37
                                        Case "vchUser38"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser38
                                        Case "vchUser39"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser39
                                        Case "vchUser40"
                                            varValorPadre = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proUser40
                                    End Select
                                    
                                    'Validar que todos los padres seleccionados deben tener el mismo valor
                                    If varValorPadreAnterior <> varValorPadre Then
                                        MsgBox "Este campo no puede ser modificado, porque los registros seleccionados no tienen el mismo valor para el padre.", vbInformation, App.Title
                                        Exit Sub
                                    End If
                                End If 'Fin de si no es el primer registro
                            End If 'Fin de si el registro se encontraba seleccionado
                        Next varContador 'Fin del ciclo de registros seleccionados
                
                        'Validar que el campo padre tenga un valor seleccionado
                        If Trim(varValorPadre) = "" Or varValorPadre = "0" Then
                            MsgBox "La propiedad [" + Me.proDatosProducto.proParametrosProducto.Item(varColumna).proEtiqueta + "] no se puede modificar porque su campo padre no posee un valor seleccionado.", vbInformation, App.Title
                            Me.cboNombreColumna.ListIndex = -1
                            Exit Sub
                        End If
                        
                    End If 'Fin de si se encontraba más de un registro seleccionado
                    
                    'Hacer consulta
                    Me.proDatosProducto.proParametrosProducto.Item(varColumna).proValorIdPadre = varValorPadre
                    
                    ''* 1.0.100 Inicio Se pasa la propiedad del id del cliente
                    If Not Me.proDatosProducto.proParametrosProducto.Item(varColumna).MetConsultarValores(Me.proiClienteId) Then
                    '* 1.0.100 Fin
                        MsgBox "Error al consultar los valores", vbCritical, App.Title
                        Exit Sub
                    End If
                End If 'Fin de si el origen no es A
            End If 'Fin de si tiene un campo padre
            
            Call SubFLlenarComboValores(varColumna)
            Me.cboNombreValor.Visible = True
        Case "T"
            Me.cboNombreValor.Visible = False
            Me.txtValor.Visible = True
            Me.chkValor.Visible = False
            Me.dtValor.Visible = False
        Case "F"
            Me.cboNombreValor.Visible = False
            Me.txtValor.Visible = False
            Me.chkValor.Visible = False
            Me.dtValor.Visible = True
            Me.dtValor.Value = Now
        Case "B"
            Me.cboNombreValor.Visible = False
            Me.txtValor.Visible = False
            Me.chkValor.Visible = True
            Me.dtValor.Visible = False
    End Select
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cboNombreValor_Click()
    On Error GoTo ErrManager
    
    Me.cboCodigovalor.ListIndex = Me.cboNombreValor.ListIndex
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub cmdGuardar_Click()
    Dim varContador As Integer
    Dim varContadorAux As Integer
    Dim varEncontro As Boolean
    Dim varColumna As Integer
    Dim varCodigos As String
    Dim varNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto
    
    Dim varPadre As String
    Dim varEntro As Boolean
    Dim varCuentaEntradas As Integer
    Dim varContadorAux2 As Integer
    Dim varPrimerRegistro As Integer
    Dim varEtiqueta As String
    
    Dim varCampoPadre1 As String
    Dim varCampoPadre2 As String
    Dim varCampoPadre3 As String
    Dim varValorPadre1 As String
    Dim varValorPadre2 As String
    Dim varValorPadre3 As String
    
    On Error GoTo ErrManager
    
    Screen.MousePointer = 11
    
    If Me.cboNombreColumna.ListIndex = -1 Then
        MsgBox "Debe seleccionar el campo a modificar.", vbInformation, App.Title
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    For varColumna = 1 To Me.proDatosProducto.proParametrosProducto.Count
        If Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampo) = Trim(Me.cboCodigoColumna.Text) Then
            Exit For
        End If
    Next varColumna
    
    Select Case Me.proDatosProducto.proParametrosProducto.Item(varColumna).proTipo
        Case "L"
            If Trim(Me.cboNombreValor.Text) = "" Then
                MsgBox "Debe seleccionar el valor que desea almacenar.", vbInformation, App.Title
                Screen.MousePointer = 0
                Exit Sub
            End If
        Case "T"
            If Trim(Me.txtValor.Text) = "" Then
                MsgBox "Debe colocar el valor a modificar", vbInformation, App.Title
                Screen.MousePointer = 0
                Exit Sub
            End If
    End Select
    
        
    If Me.proOrigen = "A" Then 'Registros actuales
    
        'Si el parametro requiere control de uso, lo valida
        If Me.proDatosProducto.proParametrosProducto.Item(varColumna).proValidarRepetidos = 1 Then
                
            'Buscar el primer registro seleccionado
            For varPrimerRegistro = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
                If Me.proDatosProducto.proDetalleDatosProducto.Item(varPrimerRegistro).proSeleccion = 1 Then
                    Exit For
                End If
            Next varPrimerRegistro
            
            'Asignar el primer registro a la clase
            Set Me.proDetalleDatosProducto = Me.proDatosProducto.proDetalleDatosProducto.Item(varPrimerRegistro)
            
            'Buscar los padres y los valores respectivos
            varPadre = Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampoPadre)
            varEntro = False
            varCuentaEntradas = 0
            
            While Trim(varPadre) <> ""
                varEntro = True
                varCuentaEntradas = varCuentaEntradas + 1
                
                'Buscar el padre
                For varContadorAux2 = 1 To Me.proDatosProducto.proParametrosProducto.Count
                    If Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux2).proCampo) = Trim(varPadre) Then
                        Exit For
                    End If
                Next varContadorAux2
                
                If varCuentaEntradas = 1 Then
                    varCampoPadre1 = Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux2).proCampo)
                    varCampoPadre2 = Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampo)
                    varCampoPadre3 = ""
                    
                    'Asignar el valor del padre 1
                    Select Case Trim(varCampoPadre1)
                        Case "vchUser1"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser1
                        Case "vchUser2"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser2
                        Case "vchUser3"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser3
                        Case "vchUser4"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser4
                        Case "vchUser5"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser5
                        Case "vchUser6"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser6
                        Case "vchUser7"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser7
                        Case "vchUser8"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser8
                        Case "vchUser9"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser9
                        Case "vchUser10"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser10
                        Case "vchUser11"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser11
                        Case "vchUser12"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser12
                        Case "vchUser13"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser13
                        Case "vchUser14"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser14
                        Case "vchUser15"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser15
                        Case "vchUser16"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser16
                        Case "vchUser7"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser17
                        Case "vchUser18"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser18
                        Case "vchUser19"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser19
                        Case "vchUser20"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser20
                        Case "vchUser21"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser21
                        Case "vchUser22"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser22
                        Case "vchUser23"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser23
                        Case "vchUser24"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser24
                        Case "vchUser25"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser25
                        Case "vchUser26"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser26
                        Case "vchUser27"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser27
                        Case "vchUser28"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser28
                        Case "vchUser29"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser29
                        Case "vchUser30"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser30
                        Case "vchUser31"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser31
                        Case "vchUser32"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser32
                        Case "vchUser33"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser33
                        Case "vchUser34"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser34
                        Case "vchUser35"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser35
                        Case "vchUser36"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser36
                        Case "vchUser37"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser37
                        Case "vchUser38"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser38
                        Case "vchUser39"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser39
                        Case "vchUser40"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser40
                    End Select
                
                    Select Case Me.proDatosProducto.proParametrosProducto.Item(varColumna).proTipo
                        Case "L"
                            varValorPadre2 = Me.cboCodigovalor.Text
                        Case "T"
                            varValorPadre2 = Me.txtValor.Text
                        Case "F"
                            varValorPadre2 = Me.dtValor.Value
                        Case "B"
                            varValorPadre2 = Me.chkValor.Value
                    End Select
                   
                    varValorPadre3 = ""
                Else
                    varCampoPadre3 = Trim(varCampoPadre2)
                    varCampoPadre2 = Trim(varCampoPadre1)
                    varCampoPadre1 = Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux2).proCampo)
                    
                    varValorPadre3 = Trim(varValorPadre2)
                    varValorPadre2 = Trim(varValorPadre1)
                    
                    'Asignar el valor del padre 1
                    Select Case Trim(varCampoPadre1)
                        Case "vchUser1"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser1
                        Case "vchUser2"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser2
                        Case "vchUser3"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser3
                        Case "vchUser4"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser4
                        Case "vchUser5"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser5
                        Case "vchUser6"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser6
                        Case "vchUser7"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser7
                        Case "vchUser8"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser8
                        Case "vchUser9"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser9
                        Case "vchUser10"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser10
                        Case "vchUser11"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser11
                        Case "vchUser12"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser12
                        Case "vchUser13"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser13
                        Case "vchUser14"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser14
                        Case "vchUser15"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser15
                        Case "vchUser16"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser16
                        Case "vchUser7"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser17
                        Case "vchUser18"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser18
                        Case "vchUser19"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser19
                        Case "vchUser20"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser20
                        Case "vchUser21"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser21
                        Case "vchUser22"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser22
                        Case "vchUser23"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser23
                        Case "vchUser24"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser24
                        Case "vchUser25"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser25
                        Case "vchUser26"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser26
                        Case "vchUser27"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser27
                        Case "vchUser28"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser28
                        Case "vchUser29"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser29
                        Case "vchUser30"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser30
                        Case "vchUser31"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser31
                        Case "vchUser32"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser32
                        Case "vchUser33"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser33
                        Case "vchUser34"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser34
                        Case "vchUser35"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser35
                        Case "vchUser36"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser36
                        Case "vchUser37"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser37
                        Case "vchUser38"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser38
                        Case "vchUser39"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser39
                        Case "vchUser40"
                            varValorPadre1 = Me.proDetalleDatosProducto.proUser40
                    End Select
                End If
                varPadre = Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux2).proCampoPadre)
            Wend
            
            If Not varEntro Then
                varCampoPadre1 = Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampo)
                varCampoPadre2 = ""
                varCampoPadre3 = ""
                Select Case Trim(varCampoPadre1)
                    Case "vchUser1"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser1
                    Case "vchUser2"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser2
                    Case "vchUser3"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser3
                    Case "vchUser4"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser4
                    Case "vchUser5"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser5
                    Case "vchUser6"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser6
                    Case "vchUser7"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser7
                    Case "vchUser8"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser8
                    Case "vchUser9"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser9
                    Case "vchUser10"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser10
                    Case "vchUser11"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser11
                    Case "vchUser12"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser12
                    Case "vchUser13"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser13
                    Case "vchUser14"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser14
                    Case "vchUser15"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser15
                    Case "vchUser16"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser16
                    Case "vchUser7"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser17
                    Case "vchUser18"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser18
                    Case "vchUser19"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser19
                    Case "vchUser20"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser20
                    Case "vchUser21"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser21
                    Case "vchUser22"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser22
                    Case "vchUser23"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser23
                    Case "vchUser24"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser24
                    Case "vchUser25"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser25
                    Case "vchUser26"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser26
                    Case "vchUser27"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser27
                    Case "vchUser28"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser28
                    Case "vchUser29"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser29
                    Case "vchUser30"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser30
                    Case "vchUser31"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser31
                    Case "vchUser32"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser32
                    Case "vchUser33"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser33
                    Case "vchUser34"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser34
                    Case "vchUser35"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser35
                    Case "vchUser36"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser36
                    Case "vchUser37"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser37
                    Case "vchUser38"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser38
                    Case "vchUser39"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser39
                    Case "vchUser40"
                        varValorPadre1 = Me.proDetalleDatosProducto.proUser40
                End Select
            End If
            If Trim(varValorPadre1) <> "" And varValorPadre1 <> "0" Then
            
                If Not Me.proDatosProducto.proParametrosProducto.Item(varColumna).MetValidarInformacionCampo(varCampoPadre1, varCampoPadre2, varCampoPadre3, varValorPadre1, varValorPadre2, varValorPadre3, Me.proDatosProducto.proDatosProductoId, varEtiqueta) Then
                    MsgBox "El campo [" & varEtiqueta & "] tiene un valor que ya fue usado en otro servicio.", vbInformation, App.Title
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
    
        'Guardar el encabezado - Si es la primera vez lo inserta - Si no lo actualiza
        If Not Me.proDatosProducto.MetGuardar Then
            MsgBox "Error al actualizar la información del producto.", vbCritical, App.Title
            Exit Sub
        End If
        
        'Inserta o actualiza la información de los incidentes
        If Not Me.proDatosProducto.MetGuardarColeccionIncidentes Then
            MsgBox "Error al almacenar el incidente asociado.", vbCritical, App.Title
        End If
        
        For varContador = 1 To Me.proDatosProducto.proDetalleDatosProducto.Count
            'Validar si el registro fue seleccionado para modificación o no
            If Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proSeleccion = "1" Then
                
                varCodigos = varCodigos & Me.proDatosProducto.proDetalleDatosProducto.Item(varContador).proDetalleDatosProductoId & ", "
                
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proCampo = Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampo)
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proCodigos = Mid(varCodigos, 1, Len(varCodigos) - 1)
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proIncidentId = Me.proDatosProducto.proIncidentId
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proProductNumber = Me.proDatosProducto.proProductNumber
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proTabla = "1"
                
                Select Case Me.proDatosProducto.proParametrosProducto.Item(varColumna).proTipo
                    Case "L"
                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proValor = Me.cboCodigovalor.Text
                    Case "T"
                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proValor = Me.txtValor.Text
                    Case "F"
                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proValor = Me.dtValor.Value
                    Case "B"
                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proValor = Me.chkValor.Value
                End Select
                
                If Me.proDatosProducto.proNovedadDetalleDatosProducto.MetActualizarColumna Then
                    If Not Me.proDatosProducto.MetConsultarNovedadDetalleDatosProducto Then
                        MsgBox "Error al agregar el elemento a la colección.", vbCritical, App.Title
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                Else
                    MsgBox "Error al actualizar la información del registro.", vbCritical, App.Title
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        Next varContador
    Else
                
        'Si el parametro requiere control de uso, lo valida
        If Me.proDatosProducto.proParametrosProducto.Item(varColumna).proValidarRepetidos = 1 Then
                
            'Buscar el primer registro seleccionado
            For varPrimerRegistro = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
                If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varPrimerRegistro).proSeleccion = 1 Then
                    Exit For
                End If
            Next varPrimerRegistro
            
            'Asignar el primer registro a la clase
            Set Me.proNovedadDetalleDatosProducto = Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varPrimerRegistro)
            
            'Buscar los padres y los valores respectivos
            varPadre = Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampoPadre)
            varEntro = False
            varCuentaEntradas = 0
            
            While Trim(varPadre) <> ""
                varEntro = True
                varCuentaEntradas = varCuentaEntradas + 1
                
                'Buscar el padre
                For varContadorAux2 = 1 To Me.proDatosProducto.proParametrosProducto.Count
                    If Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux2).proCampo) = Trim(varPadre) Then
                        Exit For
                    End If
                Next varContadorAux2
                
                If varCuentaEntradas = 1 Then
                    varCampoPadre1 = Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux2).proCampo)
                    varCampoPadre2 = Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampo)
                    varCampoPadre3 = ""
                    
                    'Asignar el valor del padre 1
                    Select Case Trim(varCampoPadre1)
                        Case "vchUser1"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser1
                        Case "vchUser2"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser2
                        Case "vchUser3"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser3
                        Case "vchUser4"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser4
                        Case "vchUser5"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser5
                        Case "vchUser6"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser6
                        Case "vchUser7"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser7
                        Case "vchUser8"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser8
                        Case "vchUser9"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser9
                        Case "vchUser10"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser10
                        Case "vchUser11"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser11
                        Case "vchUser12"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser12
                        Case "vchUser13"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser13
                        Case "vchUser14"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser14
                        Case "vchUser15"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser15
                        Case "vchUser16"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser16
                        Case "vchUser7"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser17
                        Case "vchUser18"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser18
                        Case "vchUser19"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser19
                        Case "vchUser20"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser20
                        Case "vchUser21"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser21
                        Case "vchUser22"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser22
                        Case "vchUser23"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser23
                        Case "vchUser24"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser24
                        Case "vchUser25"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser25
                        Case "vchUser26"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser26
                        Case "vchUser27"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser27
                        Case "vchUser28"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser28
                        Case "vchUser29"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser29
                        Case "vchUser30"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser30
                        Case "vchUser31"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser31
                        Case "vchUser32"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser32
                        Case "vchUser33"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser33
                        Case "vchUser34"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser34
                        Case "vchUser35"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser35
                        Case "vchUser36"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser36
                        Case "vchUser37"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser37
                        Case "vchUser38"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser38
                        Case "vchUser39"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser39
                        Case "vchUser40"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser40
                    End Select
                
                    Select Case Me.proDatosProducto.proParametrosProducto.Item(varColumna).proTipo
                        Case "L"
                            varValorPadre2 = Me.cboCodigovalor.Text
                        Case "T"
                            varValorPadre2 = Me.txtValor.Text
                        Case "F"
                            varValorPadre2 = Me.dtValor.Value
                        Case "B"
                            varValorPadre2 = Me.chkValor.Value
                    End Select
                   
                    varValorPadre3 = ""
                Else
                    varCampoPadre3 = Trim(varCampoPadre2)
                    varCampoPadre2 = Trim(varCampoPadre1)
                    varCampoPadre1 = Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux2).proCampo)
                    
                    varValorPadre3 = Trim(varValorPadre2)
                    varValorPadre2 = Trim(varValorPadre1)
                    
                    'Asignar el valor del padre 1
                    Select Case Trim(varCampoPadre1)
                        Case "vchUser1"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser1
                        Case "vchUser2"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser2
                        Case "vchUser3"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser3
                        Case "vchUser4"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser4
                        Case "vchUser5"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser5
                        Case "vchUser6"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser6
                        Case "vchUser7"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser7
                        Case "vchUser8"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser8
                        Case "vchUser9"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser9
                        Case "vchUser10"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser10
                        Case "vchUser11"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser11
                        Case "vchUser12"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser12
                        Case "vchUser13"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser13
                        Case "vchUser14"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser14
                        Case "vchUser15"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser15
                        Case "vchUser16"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser16
                        Case "vchUser7"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser17
                        Case "vchUser18"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser18
                        Case "vchUser19"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser19
                        Case "vchUser20"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser20
                        Case "vchUser21"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser21
                        Case "vchUser22"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser22
                        Case "vchUser23"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser23
                        Case "vchUser24"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser24
                        Case "vchUser25"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser25
                        Case "vchUser26"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser26
                        Case "vchUser27"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser27
                        Case "vchUser28"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser28
                        Case "vchUser29"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser29
                        Case "vchUser30"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser30
                        Case "vchUser31"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser31
                        Case "vchUser32"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser32
                        Case "vchUser33"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser33
                        Case "vchUser34"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser34
                        Case "vchUser35"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser35
                        Case "vchUser36"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser36
                        Case "vchUser37"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser37
                        Case "vchUser38"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser38
                        Case "vchUser39"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser39
                        Case "vchUser40"
                            varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser40
                    End Select
                End If
                varPadre = Trim(Me.proDatosProducto.proParametrosProducto.Item(varContadorAux2).proCampoPadre)
            Wend
            
            If Not varEntro Then
                varCampoPadre1 = Trim(Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampo)
                varCampoPadre2 = ""
                varCampoPadre3 = ""
                Select Case Trim(varCampoPadre1)
                    Case "vchUser1"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser1
                    Case "vchUser2"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser2
                    Case "vchUser3"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser3
                    Case "vchUser4"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser4
                    Case "vchUser5"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser5
                    Case "vchUser6"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser6
                    Case "vchUser7"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser7
                    Case "vchUser8"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser8
                    Case "vchUser9"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser9
                    Case "vchUser10"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser10
                    Case "vchUser11"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser11
                    Case "vchUser12"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser12
                    Case "vchUser13"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser13
                    Case "vchUser14"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser14
                    Case "vchUser15"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser15
                    Case "vchUser16"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser16
                    Case "vchUser7"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser17
                    Case "vchUser18"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser18
                    Case "vchUser19"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser19
                    Case "vchUser20"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser20
                    Case "vchUser21"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser21
                    Case "vchUser22"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser22
                    Case "vchUser23"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser23
                    Case "vchUser24"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser24
                    Case "vchUser25"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser25
                    Case "vchUser26"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser26
                    Case "vchUser27"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser27
                    Case "vchUser28"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser28
                    Case "vchUser29"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser29
                    Case "vchUser30"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser30
                    Case "vchUser31"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser31
                    Case "vchUser32"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser32
                    Case "vchUser33"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser33
                    Case "vchUser34"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser34
                    Case "vchUser35"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser35
                    Case "vchUser36"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser36
                    Case "vchUser37"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser37
                    Case "vchUser38"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser38
                    Case "vchUser39"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser39
                    Case "vchUser40"
                        varValorPadre1 = Me.proNovedadDetalleDatosProducto.proUser40
                End Select
            End If
            
            If Trim(varValorPadre1) <> "" And varValorPadre1 <> "0" Then
            
                If Not Me.proDatosProducto.proParametrosProducto.Item(varColumna).MetValidarInformacionCampo(varCampoPadre1, varCampoPadre2, varCampoPadre3, varValorPadre1, varValorPadre2, varValorPadre3, Me.proDatosProducto.proDatosProductoId, varEtiqueta) Then
                    MsgBox "El campo [" & varEtiqueta & "] tiene un valor que ya fue usado en otro servicio.", vbInformation, App.Title
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        End If
        
        'Registros insertados
        For varContador = 1 To Me.proDatosProducto.proNovedadDetalleDatosProducto.Count
            
            'Validar si el registro fue seleccionado para modificación o no
            If Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proSeleccion = "1" Then
            
                varCodigos = varCodigos & Me.proDatosProducto.proNovedadDetalleDatosProducto.Item(varContador).proNovedadDetalleDatosProductoId & ", "
                
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proCampo = Me.proDatosProducto.proParametrosProducto.Item(varColumna).proCampo
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proCodigos = Mid(varCodigos, 1, Len(varCodigos) - 2)
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proDatosProductoId = Me.proDatosProducto.proDatosProductoId
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proIncidentId = Me.proDatosProducto.proIncidentId
                Me.proDatosProducto.proNovedadDetalleDatosProducto.proProductNumber = Me.proDatosProducto.proProductNumber
                
                Select Case Me.proDatosProducto.proParametrosProducto.Item(varColumna).proTipo
                    Case "L"
                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proValor = Me.cboCodigovalor.Text
                    Case "T"
                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proValor = Me.txtValor.Text
                    Case "F"
                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proValor = Me.dtValor.Value
                    Case "B"
                        Me.proDatosProducto.proNovedadDetalleDatosProducto.proValor = Me.chkValor.Value
                End Select
            End If
        Next varContador
        If Me.proDatosProducto.proNovedadDetalleDatosProducto.MetActualizarColumna Then
            If Not Me.proDatosProducto.MetConsultarNovedadDetalleDatosProducto Then
                MsgBox "Error al agregar el elemento a la colección.", vbCritical, App.Title
                Screen.MousePointer = 0
                Exit Sub
            End If
        Else
            MsgBox "Error al actualizar la información del registro.", vbCritical, App.Title
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    Screen.MousePointer = 0
    Unload Me
    Exit Sub
ErrManager:
    Screen.MousePointer = 0
    SubGMuestraError
End Sub

Private Sub cmdSalir_Click()
    On Error GoTo ErrManager
    
    Unload Me
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    On Error GoTo ErrManager
    
    'Consultar los datos del incidente
    Set varProceso = New claProceso
    Set varProceso.proConexion = Me.proConexion
    
    varProceso.proIncidentId = Me.proDatosProducto.proIncidentId
    
    If Not varProceso.MetConsultaDatosIncidente Then
        MsgBox "Error al buscar la información del incidente.", vbCritical, App.Title
        Exit Sub
    End If
    
    Call SubFLlenarComboColumnas
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboColumnas()
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.cboCodigoColumna.Clear
    Me.cboNombreColumna.Clear
    
    For varContador = 1 To Me.proDatosProducto.proParametrosProducto.Count
        If FunFCamposObligatorios(Me.proDatosProducto.proParametrosProducto.Item(varContador).proCampo) Then
            Me.cboCodigoColumna.AddItem Me.proDatosProducto.proParametrosProducto.Item(varContador).proCampo
            Me.cboNombreColumna.AddItem Me.proDatosProducto.proParametrosProducto.Item(varContador).proEtiqueta
        End If
    Next varContador
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub

Private Sub SubFLlenarComboValores(parRegistro As Integer)
    Dim varContador As Integer
    On Error GoTo ErrManager
    
    Me.cboCodigovalor.Clear
    Me.cboNombreValor.Clear
    
    For varContador = 1 To Me.proDatosProducto.proParametrosProducto.Item(parRegistro).proValores.Count
        Me.cboCodigovalor.AddItem Me.proDatosProducto.proParametrosProducto.Item(parRegistro).proValores.Item(varContador).proValorID
        Me.cboNombreValor.AddItem Me.proDatosProducto.proParametrosProducto.Item(parRegistro).proValores.Item(varContador).proValorDesc
    Next varContador
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Public Sub SubFAsignarRegistroAModificar(ByRef parDetalleDatosProducto As claDetalleDatosProducto, ByRef parNovedadDetalleDatosProducto As claNovedadDetalleDatosProducto, ByVal parTipoNovedadId As Integer)
    On Error GoTo ErrManager
    
    parNovedadDetalleDatosProducto.proDatosProductoId = parDetalleDatosProducto.proDatosProductoId
    parNovedadDetalleDatosProducto.proDetalleDatosProductoId = parDetalleDatosProducto.proDetalleDatosProductoId
    parNovedadDetalleDatosProducto.proIncidentId = Me.proDatosProducto.proIncidentId
    parNovedadDetalleDatosProducto.proRecordStatus = parDetalleDatosProducto.proRecordStatus
    parNovedadDetalleDatosProducto.proStatusId = parDetalleDatosProducto.proStatusId
    parNovedadDetalleDatosProducto.proTipoNovedadId = parTipoNovedadId
    parNovedadDetalleDatosProducto.proUser1 = parDetalleDatosProducto.proUser1
    parNovedadDetalleDatosProducto.proUser2 = parDetalleDatosProducto.proUser2
    parNovedadDetalleDatosProducto.proUser3 = parDetalleDatosProducto.proUser3
    parNovedadDetalleDatosProducto.proUser4 = parDetalleDatosProducto.proUser4
    parNovedadDetalleDatosProducto.proUser5 = parDetalleDatosProducto.proUser5
    parNovedadDetalleDatosProducto.proUser6 = parDetalleDatosProducto.proUser6
    parNovedadDetalleDatosProducto.proUser7 = parDetalleDatosProducto.proUser7
    parNovedadDetalleDatosProducto.proUser8 = parDetalleDatosProducto.proUser8
    parNovedadDetalleDatosProducto.proUser9 = parDetalleDatosProducto.proUser9
    parNovedadDetalleDatosProducto.proUser10 = parDetalleDatosProducto.proUser10
    parNovedadDetalleDatosProducto.proUser11 = parDetalleDatosProducto.proUser11
    parNovedadDetalleDatosProducto.proUser12 = parDetalleDatosProducto.proUser12
    parNovedadDetalleDatosProducto.proUser13 = parDetalleDatosProducto.proUser13
    parNovedadDetalleDatosProducto.proUser14 = parDetalleDatosProducto.proUser14
    parNovedadDetalleDatosProducto.proUser15 = parDetalleDatosProducto.proUser15
    parNovedadDetalleDatosProducto.proUser16 = parDetalleDatosProducto.proUser16
    parNovedadDetalleDatosProducto.proUser17 = parDetalleDatosProducto.proUser17
    parNovedadDetalleDatosProducto.proUser18 = parDetalleDatosProducto.proUser18
    parNovedadDetalleDatosProducto.proUser19 = parDetalleDatosProducto.proUser19
    parNovedadDetalleDatosProducto.proUser20 = parDetalleDatosProducto.proUser20
    parNovedadDetalleDatosProducto.proUser21 = parDetalleDatosProducto.proUser21
    parNovedadDetalleDatosProducto.proUser22 = parDetalleDatosProducto.proUser22
    parNovedadDetalleDatosProducto.proUser23 = parDetalleDatosProducto.proUser23
    parNovedadDetalleDatosProducto.proUser24 = parDetalleDatosProducto.proUser24
    parNovedadDetalleDatosProducto.proUser25 = parDetalleDatosProducto.proUser25
    parNovedadDetalleDatosProducto.proUser26 = parDetalleDatosProducto.proUser26
    parNovedadDetalleDatosProducto.proUser27 = parDetalleDatosProducto.proUser27
    parNovedadDetalleDatosProducto.proUser28 = parDetalleDatosProducto.proUser28
    parNovedadDetalleDatosProducto.proUser29 = parDetalleDatosProducto.proUser29
    parNovedadDetalleDatosProducto.proUser30 = parDetalleDatosProducto.proUser30
    parNovedadDetalleDatosProducto.proUser31 = parDetalleDatosProducto.proUser31
    parNovedadDetalleDatosProducto.proUser32 = parDetalleDatosProducto.proUser32
    parNovedadDetalleDatosProducto.proUser33 = parDetalleDatosProducto.proUser33
    parNovedadDetalleDatosProducto.proUser34 = parDetalleDatosProducto.proUser34
    parNovedadDetalleDatosProducto.proUser35 = parDetalleDatosProducto.proUser35
    parNovedadDetalleDatosProducto.proUser36 = parDetalleDatosProducto.proUser36
    parNovedadDetalleDatosProducto.proUser37 = parDetalleDatosProducto.proUser37
    parNovedadDetalleDatosProducto.proUser38 = parDetalleDatosProducto.proUser38
    parNovedadDetalleDatosProducto.proUser39 = parDetalleDatosProducto.proUser39
    parNovedadDetalleDatosProducto.proUser40 = parDetalleDatosProducto.proUser40
    
    Exit Sub
ErrManager:
    SubGMuestraError
End Sub


Public Function FunFCamposObligatorios(parCampo As String) As Boolean
    Dim varCategoria As Categoria
    Dim varContador As Integer
    On Error GoTo ErrManager
        
    varCategoria = varProceso.proIncidentCategory
    If varCategoria = Venta Then
        If Trim(varProceso.proOTId) = "" Then
            For varContador = 1 To Me.proDatosProducto.proParametrosProducto.Count
                If Trim(parCampo) = Trim(Me.proDatosProducto.proParametrosProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proDatosProducto.proParametrosProducto.Item(varContador).proObligatorioVenta = True Then
                FunFCamposObligatorios = True
            Else
                FunFCamposObligatorios = False
            End If
        Else
            For varContador = 1 To Me.proDatosProducto.proParametrosProducto.Count
                If Trim(parCampo) = Trim(Me.proDatosProducto.proParametrosProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proDatosProducto.proParametrosProducto.Item(varContador).proObligatorioOT = True Then
                FunFCamposObligatorios = True
            Else
                FunFCamposObligatorios = False
            End If
        End If
    ElseIf varCategoria = Atencion Then
        If Trim(varProceso.proOTId) = "" Then
            For varContador = 1 To Me.proDatosProducto.proParametrosProducto.Count
                If Trim(parCampo) = Trim(Me.proDatosProducto.proParametrosProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proDatosProducto.proParametrosProducto.Item(varContador).proObligatorioAtencion = True Then
                FunFCamposObligatorios = True
            Else
                FunFCamposObligatorios = False
            End If
        Else
            For varContador = 1 To Me.proDatosProducto.proParametrosProducto.Count
                If Trim(parCampo) = Trim(Me.proDatosProducto.proParametrosProducto.Item(varContador).proCampo) Then
                    Exit For
                End If
            Next varContador
            
            If Me.proDatosProducto.proParametrosProducto.Item(varContador).proObligatorioOT = True Then
                FunFCamposObligatorios = True
            Else
                FunFCamposObligatorios = False
            End If
        End If
    End If
    Exit Function
ErrManager:
    SubGMuestraError
End Function

