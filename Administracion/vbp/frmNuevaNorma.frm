VERSION 5.00
Begin VB.Form frmNuevaNorma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplicación de Normas"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "frmNuevaNorma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Height          =   390
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Guardar la configuración"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   1695
      TabIndex        =   7
      ToolTipText     =   "Cancelar los cambios"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5160
      Top             =   960
   End
   Begin VB.ComboBox cboNormaId 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox cboUsoServicioId 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox cboTipoLineaId 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox cbociudadnombre 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Ciudad de la configuración"
      Top             =   60
      Width           =   3015
   End
   Begin VB.Frame fraestratos 
      Caption         =   "Estratos"
      Height          =   2295
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   5175
      Begin VB.CheckBox chkestrato 
         Caption         =   "chkestrato"
         CausesValidation=   0   'False
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   4095
      End
   End
   Begin VB.ComboBox cboNormaNombre 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Norma a aplicar en la configuración"
      Top             =   580
      Width           =   3015
   End
   Begin VB.ComboBox cboUsoServicioNombre 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Uso del servicio"
      Top             =   1100
      Width           =   3015
   End
   Begin VB.ComboBox cboTipoLineaNombre 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Tipo de línea"
      Top             =   1620
      Width           =   3015
   End
   Begin VB.ComboBox cboCiudadid 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de línea"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Norma"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   640
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ciudad"
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Uso del servicio"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1160
      Width           =   1125
   End
End
Attribute VB_Name = "frmNuevaNorma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public proConexion As ADODB.Connection
Public proEstratos As colEstratoCiudad
Public proNormas As colNormaCiudad
Public proNorma As claNorma
Public proAccion As String
Private Sub cbociudadnombre_click()
  Dim varContador As Integer
    cboNormaId.Clear
    cboNormaNombre.Clear
    If cbociudadnombre.ListIndex > 0 Then
        Set proNormas = New colNormaCiudad
        Set proNormas.proConexion = Me.proConexion
        cboCiudadid.ListIndex = cbociudadnombre.ListIndex
        proNormas.FunGConsulta cboCiudadid.Text, 1
        cboNormaNombre.AddItem "Seleccione una Norma"
        cboNormaId.AddItem "0"
        For varContador = 1 To proNormas.Count
            cboNormaNombre.AddItem proNormas.Item(varContador).proNombreNorma
            cboNormaId.AddItem proNormas.Item(varContador).proNormaCiudadId
            If proAccion = "M" Then
              If proNormas.Item(varContador).proNormaCiudadId = Me.proNorma.proNormaCiudadId Then cboNormaNombre.ListIndex = varContador
            End If
        Next
        If proAccion = "M" Then
            cboNormaId.ListIndex = cboNormaNombre.ListIndex
        Else
            If proNormas.Count = 1 Then
                cboNormaNombre.ListIndex = 1
            Else
                cboNormaNombre.ListIndex = 0
            End If
        End If
    Else
        cboNormaNombre.AddItem "Seleccione una Norma"
        cboNormaId.AddItem "0"
        cboNormaNombre.ListIndex = 0
    End If
End Sub

Private Sub cboNormaNombre_click()
    Timer1.Enabled = True
End Sub

Private Sub cboTipoLineaNombre_click()
    Timer1.Enabled = True
End Sub

Private Sub cboUsoServicioNombre_click()
    Timer1.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    Dim varciudad As Long
    Dim varTipoLinea As Long
    Dim varNorma As Long
    Dim varUsoServicio As Long
    Dim varContador  As Integer
    cboCiudadid.ListIndex = cbociudadnombre.ListIndex
    cboNormaId.ListIndex = cboNormaNombre.ListIndex
    cboTipoLineaId.ListIndex = cboTipoLineaNombre.ListIndex
    cboUsoServicioId.ListIndex = cboUsoServicioNombre.ListIndex
    varciudad = cboCiudadid.Text
    varTipoLinea = cboTipoLineaId.Text
    varNorma = IIf(cboNormaId.Text = "", 0, cboNormaId.Text)
    varUsoServicio = IIf(cboUsoServicioId.Text = "", 0, cboUsoServicioId.Text)
    If varciudad = 0 Or varTipoLinea = 0 Or varNorma = 0 Or varUsoServicio = 0 Then
       MsgBox "No se ha definido la totalidad de elementos", vbInformation, "Seleccione un valor para cada campo"
    Else
        If proAccion = "N" Then
            Set proNorma = New claNorma
            Set proNorma.proConexion = Me.proConexion
            proNorma.proCiudadId = varciudad
            proNorma.proNormaCiudadId = varNorma
            proNorma.proTipoLineaId = varTipoLinea
            proNorma.proUsoServicioId = varUsoServicio
            proNorma.FunGInsertar
            If proNorma.proNormaId <> 0 Then
                For varContador = 1 To Me.chkestrato.Count - 1
                    If chkestrato(varContador).Value = 1 Then
                        proNorma.FunGInsertarEstrato (chkestrato(varContador).Tag)
                    End If
                Next
                MsgBox "Norma configurada con éxito", vbInformation
                Unload Me
            End If
        Else
            For varContador = 1 To Me.chkestrato.Count - 1
                Dim varestrato As claEstratoCiudad
                Set varestrato = proEstratos.Item(varContador)
                If chkestrato(varContador).Value = 1 And varestrato.proSeleccionado = False Then
                    proNorma.FunGInsertarEstrato chkestrato(varContador).Tag
                End If
                If chkestrato(varContador).Value = 0 And varestrato.proSeleccionado = True Then
                    proNorma.FunGEliminarEstrato chkestrato(varContador).Tag
                End If
            Next
            MsgBox "Modificación realizada con éxito", vbInformation
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Activate()
    If proAccion = "M" Then
        Me.cmdGuardar.SetFocus
    Else
        cbociudadnombre.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim varContador As Integer
    Dim varciudades As colCiudadOnyx
    Set varciudades = New colCiudadOnyx
    Set varciudades.proConexion = Me.proConexion
    Dim varciudad As Long
    varciudades.FunGConsulta
    If proAccion = "M" Then
       varciudad = Me.proNorma.proCiudadId
    End If
    FunGLlenarCombosCiudad cboCiudadid, cbociudadnombre, varciudades, "Seleccione una ciudad", varciudad
    If proAccion = "M" Then
        cbociudadnombre.Enabled = False
    End If
    Dim varTipoLinea As colValoresCampoProducto
    Set varTipoLinea = New colValoresCampoProducto
    Set varTipoLinea.proConexion = Me.proConexion
    varTipoLinea.proProductNumber = "1810"
    'Consulta de tipo de línea
    varTipoLinea.proCampo = "vchuser1"
    varTipoLinea.MetConsultarValoresxProducto
    Me.cboTipoLineaNombre.Clear
    Me.cboTipoLineaId.Clear
    cboTipoLineaNombre.AddItem "Seleccione un Tipo"
    cboTipoLineaId.AddItem "0"
    For varContador = 1 To varTipoLinea.Count
        cboTipoLineaNombre.AddItem varTipoLinea.Item(varContador).proValorDesc
        cboTipoLineaId.AddItem varTipoLinea.Item(varContador).proValorId
        If proAccion = "M" Then
            If Me.proNorma.proTipoLineaId = varTipoLinea.Item(varContador).proValorId Then cboTipoLineaId.ListIndex = varContador
        End If
    Next
    Dim varUsoServicio As colEstratos
    Set varUsoServicio = New colEstratos
    Set varUsoServicio.proConexion = Me.proConexion
    varUsoServicio.MetConsultar
    'Consulta de tipo de línea
    Me.cboUsoServicioNombre.Clear
    Me.cboUsoServicioId.Clear
    cboUsoServicioNombre.AddItem "Seleccione un Uso de Servicio"
    cboUsoServicioId.AddItem "0"
    For varContador = 1 To varUsoServicio.Count
       cboUsoServicioNombre.AddItem varUsoServicio.Item(varContador).proDescripcion
       cboUsoServicioId.AddItem varUsoServicio.Item(varContador).proEstratoID
       If proAccion = "M" Then
            If Me.proNorma.proUsoServicioId = varUsoServicio.Item(varContador).proEstratoID Then cboUsoServicioId.ListIndex = varContador
        End If
    Next
    If proAccion = "M" Then
        cboTipoLineaNombre.ListIndex = cboTipoLineaId.ListIndex
        cboUsoServicioNombre.ListIndex = cboUsoServicioId.ListIndex
        cboUsoServicioNombre.Enabled = False
        cboTipoLineaNombre.Enabled = False
        cboNormaNombre.Enabled = False
    Else
        If varTipoLinea.Count = 1 Then
          cboTipoLineaNombre.ListIndex = 1
        Else
          cboTipoLineaNombre.ListIndex = 0
        End If
        If varUsoServicio.Count = 1 Then
          cboUsoServicioNombre.ListIndex = 1
        Else
          cboUsoServicioNombre.ListIndex = 0
        End If
    End If
End Sub
Private Sub MetConsultarEstrato()
    Dim varciudad As Long
    Dim varTipoLinea As Long
    Dim varNorma As Long
    Dim varUsoServicio As Long
    Dim varContador  As Integer
    cboCiudadid.ListIndex = cbociudadnombre.ListIndex
    cboNormaId.ListIndex = cboNormaNombre.ListIndex
    cboTipoLineaId.ListIndex = cboTipoLineaNombre.ListIndex
    cboUsoServicioId.ListIndex = cboUsoServicioNombre.ListIndex
    varciudad = cboCiudadid.Text
    varTipoLinea = cboTipoLineaId.Text
    varNorma = IIf(cboNormaId.Text = "", 0, cboNormaId.Text)
    varUsoServicio = IIf(cboUsoServicioId.Text = "", 0, cboUsoServicioId.Text)
    Dim i, j As Integer
    i = Me.chkestrato.Count
    For j = 1 To i - 1
      Unload chkestrato(j)
    Next
    Me.fraestratos.Refresh
    If varciudad = 0 Or varTipoLinea = 0 Or varNorma = 0 Or varUsoServicio = 0 Then
    Else
        Set proEstratos = New colEstratoCiudad
        Set proEstratos.proConexion = Me.proConexion
        If proAccion <> "M" Then
            varNorma = 0
        Else
            varNorma = proNorma.proNormaId
        End If
        proEstratos.FunGConsultaNorma varNorma, varTipoLinea, varUsoServicio, varciudad
        For varContador = 1 To proEstratos.Count
            With proEstratos.Item(varContador)
                i = Me.chkestrato.Count
                Load Me.chkestrato(i)
                chkestrato(i).Caption = .proNombreEstrato
                chkestrato(i).Tag = .proEstratoCiudadId
                chkestrato(i).Value = IIf(.proSeleccionado = True, 1, 0)
                chkestrato(i).Top = chkestrato(i - 1).Top + chkestrato(i - 1).Height + 20
                chkestrato(i).Visible = True
            End With
        Next
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    MetConsultarEstrato
End Sub
