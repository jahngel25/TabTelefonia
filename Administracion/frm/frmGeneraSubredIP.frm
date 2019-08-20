VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmGeneraSubredIP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Herramienta de creación de subredes"
   ClientHeight    =   2520
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5955
   Icon            =   "frmGeneraSubredIP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4590
      TabIndex        =   5
      ToolTipText     =   "EliminarTramo"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "&Mostrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1980
      TabIndex        =   3
      ToolTipText     =   "EliminarTramo"
      Top             =   1830
      Width           =   1215
   End
   Begin Threed.SSPanel pnlCampo 
      Height          =   2475
      Left            =   0
      TabIndex        =   6
      Top             =   60
      Width           =   5955
      _Version        =   65536
      _ExtentX        =   10504
      _ExtentY        =   4366
      _StockProps     =   15
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      Begin EDCAdminVoz.EditIPBox EditIPBoxFinal 
         Height          =   315
         Left            =   1020
         TabIndex        =   2
         Top             =   1350
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
      End
      Begin EDCAdminVoz.EditIPBox EditIPBoxInicial 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   930
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
      End
      Begin VB.ComboBox cmbSubred 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   510
         Width           =   2145
      End
      Begin VB.ListBox lstResultado 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   3300
         TabIndex        =   4
         Top             =   510
         Width           =   2505
      End
      Begin Threed.SSPanel pnlTituloCampo 
         Height          =   255
         Left            =   3300
         TabIndex        =   7
         Top             =   240
         Width           =   2505
         _Version        =   65536
         _ExtentX        =   4419
         _ExtentY        =   450
         _StockProps     =   15
         Caption         =   "Generación de subredes IP"
         ForeColor       =   16777215
         BackColor       =   12620376
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "IP Final"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1380
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "IP Inicial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   930
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Subred"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   510
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmGeneraSubredIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmGeneraSubredIP
' Fecha  : 24/09/2004 08:08
' Author    : Germán A. Fajardo G -  Informática & Tecnologia LTDA.
' Propósito   : Herramienta para gweneración asistida de subredes y  direcciones IP
'---------------------------------------------------------------------------------------


Option Explicit

Public proConexion As ADODB.Connection

Public procolValordatos As colValordatos
Public proclaValordatos As claValordatos

Public procolValoresCampoProducto As colValoresCampoProducto
Public proclaValoresCampoProducto As claValoresCampoProducto

Public proCampoSubred As String
Public proCampoIP As String
Public proIdPadre As String
Public proProductNumber As String

Dim iIpsPorGrupo As Integer

Private Sub cmbSubred_Click()
   On Error GoTo ErrorManager

    If cmbSubred.ListIndex > -1 Then
        Me.cmdMostrar.Enabled = True
        lstResultado.Clear
        iIpsPorGrupo = 2 ^ (32 - Val(Right(cmbSubred.List(cmbSubred.ListIndex), 2)))
        Call Me.EditIPBoxInicial.LlenarCombo(0, 256, iIpsPorGrupo, True)
    End If

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdGenerar_Click()
   On Error GoTo ErrorManager

        Call Generar(True)
        Unload Me
      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub cmdMostrar_Click()
   On Error GoTo ErrorManager

    If (EditIPBoxInicial.Text <> "" And EditIPBoxFinal.Text <> "") And (CDbl(EditIPBoxInicial.TextItem(3)) < CDbl(EditIPBoxFinal.TextItem(3)) - 1) Then
        Call Generar(False)
    Else
        MsgBox ("Lasdirecciones IP no son válidas")
    End If

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub EditIPBoxInicial_Change()
   On Error GoTo ErrorManager

    Me.EditIPBoxFinal.TextItem(0) = Me.EditIPBoxInicial.TextItem(0)
    Me.EditIPBoxFinal.TextItem(1) = Me.EditIPBoxInicial.TextItem(1)
    Me.EditIPBoxFinal.TextItem(2) = Me.EditIPBoxInicial.TextItem(2)

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub EditIPBoxInicial_ComboClick()
   On Error GoTo ErrorManager

    Call EditIPBoxFinal.LlenarCombo(Me.EditIPBoxInicial.TextItem(3) - 1, 255, iIpsPorGrupo, False)

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub EditIPBoxInicial_DoLostFocus()
   On Error GoTo ErrorManager

        EditIPBoxFinal.SetFocus

      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Private Sub Form_Load()
    Call IniciarForma
End Sub

Sub IniciarForma()
   On Error GoTo ErrorManager

    Me.cmbSubred.AddItem "/32"
    Me.cmbSubred.AddItem "/31"
    Me.cmbSubred.AddItem "/30"
    Me.cmbSubred.AddItem "/29"
    Me.cmbSubred.AddItem "/28"
    Me.cmbSubred.AddItem "/27"
    Me.cmbSubred.AddItem "/26"
    Me.cmbSubred.AddItem "/25"
    Me.cmbSubred.AddItem "/24"
    Me.EditIPBoxInicial.MostrarCombo = True
    Me.EditIPBoxFinal.MostrarCombo = True
    Me.EditIPBoxFinal.ItemEnabled(0) = False
    Me.EditIPBoxFinal.ItemEnabled(1) = False
    Me.EditIPBoxFinal.ItemEnabled(2) = False
    
      Exit Sub
ErrorManager:
    SubGMuestraError
End Sub

Sub Generar(bGenerarenBase As Boolean)
    Dim iGrupos As Long
    Dim iDiferencia As Long
    Dim iIpsPorGrupo As Long
    Dim varclaValoresCampoProducto As claValoresCampoProducto
    Dim varValordatos As claValor
    Dim iContIP As Long
    Dim tmpiValorPadre As Long
   On Error GoTo ErrorManager
   Screen.MousePointer = 11
    lstResultado.Clear
    iIpsPorGrupo = 2 ^ (32 - Val(Right(cmbSubred.List(cmbSubred.ListIndex), 2)))
    iDiferencia = CDbl(EditIPBoxFinal.TextItem(3)) - IIf(CDbl(EditIPBoxInicial.TextItem(3)) = 0, -1, 0)
    If iIpsPorGrupo > iDiferencia Then
        MsgBox "El Rango no permite generar para este tipo de subred"
        Exit Sub
    End If
        For iGrupos = CDbl(EditIPBoxInicial.TextItem(3)) To CDbl(EditIPBoxFinal.TextItem(3)) Step iIpsPorGrupo
            If bGenerarenBase Then
                
                Set varValordatos = New claValor
                Set varValordatos.proConexion = Me.proConexion
                Set varclaValoresCampoProducto = New claValoresCampoProducto
                Set varclaValoresCampoProducto.proConexion = Me.proConexion
                'Inserta valor de grupo
                varValordatos.proValorId = ""
                varValordatos.proValorDesc = EditIPBoxInicial.TextItem(0) & "." & EditIPBoxInicial.TextItem(1) & "." & Trim(EditIPBoxInicial.TextItem(2)) & "." & iGrupos & Me.cmbSubred.Text
                varValordatos.proRecordStatus = 1
                If varValordatos.MetModificar Then
                End If
                'Inserta Relación del grupo
                tmpiValorPadre = varValordatos.proValorId
                varclaValoresCampoProducto.proProductNumber = proProductNumber
                varclaValoresCampoProducto.proCampo = proCampoSubred
                varclaValoresCampoProducto.proValorId = varValordatos.proValorId
                varclaValoresCampoProducto.proValorIdPadre = Me.proIdPadre
                If Not varclaValoresCampoProducto.MetValidarExistencia Then
                    varclaValoresCampoProducto.MetInsertar
                End If
                For iContIP = iGrupos To iGrupos + iIpsPorGrupo - 1
                    'Inserta valor del IP
                    varValordatos.proValorId = ""
                    varValordatos.proValorDesc = EditIPBoxInicial.TextItem(0) & "." & EditIPBoxInicial.TextItem(1) & "." & Trim(EditIPBoxInicial.TextItem(2)) & "." & Trim(iContIP)
                    varValordatos.proRecordStatus = 1
                    If varValordatos.MetModificar Then
                        'Inserta Relación del IP
                        varclaValoresCampoProducto.proProductNumber = proProductNumber
                        varclaValoresCampoProducto.proCampo = proCampoIP
                        varclaValoresCampoProducto.proValorId = varValordatos.proValorId
                        varclaValoresCampoProducto.proValorIdPadre = tmpiValorPadre
                        If Not varclaValoresCampoProducto.MetValidarExistencia Then
                            varclaValoresCampoProducto.MetInsertar
                        End If
                    End If
                Next
            Else
                Me.lstResultado.AddItem EditIPBoxInicial.TextItem(0) & "." & EditIPBoxInicial.TextItem(1) & "." & Trim(EditIPBoxInicial.TextItem(2)) & "." & iGrupos & Me.cmbSubred.Text
            End If
        Next
        Screen.MousePointer = 0
      Exit Sub
ErrorManager:
    SubGMuestraError
    Screen.MousePointer = 0
End Sub

