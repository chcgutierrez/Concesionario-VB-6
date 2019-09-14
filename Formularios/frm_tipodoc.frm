VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTipoDoc 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestra - Tipo Documento"
   ClientHeight    =   5385
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6285
   Icon            =   "frm_tipodoc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dtgTipoDoc 
      Height          =   1455
      Left            =   300
      TabIndex        =   12
      Top             =   3320
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "id_tipodoc"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "tipo_doc"
         Caption         =   "Tipo Doc."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "des_tip_doc"
         Caption         =   "Descripcion"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "est_tip_doc"
         Caption         =   "Estado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "fec_act"
         Caption         =   "Modificado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "obs_gen"
         Caption         =   "Observaciones"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
            WrapText        =   -1  'True
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
            WrapText        =   -1  'True
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
            WrapText        =   -1  'True
            ColumnWidth     =   3495,118
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
            WrapText        =   -1  'True
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column04 
            ColumnAllowSizing=   0   'False
            WrapText        =   -1  'True
         EndProperty
         BeginProperty Column05 
            ColumnAllowSizing=   0   'False
            WrapText        =   -1  'True
            ColumnWidth     =   1500,095
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   1680
   End
   Begin VB.TextBox txtObser 
      Height          =   735
      Left            =   1600
      TabIndex        =   8
      Text            =   "CARGA_INICIAL"
      Top             =   2250
      Width           =   4380
   End
   Begin VB.OptionButton optActivo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Activo"
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   1695
      Width           =   780
   End
   Begin VB.OptionButton optInactivo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Inactivo"
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   1695
      Width           =   900
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   ">>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtDescTipoDoc 
      Height          =   315
      Left            =   1635
      TabIndex        =   2
      Text            =   "PERMISO ESPECIAL DE PERMANENCIA"
      Top             =   1050
      Width           =   3735
   End
   Begin VB.TextBox txtTipoDoc 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1635
      TabIndex        =   0
      Text            =   "PE"
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   1455
      Width           =   2055
   End
   Begin MSComctlLib.ImageList img_lista 
      Left            =   4560
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_tipodoc.frx":058A
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_tipodoc.frx":0B24
            Key             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_tipodoc.frx":10BE
            Key             =   "borrar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_tipodoc.frx":1658
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_tipodoc.frx":1BF2
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_tipodoc.frx":218C
            Key             =   "guardar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_tipodoc.frx":2726
            Key             =   "imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_tipodoc.frx":2CC0
            Key             =   "salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb_botones 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "img_lista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnNuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnEditar"
            Object.ToolTipText     =   "Editar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnBorrar"
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnCancelar"
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnBuscar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnGuardar"
            Object.ToolTipText     =   "Guardar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnImprimir"
            Object.ToolTipText     =   "Reporte"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnSalir"
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   5010
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Ver 1.0.0"
            TextSave        =   "13/09/2019"
            Key             =   "sbrPan01"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
            Key             =   "sbrPan02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3334
            Key             =   "sbrPan03"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Ver 1.0.0"
            TextSave        =   "Ver 1.0.0"
            Key             =   "sbrPan04"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Existentes"
      DataSource      =   "360"
      Height          =   195
      Left            =   320
      TabIndex        =   13
      Top             =   3080
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   4440
      Left            =   60
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
      DataSource      =   "360"
      Height          =   195
      Left            =   430
      TabIndex        =   9
      Top             =   2250
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Descr. Documento"
      Height          =   195
      Left            =   195
      TabIndex        =   3
      Top             =   1050
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      Height          =   195
      Left            =   705
      TabIndex        =   1
      Top             =   600
      Width           =   825
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuArchivo_Nuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnuArchivo_Cancelar 
         Caption         =   "&Cancelar"
      End
      Begin VB.Menu mnuArchivo_Guardar 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu mnuArchivo_Buscar 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnuArchivo_Editar 
         Caption         =   "&Editar"
      End
      Begin VB.Menu mnuArchivo_Cargar 
         Caption         =   "&Desde &Archivo"
      End
      Begin VB.Menu mnuArchivo_Imprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu mnuArchivo_Salir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
   End
   Begin VB.Menu mnuVista 
      Caption         =   "&Vista"
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
   End
End
Attribute VB_Name = "frmTipoDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytFlagModifica As Byte

Private Sub cmdValidar_Click()
    If Me.txtTipoDoc.Text <> "" Then
        mnuArchivo_Buscar_Click
    Else
        MsgBox "Debe ingresar un criterio", vbInformation + vbOKOnly, "Consultar"
        Me.txtTipoDoc.SetFocus
    End If
End Sub

Private Sub Form_Load()
    mnuArchivo_Nuevo_Click
    PrenderMenus Me, tlb_botones, 275
'    AbrirTipodoc
End Sub

Private Sub mnuArchivo_Buscar_Click()
On Error GoTo ControlError

With cmdSQL
.ActiveConnection = ConexSQL
.CommandType = adCmdStoredProc
.CommandText = "sp_buscar_tipodoc"
.Parameters.Refresh
    .Parameters("@tipo_doc").Value = Me.txtTipoDoc.Text
Set rstSQL = cmdSQL.Execute
End With
Set cmdSQL = Nothing
Set cmdSQL.ActiveConnection = Nothing
'cmdSQL.ActiveConnection.Close
If rstSQL.RecordCount > 0 Then
    If MsgBox("El registro ya existe. ¿Mostrar Datos?", vbQuestion + vbYesNo, "Consultar") = vbYes Then
        Me.txtDescTipoDoc.Text = rstSQL("des_tip_doc").Value
            If rstSQL("est_tip_doc").Value = "A" Then
                Me.optActivo = True
            Else
                Me.optInactivo = True
            End If
        Me.txtObser.Text = rstSQL("obs_gen").Value
    End If
    Set rstSQL = Nothing
Else
Me.txtTipoDoc.Enabled = False
Me.cmdValidar.Enabled = False
Me.txtDescTipoDoc.Enabled = True
Me.optActivo.Enabled = True
Me.optInactivo.Enabled = True
Me.txtObser.Enabled = True
End If
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

Private Sub mnuArchivo_Cancelar_Click()
On Error GoTo ControlError

Me.txtTipoDoc.Enabled = False
Me.txtDescTipoDoc.Enabled = False
Me.txtObser.Enabled = False
Me.optActivo.Enabled = False
Me.optInactivo.Enabled = False

ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

Private Sub mnuArchivo_Guardar_Click()
On Error GoTo ControlError

Dim estTipoDoc As String

If Me.optActivo = True Then
estTipoDoc = "A"
ElseIf Me.optInactivo = True Then
estTipoDoc = "I"
End If

With cmdSQL
.ActiveConnection = ConexSQL
.CommandType = adCmdStoredProc
.CommandText = "sp_guardar_tipodoc"
.Parameters.Refresh
    .Parameters("@tipo_doc").Value = Me.txtTipoDoc.Text
    .Parameters("@des_tip_doc").Value = Me.txtDescTipoDoc.Text
    .Parameters("@est_tip_doc").Value = estTipoDoc
    .Parameters("@obs_gen").Value = Me.txtObser.Text
.Execute
End With
Set cmdSQL = Nothing
Set cmdSQL.ActiveConnection = Nothing
'cmdSQL.ActiveConnection.Close
mnuArchivo_Cancelar_Click
MsgBox "Datos Guardados Correctamente", vbInformation + vbOKOnly, "Guardar"
AbrirTipodoc
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

Private Sub mnuArchivo_Nuevo_Click()

On Error GoTo ControlError

Me.txtTipoDoc.Text = ""
Me.txtTipoDoc.Enabled = True
Me.cmdValidar.Enabled = True
Me.txtDescTipoDoc.Text = ""
Me.txtDescTipoDoc.Enabled = False
Me.txtObser.Text = ""
Me.txtObser.Enabled = False
Me.optActivo.Enabled = False
Me.optActivo.Value = False
Me.optInactivo.Enabled = False
Me.optInactivo.Value = False

ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc

End Sub

Private Sub mnuArchivo_Salir_Click()
    If MsgBox("¿Cerrar el Formulario?", vbQuestion + vbYesNo, "Cerrar") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub tlb_botones_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key

        Case "btnNuevo": mnuArchivo_Nuevo_Click
        
        Case "btnGuardar": mnuArchivo_Guardar_Click
        
        Case "btnSalir": mnuArchivo_Salir_Click

    End Select

End Sub

Private Sub AbrirTipodoc()
On Error GoTo ControlError

With cmdSQL
.ActiveConnection = ConexSQL
.CommandType = adCmdStoredProc
.CommandText = "sp_mostrar_tipodoc"
    With rstSQL
        If .State = 1 Then .Close
        Set rstSQL = cmdSQL.Execute
        Set Me.dtgTipoDoc.DataSource = rstSQL
    End With
End With
Set cmdSQL = Nothing
Set cmdSQL.ActiveConnection = Nothing
'cmdSQL.ActiveConnection.Close
Set rstSQL = Nothing
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc

End Sub
