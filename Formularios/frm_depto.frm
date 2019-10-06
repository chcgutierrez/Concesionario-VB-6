VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDepto 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestra - Departamento"
   ClientHeight    =   5790
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6390
   Icon            =   "frm_depto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesPais 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2610
      TabIndex        =   16
      Top             =   650
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   1080
   End
   Begin VB.TextBox txtObser 
      Height          =   735
      Left            =   1755
      TabIndex        =   6
      Top             =   2700
      Width           =   4380
   End
   Begin VB.TextBox txtNomDepto 
      Height          =   315
      Left            =   1755
      TabIndex        =   3
      Top             =   1550
      Width           =   3735
   End
   Begin VB.TextBox txtDepto 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1755
      TabIndex        =   1
      Top             =   1080
      Width           =   855
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
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox txtCodPais 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1755
      TabIndex        =   0
      Top             =   650
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid dtgDepto 
      Height          =   1455
      Left            =   375
      TabIndex        =   9
      Top             =   3720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2566
      _Version        =   393216
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
         DataField       =   "cod_pais"
         Caption         =   "Cod Pais"
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
         DataField       =   "cod_depto"
         Caption         =   "Cod. Depto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "nom_depto"
         Caption         =   "Nombre Depto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "est_depto"
         Caption         =   "Estado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "fec_act"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   9226
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "obs_gen"
         Caption         =   "Obervaciones"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
      EndProperty
   End
   Begin VB.OptionButton optInactivo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Inactivo"
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   2180
      Width           =   900
   End
   Begin VB.OptionButton optActivo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Activo"
      Height          =   315
      Left            =   1850
      TabIndex        =   4
      Top             =   2180
      Width           =   780
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
      Left            =   1755
      TabIndex        =   12
      Top             =   1950
      Width           =   2055
   End
   Begin MSComctlLib.ImageList img_lista 
      Left            =   4440
      Top             =   960
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
            Picture         =   "frm_depto.frx":058A
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_depto.frx":0B24
            Key             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_depto.frx":10BE
            Key             =   "borrar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_depto.frx":1658
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_depto.frx":1BF2
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_depto.frx":218C
            Key             =   "guardar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_depto.frx":2726
            Key             =   "imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_depto.frx":2CC0
            Key             =   "salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb_botones 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6390
      _ExtentX        =   11271
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
            Object.Tag             =   "Nuevo"
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
      TabIndex        =   15
      Top             =   5415
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Ver 1.0.0"
            TextSave        =   "05/10/2019"
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
            Object.Width           =   3519
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Existentes"
      DataSource      =   "360"
      Height          =   195
      Left            =   375
      TabIndex        =   14
      Top             =   3480
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   4860
      Left            =   120
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
      Left            =   600
      TabIndex        =   11
      Top             =   2700
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Departamento"
      Height          =   195
      Left            =   675
      TabIndex        =   10
      Top             =   1550
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. Departamento"
      Height          =   195
      Left            =   285
      TabIndex        =   8
      Top             =   1080
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pais"
      Height          =   195
      Left            =   1350
      TabIndex        =   7
      Top             =   650
      Width           =   300
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
Attribute VB_Name = "frmDepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytFlagModifica As Byte

Private Sub cmdValidar_Click()
    If Me.txtCodPais.Text <> "" And Me.txtDepto.Text <> "" Then
        mnuArchivo_Buscar_Click
    Else
        MsgBox "Datos Incompletos", vbInformation + vbOKOnly, "Consultar"
    End If
End Sub

Private Sub Form_Load()
    AbrirDepto
    mnuArchivo_Nuevo_Click
    PrenderMenus Me, tlb_botones, gcnstConsCompleta
End Sub

Private Sub mnuArchivo_Buscar_Click()
On Error GoTo ControlError

With cmdSQL
.ActiveConnection = ConexSQL
.CommandType = adCmdStoredProc
.CommandText = "sp_buscar_depto"
.Parameters.Append .CreateParameter("@codPais", adVarChar, adParamInput, 10, Me.txtCodPais.Text)
.Parameters.Append .CreateParameter("@codDepto", adVarChar, adParamInput, 10, Me.txtDepto.Text)
Set rstSQL = cmdSQL.Execute
End With
Set cmdSQL = Nothing
Set cmdSQL.ActiveConnection = Nothing
'cmdSQL.ActiveConnection.Close
If rstSQL.RecordCount > 0 Then
    If MsgBox("El registro ya existe. ¿Mostrar Datos?", vbQuestion + vbYesNo, "Consultar") = vbYes Then
        Me.txtNomDepto.Text = rstSQL("nom_depto").Value
            If rstSQL("est_depto").Value = "A" Then
                Me.optActivo = True
            Else
                Me.optInactivo = True
            End If
        Me.txtObser.Text = rstSQL("obs_gen").Value
    End If
    Set rstSQL = Nothing
Else
Me.txtCodPais.Enabled = False
Me.txtDepto.Enabled = False
Me.cmdValidar.Enabled = False
Me.txtNomDepto.Enabled = True
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

Me.txtCodPais.Enabled = False
Me.txtDepto.Enabled = False
Me.txtNomDepto.Enabled = False
Me.optActivo.Enabled = False
Me.optInactivo.Enabled = False
Me.txtObser.Enabled = False

ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

Private Sub mnuArchivo_Editar_Click()
On Error GoTo ControlError

Me.txtCodPais.Enabled = False
Me.txtDepto.Enabled = False
Me.cmdValidar.Enabled = False
Me.txtNomDepto.Enabled = True
Me.optActivo.Enabled = True
Me.optInactivo.Enabled = True
Me.txtObser.Enabled = True
bytFlagModifica = 1

ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

Private Sub mnuArchivo_Guardar_Click()

Dim estDepto As String
Dim rsPais As ADODB.Recordset

On Error GoTo ControlError

If Me.optActivo = True Then
estDepto = "A"
ElseIf Me.optInactivo = True Then
estDepto = "I"
End If

    If Len(Me.txtCodPais.Text) > 0 Then
        Set rsPais = TraerPais(txtCodPais.Text)
    End If

If bytFlagModifica = 0 Then

With cmdSQL
.ActiveConnection = ConexSQL
.CommandType = adCmdStoredProc
.CommandText = "sp_guardar_depto"
    .Parameters.Append .CreateParameter("@idPais", adInteger, adParamInput, 10, rsPais("id_pais").Value)
    .Parameters.Append .CreateParameter("@codDepto", adVarChar, adParamInput, 10, Me.txtDepto.Text)
    .Parameters.Append .CreateParameter("@nomDepto", adVarChar, adParamInput, 100, Me.txtNomDepto.Text)
    .Parameters.Append .CreateParameter("@estDepto", adVarChar, adParamInput, 10, estDepto)
    .Parameters.Append .CreateParameter("@obsGen", adVarChar, adParamInput, 100, Me.txtObser.Text)
.Execute
End With
Set cmdSQL = Nothing
Set cmdSQL.ActiveConnection = Nothing
'cmdSQL.ActiveConnection.Close
mnuArchivo_Cancelar_Click
MsgBox "Datos Guardados Correctamente", vbInformation + vbOKOnly, "Guardar"
AbrirDepto

Else

With cmdSQL
.ActiveConnection = ConexSQL
.CommandType = adCmdStoredProc
.CommandText = "sp_editar_depto"
    .Parameters.Append .CreateParameter("@idPais", adInteger, adParamInput, 10, rsPais("id_pais").Value)
    .Parameters.Append .CreateParameter("@codDepto", adVarChar, adParamInput, 10, Me.txtDepto.Text)
    .Parameters.Append .CreateParameter("@nomDepto", adVarChar, adParamInput, 100, Me.txtNomDepto.Text)
    .Parameters.Append .CreateParameter("@estDepto", adVarChar, adParamInput, 10, estDepto)
    .Parameters.Append .CreateParameter("@obsGen", adVarChar, adParamInput, 100, Me.txtObser.Text)
.Execute
End With
Set cmdSQL = Nothing
Set cmdSQL.ActiveConnection = Nothing
'cmdSQL.ActiveConnection.Close
mnuArchivo_Cancelar_Click
MsgBox "Datos Guardados Correctamente", vbInformation + vbOKOnly, "Guardar"
AbrirDepto

End If

Set rsPais = Nothing

ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub


Private Sub mnuArchivo_Nuevo_Click()
On Error GoTo ControlError

Me.txtCodPais.Text = ""
Me.txtCodPais.Enabled = True
Me.txtDesPais.Enabled = False
Me.txtDesPais.Text = ""
Me.txtDepto.Text = ""
Me.txtDepto.Enabled = True
Me.cmdValidar.Enabled = True
Me.txtNomDepto.Text = ""
Me.txtNomDepto.Enabled = False
Me.txtObser.Text = ""
Me.txtObser.Enabled = False
Me.optActivo.Enabled = False
Me.optActivo.Value = False
Me.optInactivo.Enabled = False
Me.optInactivo.Value = False
bytFlagModifica = 0

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
        
        Case "btnEditar": mnuArchivo_Editar_Click
        
        Case "btnGuardar": mnuArchivo_Guardar_Click
        
        Case "btnSalir": mnuArchivo_Salir_Click

    End Select
End Sub

Private Sub txtCodPais_Validate(Cancel As Boolean)

    Dim rsPais As ADODB.Recordset
    
    On Error GoTo ControlError
    
    If Len(Me.txtCodPais.Text) > 0 Then
        Set rsPais = TraerPais(txtCodPais.Text)
            If rsPais.RecordCount > 0 Then
                    If rsPais("est_pais").Value = "I" Then
                        Me.txtCodPais.SelStart = 0
                        Me.txtCodPais.SelLength = Len(Me.txtCodPais.Text)
                        MsgBox "El Pais ingresado está inactivo.", vbOKOnly, "Buscar Pais"
                        Me.txtDesPais.Text = ""
                        Cancel = True
                        Exit Sub
                    End If
                Me.txtDesPais.Text = rsPais("nom_pais").Value
            Else
                Me.txtCodPais.SelStart = 0
                Me.txtCodPais.SelLength = Len(Me.txtCodPais.Text)
                MsgBox "No existe el Pais para el criterio ingresado.", vbOKOnly, "Buscar Pais"
                Me.txtDesPais.Text = ""
                Cancel = True
                Exit Sub
            End If
    Else
         Me.txtDesPais.Text = ""
         MsgBox "Debe ingresar un criterio para realizar la busqueda.", vbOKOnly, "Criterio Inválido"
         Me.txtCodPais.SetFocus
         Cancel = True
         Exit Sub
    End If
    
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc

End Sub

Private Sub AbrirDepto()
On Error GoTo ControlError

With cmdSQL
.ActiveConnection = ConexSQL
.CommandType = adCmdStoredProc
.CommandText = "sp_mostrar_depto"
    With rstSQL
        If .State = 1 Then .Close
        Set rstSQL = cmdSQL.Execute
        Set Me.dtgDepto.DataSource = rstSQL
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
