VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
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
            TextSave        =   "30/04/2020"
            Key             =   "sbrPan01"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "MAYï¿½S"
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
100   If Me.txtCodPais.Text <> "" And Me.txtDepto.Text <> "" Then
110      mnuArchivo_Buscar_Click
120   Else
130      MsgBox "Datos Incompletos", vbInformation + vbOKOnly, "Consultar"
140   End If
End Sub

Private Sub Form_Load()
100   AbrirDepto
110   mnuArchivo_Nuevo_Click
120   PrenderMenus Me, tlb_botones, gcnstConsCompleta
End Sub

Private Sub mnuArchivo_Buscar_Click()
1000   On Error GoTo ControlError
       
1010   With cmdSQL
1020      .ActiveConnection = ConexSQL
1030      .CommandType = adCmdStoredProc
1040      .CommandText = "sp_buscar_depto"
1050      .Parameters.Append .CreateParameter("@codPais", adVarChar, adParamInput, 10, Me.txtCodPais.Text)
1060      .Parameters.Append .CreateParameter("@codDepto", adVarChar, adParamInput, 10, Me.txtDepto.Text)
1070      Set rstSQL = cmdSQL.Execute
1080   End With
       
1090   Set cmdSQL = Nothing
1100   Set cmdSQL.ActiveConnection = Nothing
       '1110   cmdSQL.ActiveConnection.Close
       
1110   If rstSQL.RecordCount > 0 Then
1120      If MsgBox("El registro ya existe. ¿Mostrar Datos?", vbQuestion + vbYesNo, "Consultar") = vbYes Then
1130         Me.txtNomDepto.Text = rstSQL("nom_depto").Value
1140         If rstSQL("est_depto").Value = "A" Then
1150            Me.optActivo = True
1160         Else
1170            Me.optInactivo = True
1180         End If
1190         Me.txtObser.Text = rstSQL("obs_gen").Value
1200      End If
1210      Set rstSQL = Nothing
1220   Else
1230      Me.txtCodPais.Enabled = False
1240      Me.txtDepto.Enabled = False
1250      Me.cmdValidar.Enabled = False
1260      Me.txtNomDepto.Enabled = True
1270      Me.optActivo.Enabled = True
1280      Me.optInactivo.Enabled = True
1290      Me.txtObser.Enabled = True
1300   End If
ExitProc:
1310   Exit Sub
ControlError:
1320   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1330   Resume ExitProc
       
End Sub

Private Sub mnuArchivo_Cancelar_Click()
1000   On Error GoTo ControlError
       
1010   Me.txtCodPais.Enabled = False
1020   Me.txtDepto.Enabled = False
1030   Me.txtNomDepto.Enabled = False
1040   Me.optActivo.Enabled = False
1050   Me.optInactivo.Enabled = False
1060   Me.txtObser.Enabled = False
       
ExitProc:
1070   Exit Sub
ControlError:
1080   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1090   Resume ExitProc
End Sub

Private Sub mnuArchivo_Editar_Click()
1000   On Error GoTo ControlError
       
1010   Me.txtCodPais.Enabled = False
1020   Me.txtDepto.Enabled = False
1030   Me.cmdValidar.Enabled = False
1040   Me.txtNomDepto.Enabled = True
1050   Me.optActivo.Enabled = True
1060   Me.optInactivo.Enabled = True
1070   Me.txtObser.Enabled = True
1080   bytFlagModifica = 1
       
ExitProc:
1090   Exit Sub
ControlError:
1100   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1110   Resume ExitProc
End Sub

Private Sub mnuArchivo_Guardar_Click()
       
       Dim estDepto As String
       Dim rsPais As ADODB.Recordset
       
1000   On Error GoTo ControlError
       
1010   If Me.optActivo = True Then
1020      estDepto = "A"
1030   ElseIf Me.optInactivo = True Then
1040      estDepto = "I"
1050   End If
       
1060   If Len(Me.txtCodPais.Text) > 0 Then
1070      Set rsPais = TraerPais(txtCodPais.Text)
1080   End If
       
1090   If bytFlagModifica = 0 Then
          
1100      With cmdSQL
1110         .ActiveConnection = ConexSQL
1120         .CommandType = adCmdStoredProc
1130         .CommandText = "sp_guardar_depto"
1140         .Parameters.Append .CreateParameter("@idPais", adInteger, adParamInput, 10, rsPais("id_pais").Value)
1150         .Parameters.Append .CreateParameter("@codDepto", adVarChar, adParamInput, 10, Me.txtDepto.Text)
1160         .Parameters.Append .CreateParameter("@nomDepto", adVarChar, adParamInput, 100, Me.txtNomDepto.Text)
1170         .Parameters.Append .CreateParameter("@estDepto", adVarChar, adParamInput, 10, estDepto)
1180         .Parameters.Append .CreateParameter("@obsGen", adVarChar, adParamInput, 100, Me.txtObser.Text)
1190         .Execute
1200      End With
1210      Set cmdSQL = Nothing
1220      Set cmdSQL.ActiveConnection = Nothing
          '1230      cmdSQL.ActiveConnection.Close
1230      mnuArchivo_Cancelar_Click
1240      MsgBox "Datos Guardados Correctamente", vbInformation + vbOKOnly, "Guardar"
1250      AbrirDepto
          
1260   Else
          
1270      With cmdSQL
1280         .ActiveConnection = ConexSQL
1290         .CommandType = adCmdStoredProc
1300         .CommandText = "sp_editar_depto"
1310         .Parameters.Append .CreateParameter("@idPais", adInteger, adParamInput, 10, rsPais("id_pais").Value)
1320         .Parameters.Append .CreateParameter("@codDepto", adVarChar, adParamInput, 10, Me.txtDepto.Text)
1330         .Parameters.Append .CreateParameter("@nomDepto", adVarChar, adParamInput, 100, Me.txtNomDepto.Text)
1340         .Parameters.Append .CreateParameter("@estDepto", adVarChar, adParamInput, 10, estDepto)
1350         .Parameters.Append .CreateParameter("@obsGen", adVarChar, adParamInput, 100, Me.txtObser.Text)
1360         .Execute
1370      End With
1380      Set cmdSQL = Nothing
1390      Set cmdSQL.ActiveConnection = Nothing
          '1400      cmdSQL.ActiveConnection.Close
1400      mnuArchivo_Cancelar_Click
1410      MsgBox "Datos Guardados Correctamente", vbInformation + vbOKOnly, "Guardar"
1420      AbrirDepto
          
1430   End If
       
1440   Set rsPais = Nothing
       
ExitProc:
1450   Exit Sub
ControlError:
1460   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1470   Resume ExitProc
End Sub

Private Sub mnuArchivo_Nuevo_Click()

1000   On Error GoTo ControlError
       
1010   Me.txtCodPais.Text = ""
1020   Me.txtCodPais.Enabled = True
1030   Me.txtDesPais.Enabled = False
1040   Me.txtDesPais.Text = ""
1050   Me.txtDepto.Text = ""
1060   Me.txtDepto.Enabled = True
1070   Me.cmdValidar.Enabled = True
1080   Me.txtNomDepto.Text = ""
1090   Me.txtNomDepto.Enabled = False
1100   Me.txtObser.Text = ""
1110   Me.txtObser.Enabled = False
1120   Me.optActivo.Enabled = False
1130   Me.optActivo.Value = False
1140   Me.optInactivo.Enabled = False
1150   Me.optInactivo.Value = False
1160   bytFlagModifica = 0
       
ExitProc:
1170   Exit Sub
ControlError:
1180   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1190   Resume ExitProc
End Sub

Private Sub mnuArchivo_Salir_Click()
100   If MsgBox("¿Cerrar el Formulario?", vbQuestion + vbYesNo, "Cerrar") = vbYes Then
110      Unload Me
120   End If
End Sub

Private Sub tlb_botones_ButtonClick(ByVal Button As MSComctlLib.Button)
1000   Select Case Button.Key
          
       Case "btnNuevo": mnuArchivo_Nuevo_Click
          
       Case "btnEditar": mnuArchivo_Editar_Click
          
       Case "btnGuardar": mnuArchivo_Guardar_Click
          
       Case "btnSalir": mnuArchivo_Salir_Click
          
1010   End Select
End Sub

Private Sub txtCodPais_DblClick()
       
       Dim blnMostrarDat As Boolean
       Dim strCodPais As String
       Dim strDescPais As String
       
1000   On Error GoTo ControlError
       
1010   blnMostrarDat = frm_bPais.BusquedaPais(strCodPais, strDescPais)
1020   txtCodPais.Text = strCodPais
1030   txtDesPais = strDescPais
1040   If Len(txtDesPais.Text) > 0 Then
1050      txtDepto.SetFocus
1060   End If
1070   Me.Refresh
1080   Exit Sub
       
ExitProc:
1090   Exit Sub
ControlError:
1100   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1110   Resume ExitProc
End Sub

Private Sub txtCodPais_Validate(Cancel As Boolean)
       
       Dim rsPais As ADODB.Recordset
       
1000   On Error GoTo ControlError
       
1010   If Len(Me.txtCodPais.Text) > 0 Then
1020      Set rsPais = TraerPais(txtCodPais.Text)
1030      If rsPais.RecordCount > 0 Then
1040         If rsPais("est_pais").Value = "I" Then
1050            Me.txtCodPais.SelStart = 0
1060            Me.txtCodPais.SelLength = Len(Me.txtCodPais.Text)
1070            MsgBox "El Pais ingresado está inactivo.", vbOKOnly, "Buscar Pais"
1080            Me.txtDesPais.Text = ""
1090            Cancel = True
1100            Exit Sub
1110         End If
1120         Me.txtDesPais.Text = rsPais("nom_pais").Value
1130      Else
1140         Me.txtCodPais.SelStart = 0
1150         Me.txtCodPais.SelLength = Len(Me.txtCodPais.Text)
1160         MsgBox "No existe el Pais para el criterio ingresado.", vbOKOnly, "Buscar Pais"
1170         Me.txtDesPais.Text = ""
1180         Cancel = True
1190         Exit Sub
1200      End If
1210   Else
1220      Me.txtDesPais.Text = ""
1230      MsgBox "Debe ingresar un criterio para realizar la busqueda.", vbOKOnly, "Criterio Inválido"
1240      Me.txtCodPais.SetFocus
1250      Cancel = True
1260      Exit Sub
1270   End If
       
ExitProc:
1280   Exit Sub
ControlError:
1290   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1300   Resume ExitProc
       
End Sub

Private Sub AbrirDepto()
1000   On Error GoTo ControlError
       
1010   With cmdSQL
1020      .ActiveConnection = ConexSQL
1030      .CommandType = adCmdStoredProc
1040      .CommandText = "sp_mostrar_depto"
1050      With rstSQL
1060         If .State = 1 Then
1070            .Close
1080         End If
1090         Set rstSQL = cmdSQL.Execute
1100         Set Me.dtgDepto.DataSource = rstSQL
1110      End With
1120   End With
1130   Set cmdSQL = Nothing
1140   Set cmdSQL.ActiveConnection = Nothing
       '1160   cmdSQL.ActiveConnection.Close 'com
1150   Set rstSQL = Nothing
ExitProc:
1160   Exit Sub
ControlError:
1170   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1180   Resume ExitProc
End Sub

Private Sub txtDepto_DblClick()
       
       Dim blnMostrarDat As Boolean
       Dim strCodDepto As String
       Dim strDescDepto As String
       
1000   On Error GoTo ControlError
       
1010   blnMostrarDat = frm_bDepto.BusqDepto(strCodDepto, strDescDepto)
1020   txtDepto.Text = strCodDepto
       '1030   txtDesPais = strDescPais 'com
1030   If Len(txtDepto.Text) > 0 Then
1040      cmdValidar_Click
1050   End If
1060   Me.Refresh
1070   Exit Sub
       
ExitProc:
1080   Exit Sub
ControlError:
1090   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1100   Resume ExitProc
End Sub
