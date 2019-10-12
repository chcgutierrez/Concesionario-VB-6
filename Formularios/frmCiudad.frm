VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCiudad 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ciudad"
   ClientHeight    =   6300
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6225
   Icon            =   "frmCiudad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDesDepto 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   19
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox txtDesPais 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   2400
      TabIndex        =   18
      Top             =   645
      Width           =   3495
   End
   Begin VB.TextBox txtCodDepto 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1515
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton optActivo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Activo"
      Height          =   315
      Left            =   1725
      TabIndex        =   5
      Top             =   2655
      Width           =   780
   End
   Begin VB.OptionButton optInactivo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Inactivo"
      Height          =   315
      Left            =   2640
      TabIndex        =   6
      Top             =   2655
      Width           =   900
   End
   Begin VB.TextBox txtCodPais 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1515
      TabIndex        =   0
      Top             =   645
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
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtCiudad 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1515
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtNomCiudad 
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Top             =   2025
      Width           =   3735
   End
   Begin VB.TextBox txtObser 
      Height          =   735
      Left            =   1530
      TabIndex        =   7
      Top             =   3180
      Width           =   4350
   End
   Begin VB.Timer Timer1 
      Left            =   3480
      Top             =   1560
   End
   Begin MSDataGridLib.DataGrid dtgCiudad 
      Height          =   1455
      Left            =   320
      TabIndex        =   8
      Top             =   4200
      Width           =   5600
      _ExtentX        =   9895
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
      EndProperty
   End
   Begin MSComctlLib.ImageList img_lista 
      Left            =   4080
      Top             =   1440
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
            Picture         =   "frmCiudad.frx":058A
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCiudad.frx":0B24
            Key             =   "editar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCiudad.frx":10BE
            Key             =   "borrar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCiudad.frx":1658
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCiudad.frx":1BF2
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCiudad.frx":218C
            Key             =   "guardar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCiudad.frx":2726
            Key             =   "imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCiudad.frx":2CC0
            Key             =   "salir"
         EndProperty
      EndProperty
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
      Left            =   1530
      TabIndex        =   9
      Top             =   2430
      Width           =   2055
   End
   Begin MSComctlLib.Toolbar tlb_botones 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
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
      TabIndex        =   17
      Top             =   5925
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Ver 1.0.0"
            TextSave        =   "11/10/2019"
            Key             =   "sbrPan01"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "MAYÚS"
            Key             =   "sbrPan02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3228
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Departamento"
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pais"
      Height          =   195
      Left            =   1100
      TabIndex        =   14
      Top             =   645
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. Ciudad"
      Height          =   195
      Left            =   510
      TabIndex        =   13
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad"
      Height          =   195
      Left            =   900
      TabIndex        =   12
      Top             =   2025
      Width           =   495
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
      DataSource      =   "360"
      Height          =   195
      Left            =   370
      TabIndex        =   11
      Top             =   3180
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   5340
      Left            =   120
      Top             =   480
      Width           =   5950
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Existentes"
      DataSource      =   "360"
      Height          =   195
      Left            =   290
      TabIndex        =   10
      Top             =   3960
      Width           =   1185
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
Attribute VB_Name = "frmCiudad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytFlagModifica As Byte

Private Sub cmdValidar_Click()
    If Me.txtCodPais.Text <> "" And Me.txtCodDepto.Text <> "" And Me.txtCiudad.Text <> "" Then
        mnuArchivo_Buscar_Click
    Else
        MsgBox "Datos Incompletos", vbInformation + vbOKOnly, "Consultar"
    End If
End Sub

Private Sub Form_Load()
'    AbrirDepto
    mnuArchivo_Nuevo_Click
    PrenderMenus Me, tlb_botones, gcnstConsCompleta
End Sub


Private Sub mnuArchivo_Buscar_Click()
On Error GoTo ControlError

With cmdSQL
.ActiveConnection = ConexSQL
.CommandType = adCmdStoredProc
.CommandText = "sp_buscar_ciudad"
.Parameters.Append .CreateParameter("@codPais", adVarChar, adParamInput, 10, Me.txtCodPais.Text)
.Parameters.Append .CreateParameter("@codDepto", adVarChar, adParamInput, 10, Me.txtCodDepto.Text)
.Parameters.Append .CreateParameter("@codCiudad", adVarChar, adParamInput, 10, Me.txtCiudad.Text)
Set rstSQL = cmdSQL.Execute
End With
Set cmdSQL = Nothing
Set cmdSQL.ActiveConnection = Nothing
'cmdSQL.ActiveConnection.Close
If rstSQL.RecordCount > 0 Then
    If MsgBox("El registro ya existe. ¿Mostrar Datos?", vbQuestion + vbYesNo, "Consultar") = vbYes Then
        Me.txtNomCiudad.Text = rstSQL("nom_ciu").Value
            If rstSQL("est_ciu").Value = "A" Then
                Me.optActivo = True
            Else
                Me.optInactivo = True
            End If
        Me.txtObser.Text = rstSQL("obs_gen").Value
    End If
    Set rstSQL = Nothing
Else
Me.txtCodPais.Enabled = False
Me.txtCodDepto.Enabled = False
Me.txtCiudad.Enabled = False
Me.cmdValidar.Enabled = False
Me.txtNomCiudad.Enabled = True
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
Me.txtCodDepto.Enabled = False
Me.cmdValidar.Enabled = False
Me.txtDesPais.Enabled = False
Me.txtDesDepto.Enabled = False
Me.txtNomCiudad.Enabled = False
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
Me.txtCodDepto.Enabled = False
Me.txtCiudad.Enabled = False
Me.cmdValidar.Enabled = False
Me.txtNomCiudad.Enabled = True
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
Dim estCiudad As String
Dim rsCiudad As ADODB.Recordset

On Error GoTo ControlError

If Me.optActivo = True Then
estCiudad = "A"
ElseIf Me.optInactivo = True Then
estCiudad = "I"
End If
    
    If Len(Me.txtCodPais.Text) > 0 And Len(Me.txtCodDepto.Text) > 0 Then
        Set rsCiudad = TraerDepto(txtCodPais.Text, txtCodDepto.Text)
    End If

If bytFlagModifica = 0 Then

With cmdSQL
.ActiveConnection = ConexSQL
.CommandType = adCmdStoredProc
.CommandText = "sp_guardar_ciudad"
    .Parameters.Append .CreateParameter("@idDepto", adInteger, adParamInput, 10, rsCiudad("id_depto").Value)
    .Parameters.Append .CreateParameter("@idPais", adInteger, adParamInput, 10, rsCiudad("id_pais").Value)
    .Parameters.Append .CreateParameter("@codCiudad", adVarChar, adParamInput, 10, Me.txtCiudad.Text)
    .Parameters.Append .CreateParameter("@nomCiudad", adVarChar, adParamInput, 100, Me.txtNomCiudad.Text)
    .Parameters.Append .CreateParameter("@estCiudad", adVarChar, adParamInput, 10, estCiudad)
    .Parameters.Append .CreateParameter("@obsGen", adVarChar, adParamInput, 100, Me.txtObser.Text)
.Execute
End With
Set cmdSQL = Nothing
Set cmdSQL.ActiveConnection = Nothing
'cmdSQL.ActiveConnection.Close
mnuArchivo_Cancelar_Click
Me.StatusBar1.Panels(3) = "Datos Guardados Correctamente"
'MsgBox "Datos Guardados Correctamente", vbInformation + vbOKOnly, "Guardar"
'AbrirDepto

Else

With cmdSQL
.ActiveConnection = ConexSQL
.CommandType = adCmdStoredProc
.CommandText = "sp_editar_ciudad"
    .Parameters.Append .CreateParameter("@idDepto", adInteger, adParamInput, 10, rsCiudad("id_depto").Value)
    .Parameters.Append .CreateParameter("@idPais", adInteger, adParamInput, 10, rsCiudad("id_pais").Value)
    .Parameters.Append .CreateParameter("@codCiudad", adVarChar, adParamInput, 10, Me.txtCiudad.Text)
    .Parameters.Append .CreateParameter("@nomCiudad", adVarChar, adParamInput, 100, Me.txtNomCiudad.Text)
    .Parameters.Append .CreateParameter("@estCiudad", adVarChar, adParamInput, 10, estCiudad)
    .Parameters.Append .CreateParameter("@obsGen", adVarChar, adParamInput, 100, Me.txtObser.Text)
.Execute
End With
Set cmdSQL = Nothing
Set cmdSQL.ActiveConnection = Nothing
'cmdSQL.ActiveConnection.Close
mnuArchivo_Cancelar_Click
Me.StatusBar1.Panels(3) = "Datos Guardados Correctamente"
'MsgBox "Datos Guardados Correctamente", vbInformation + vbOKOnly, "Guardar"
'AbrirDepto

End If

Set rsCiudad = Nothing

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
Me.txtCodDepto.Text = ""
Me.txtCodDepto.Enabled = True
Me.txtCiudad.Text = ""
Me.txtCiudad.Enabled = True
Me.cmdValidar.Enabled = True
Me.txtDesPais.Text = ""
Me.txtDesPais.Enabled = False
Me.txtDesDepto.Text = ""
Me.txtDesDepto.Enabled = False
Me.txtNomCiudad.Text = ""
Me.txtNomCiudad.Enabled = False
Me.optActivo.Enabled = False
Me.optActivo.Value = False
Me.optInactivo.Enabled = False
Me.optInactivo.Value = False
Me.txtObser.Text = ""
Me.txtObser.Enabled = False
Me.StatusBar1.Panels(3) = ""
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
'
        Case "btnEditar": mnuArchivo_Editar_Click
'
        Case "btnGuardar": mnuArchivo_Guardar_Click
        
        Case "btnSalir": mnuArchivo_Salir_Click

    End Select
End Sub


Private Sub txtCodDepto_Validate(Cancel As Boolean)
Dim rsDepto As ADODB.Recordset
    
    On Error GoTo ControlError
    
    If Len(Me.txtCodPais.Text) > 0 And Len(Me.txtCodDepto.Text) > 0 Then
        Set rsDepto = TraerDepto(txtCodPais.Text, txtCodDepto.Text)
            If rsDepto.RecordCount > 0 Then
                    If rsDepto("est_depto").Value = "I" Then
                        Me.txtCodDepto.SelStart = 0
                        Me.txtCodDepto.SelLength = Len(Me.txtCodDepto.Text)
                        MsgBox "El Departamento ingresado está inactivo.", vbOKOnly, "Buscar Departamento"
                        Me.txtDesDepto.Text = ""
                        Cancel = True
                        Exit Sub
                    End If
                Me.txtDesDepto.Text = rsDepto("nom_depto").Value
            Else
                Me.txtCodDepto.SelStart = 0
                Me.txtCodDepto.SelLength = Len(Me.txtCodPais.Text)
                MsgBox "No existe el Departamento para el criterio ingresado.", vbOKOnly, "Buscar Departamento"
                Me.txtDesDepto.Text = ""
                Cancel = True
                Exit Sub
            End If
    Else
         Me.txtDesDepto.Text = ""
         MsgBox "Debe ingresar un criterio para realizar la busqueda.", vbOKOnly, "Criterio Inválido"
         Me.txtCodDepto.SetFocus
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
