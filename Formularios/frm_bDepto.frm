VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_bDepto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Departamento"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Realizar Búsqueda"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6880
      Begin VB.TextBox txtDescDepto 
         Height          =   375
         Left            =   1250
         TabIndex        =   0
         Top             =   300
         Width           =   4335
      End
      Begin VB.CommandButton btnBusqDepto 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   280
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.CommandButton btnBsqAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dtgDeptoAct 
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "cod_depto"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nom_depto"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "est_depto"
         Caption         =   "Estado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "obs_gen"
         Caption         =   "Observaciones"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   22538
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
      EndProperty
   End
End
Attribute VB_Name = "frm_bDepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bsqDepto As Boolean


Public Function BusqDepto(ByRef strCodDepto As String, ByRef strDescDepto As String) As Boolean
       
1000   On Error GoTo ControlError
       
1010   Set dtgDeptoAct.DataSource = Nothing
1020   bsqDepto = False
1030   Me.Show vbModal
1040   If bsqDepto Then
1050      If Not dtgDeptoAct.DataSource Is Nothing Then 'Grid del form busqueda
1070         strCodDepto = dtgDeptoAct.Columns(0).Text
1080         strDescDepto = dtgDeptoAct.Columns(1).Text
1090         BusqDepto = True
1100      Else
1110         BusqDepto = False
1120      End If
1130   Else
1140      BusqDepto = False
1150   End If
       
1160   Exit Function
ControlError:
1170   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
       
End Function

Private Sub btnBsqAceptar_Click()
100   On Error GoTo ControlError
110   bsqDepto = True
120   Me.Hide
ExitProc:
130   Exit Sub
ControlError:
140   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
         ". Descripción del error: " & Err.Description, vbCritical, App.Title
150   Resume ExitProc
End Sub

Private Sub btnBusqDepto_Click()
       
       Dim rsBusqDepto As ADODB.Recordset
       Dim CodPais As String
       
1000   On Error GoTo ControlError
1010   CodPais = TraerPaisDesc(frmDepto.txtDesPais.Text)
1020   If Len(Me.txtDescDepto.Text) > 0 Then
1030      Set rsBusqDepto = TraerDeptoDesc(CodPais, txtDescDepto.Text)
1040      If rsBusqDepto.RecordCount > 0 Then
1050         Set dtgDeptoAct.DataSource = rsBusqDepto
1060         dtgDeptoAct.Columns("Codigo").Width = 900
1070         dtgDeptoAct.Columns("Nombre").Width = 2300
1080         dtgDeptoAct.Columns("Estado").Width = 800
1090         dtgDeptoAct.Columns("Observaciones").Width = 1300
1100      Else
1110         Set dtgDeptoAct.DataSource = Nothing
1120         Me.txtDescDepto.SelStart = 0
1130         Me.txtDescDepto.SelLength = Len(Me.txtDescDepto.Text)
1140         MsgBox "El departamento no existe o está inactivo", vbOKOnly, "Buscar Departamento"
             '                Cancel
1150         Exit Sub
1160      End If
1170   Else
1180      MsgBox "Debe ingresar un criterio para realizar la busqueda.", vbOKOnly, "Criterio Inválido"
1190      Me.txtDescDepto.SetFocus
1200      Set dtgDeptoAct.DataSource = Nothing
          '         Cancel = True
1210      Exit Sub
1220   End If
       
ExitProc:
1230   Exit Sub
ControlError:
1240   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
1250   Resume ExitProc
End Sub

Private Sub dtgDeptoAct_DblClick()
100   On Error GoTo ControlError
110   bsqDepto = True
120   Me.Hide
ExitProc:
130   Exit Sub
ControlError:
140   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
         ". Descripción del error: " & Err.Description, vbCritical, App.Title
150   Resume ExitProc
End Sub

