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
       
    On Error GoTo ControlError
    
        Set dtgDeptoAct.DataSource = Nothing
        bsqDepto = False
        Me.Show vbModal
        If bsqDepto Then
            If Not dtgDeptoAct.DataSource Is Nothing Then 'Grid del form busqueda
                strCodDepto = dtgDeptoAct.Columns(0).Text
                strDescDepto = dtgDeptoAct.Columns(1).Text
                BusqDepto = True
            Else
                BusqDepto = False
            End If
        Else
            BusqDepto = False
        End If
        
    Exit Function
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
          
End Function

Private Sub btnBsqAceptar_Click()
On Error GoTo ControlError
    bsqDepto = True
    Me.Hide
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

Private Sub btnBusqDepto_Click()

Dim rsBusqDepto As ADODB.Recordset
    
    On Error GoTo ControlError
'    CodPais = TraerPaisDesc(txtDescPais.Text)
    If Len(Me.txtDescDepto.Text) > 0 Then
        Set rsBusqDepto = TraerDeptoDesc(CodPais, txtDescDepto.Text)
            If rsBusqDepto.RecordCount > 0 Then
                Set dtgDeptoAct.DataSource = rsBusqDepto
                dtgDeptoAct.Columns("Codigo").Width = 900
                dtgDeptoAct.Columns("Nombre").Width = 2300
                dtgDeptoAct.Columns("Estado").Width = 800
                dtgDeptoAct.Columns("Observaciones").Width = 1300
            Else
                Set dtgDeptoAct.DataSource = Nothing
                Me.txtDescDepto.SelStart = 0
                Me.txtDescDepto.SelLength = Len(Me.txtDescDepto.Text)
                MsgBox "El departamento no existe o está inactivo", vbOKOnly, "Buscar Departamento"
'                Cancel
                Exit Sub
            End If
    Else
         MsgBox "Debe ingresar un criterio para realizar la busqueda.", vbOKOnly, "Criterio Inválido"
         Me.txtDescDepto.SetFocus
         Set dtgDeptoAct.DataSource = Nothing
'         Cancel = True
         Exit Sub
    End If
    
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

Private Sub dtgDeptoAct_DblClick()
On Error GoTo ControlError
    bsqDepto = True
    Me.Hide
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

