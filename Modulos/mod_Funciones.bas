Attribute VB_Name = "mod_Funciones"
Option Explicit

'***************************************************************************************
'Constantes de control para botones, menus y combinar acciones
'***************************************************************************************
Public Const cnstMnuNuevo As Integer = 256
Public Const cnstMnuEditar As Integer = 128
Public Const cnstMnuBorrar As Integer = 64
Public Const cnstMnuCancelar As Integer = 32
Public Const cnstMnuBuscar As Integer = 16
Public Const cnstMnuCargar As Integer = 8
Public Const cnstMnuGuardar As Integer = 4
Public Const cnstMnuImprimir As Integer = 2
Public Const cnstMnuCerrar As Integer = 1
Public Const cnstMnuNada As Integer = 0

Public Const m_ColorNuevo As Long = &HFFF6F0
Public Const m_ColorExistente As Long = vbWhite

'***************************************************************************************
'Acciones para combinar con cada boton y menu y activar o no
'***************************************************************************************
Public Enum TipoAccionMenu
   gcnstReporte = cnstMnuImprimir + cnstMnuGuardar + cnstMnuNuevo + cnstMnuCerrar                                         'Habilita Nuevo, Guardar, Imprimir y Salir
   gcnstEntrar = cnstMnuNuevo + cnstMnuBuscar + cnstMnuCargar + cnstMnuCerrar                                          ' habilita Nuevo, buscar, consultar y Salir
   gcnstNuevo = cnstMnuCancelar + cnstMnuGuardar + cnstMnuCerrar                                                        ' habilita Cancelar, Guardar y Salir
   gcnstEditar = cnstMnuCancelar + cnstMnuBuscar + cnstMnuGuardar + cnstMnuCerrar                                   ' habilita Cancelar, Consultar, Guardar y Salir
   gcnstEliminar = cnstMnuNuevo + cnstMnuBuscar + cnstMnuBorrar + cnstMnuCerrar                                       ' habilita Nuevo, buscar, consultar y Salir
   gcnstCancelar = cnstMnuNuevo + cnstMnuBuscar + cnstMnuCargar + cnstMnuCerrar                                        ' habilita Nuevo, buscar, consultar y Salir
   gcnstGuardar = cnstMnuNuevo + cnstMnuEditar + cnstMnuBorrar + cnstMnuBuscar + cnstMnuImprimir + cnstMnuCerrar    ' habilita Nuevo, Modificar, Eliminar, Consultar, Imprimir y Salir
   gcnstCargar = cnstMnuCancelar + cnstMnuCerrar + cnstMnuBuscar + cnstMnuCargar                                                    ' habilita Cancelar, Salir y Buscar
   gcnstBuscarImprimir = cnstMnuCancelar + cnstMnuCerrar + cnstMnuBuscar + cnstMnuImprimir                                                      ' habilita Cancelar, Salir y Buscar
   gcnstConsCompleta = cnstMnuNuevo + cnstMnuEditar + cnstMnuGuardar + cnstMnuBorrar + cnstMnuCancelar + cnstMnuImprimir + cnstMnuCerrar ' habilita Nuevo, Modificar, Eliminar, Cancelar, Imprimir y Salir
   gcnstBuscar = cnstMnuEditar + cnstMnuBorrar + cnstMnuCancelar + cnstMnuImprimir + cnstMnuCerrar                    ' habilita Modificar, Eliminar, Cancelar, Imprimir y Salir
   gcnstCerrar = cnstMnuCerrar                                                                                          ' habilita salir
   gcnstNada = 0                                                                                                  ' No habilita ninguno
   gcnstPredet = cnstMnuCancelar + cnstMnuCerrar + cnstMnuBuscar                                                        ' habilita Cancelar y Salir
End Enum

'***************************************************************************************
'Nombre: PrenderMenus. Rutina para la habilitar los menus.
'Parámetros:
'* (Obligatorio) ByVal Formulario:
'* (Obligatorio) ByVal tlbmenu:
'* (Obligatorio) ByVal bytCodigo:
'* (Opcional) ByVal intBotonAct:
'* (Opcional) ByVal blnPrender:
'***************************************************************************************
Public Sub PrenderMenus(ByVal Formulario As Object, ByVal tlbFormulario As Toolbar, _
                        ByVal bytCodigo As TipoAccionMenu, Optional ByVal intBotonAct As Integer = 0, _
                        Optional ByVal blnPrender As Boolean = True)
       'On Error GoTo ControlError
1000   If Not tlbFormulario Is Nothing Then
1010      With tlbFormulario
1020         .Buttons("btnNuevo").Enabled = ((cnstMnuNuevo And bytCodigo) Or (cnstMnuNuevo And intBotonAct)) And blnPrender
1030         .Buttons("btnEditar").Enabled = ((cnstMnuEditar And bytCodigo) Or (cnstMnuEditar And intBotonAct)) And blnPrender
1040         .Buttons("btnBorrar").Enabled = ((cnstMnuBorrar And bytCodigo) Or (cnstMnuBorrar And intBotonAct)) And blnPrender
1050         .Buttons("btnCancelar").Enabled = (cnstMnuCancelar And bytCodigo) Or (cnstMnuCancelar And intBotonAct) And blnPrender
             '            .Buttons("btnCargar").Enabled = ((cnstMnuCargar And bytCodigo) Or (cnstMnuCargar And intBotonAct)) And blnPrender
1060         .Buttons("btnBuscar").Enabled = ((cnstMnuBuscar And bytCodigo) Or (cnstMnuBuscar And intBotonAct)) And blnPrender
1070         .Buttons("btnGuardar").Enabled = ((cnstMnuGuardar And bytCodigo) Or (cnstMnuGuardar And intBotonAct)) And blnPrender
1080         .Buttons("btnImprimir").Enabled = (cnstMnuImprimir And bytCodigo) Or (cnstMnuImprimir And intBotonAct) And blnPrender
1090         .Buttons("btnSalir").Enabled = (cnstMnuCerrar And bytCodigo) Or (cnstMnuCerrar And intBotonAct) And blnPrender
1100         Formulario.mnuArchivo_Nuevo.Enabled = .Buttons("btnNuevo").Enabled
1110         Formulario.mnuArchivo_Editar.Enabled = .Buttons("btnEditar").Enabled
1120         Formulario.mnuArchivo_Cancelar.Enabled = .Buttons("btnCancelar").Enabled
             '            Formulario.mnuArchivo_Borrar.Enabled = .Buttons("btnBorrar").Enabled
             '            Formulario.mnuArchivo_Cargar.Enabled = .Buttons("btnCargar").Enabled
1130         Formulario.mnuArchivo_Buscar.Enabled = .Buttons("btnBuscar").Enabled
1140         Formulario.mnuArchivo_Guardar.Enabled = .Buttons("btnGuardar").Enabled
1150         Formulario.mnuArchivo_Imprimir.Enabled = .Buttons("btnImprimir").Enabled
1160         Formulario.mnuArchivo_Salir.Enabled = .Buttons("btnSalir").Enabled
1170      End With
1180   End If
1190   Exit Sub
ControlError:
1200   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error:" & Err.Description, vbCritical, App.Title
End Sub

Public Sub TextoMayusc(ByVal Frm As Form)
1000   On Error GoTo ControlError
       Dim Ctrl As Control
1010   For Each Ctrl In Frm.Controls
1020      If TypeOf Ctrl Is TextBox Then
1030         Ctrl.Text = UCase(Ctrl.Text)
1040         Ctrl.SelStart = Len(Ctrl)
1050      End If
1060   Next
1070   Exit Sub
ControlError:
1080   MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
End Sub
