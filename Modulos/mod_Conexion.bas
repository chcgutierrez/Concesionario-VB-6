Attribute VB_Name = "mod_Conexion"
Sub main()
On Error GoTo ControlError
'Conectar a la BD
With BD_SQL
.CursorLocation = adUseClient 'Soy cliente de la BD
.Open "Provider=SQLOLEDB.1;" & _
      "Integrated Security=SSPI;" & _
      "Persist Security Info=False;" & _
      "Initial Catalog=farmacia;" & _
      "Data Source=CLIENTE-PC"
'frm_login.Show
End With
ExitProc:
Exit Sub
ControlError:
MsgBox "Ha ocurrido un error en la aplicación." & vbLf & vbLf & "Error: " & CStr(Err.Number) & _
          ". Descripción del error: " & Err.Description, vbCritical, App.Title
Resume ExitProc
End Sub

'Sub Abrir_tblUSUARIO() 'Conectar a tabla en la BD
'With RS_tblUSUARIO
'If .State = 1 Then .Close
'.Open "select * from usuario", BD_SQL, adOpenStatic, adLockOptimistic
'End With
'End Sub
'
'Sub buscar()
''busca el usuario en la BD
'With RS_tblUSUARIO
'.Requery
'.Find "id_usuario='" & Trim(frm_login.txt_usuario.Text) & "'" 'busca en el RS
''recorre el RS hasta el final, y si no encuentra
' If .EOF Then
'MsgBox "Usuario no encontrado", vbInformation, "Usuario" 'si no trae nada muestra el mensaje
'frm_login.txt_usuario.Text = ""
'frm_login.txt_usuario.SetFocus
' Exit Sub
'  Else
''si el usuario existe, valida a clave
'  If !clave_usuario = Trim(frm_login.txt_clave.Text) Then
''si todo esta ok
'    MsgBox "Ha iniciado una nueva sesión.", vbOKOnly, "Login"
'    frm_main.Show 'muestro el form principal
'    Unload frm_login 'cierro el form login
'  Else
''si no encuentra la clave
'    MsgBox "Contraseña Incorrecta", vbInformation, "Contraseña" 'si no trae nada muestra el mensaje
'    frm_login.txt_clave.Text = ""
'    frm_login.txt_clave.SetFocus
'  Exit Sub
'End If
'End If
'End With
'End Sub
'
'Sub login() 'valida los datos ingresados
'If frm_login.txt_usuario.Text = "" Then
'MsgBox "No se ha ingresado el usuario", vbInformation, "Ingresar Usuario"
'frm_login.txt_usuario.SetFocus
'If frm_login.txt_clave.Text = "" Then
'MsgBox "No se ha ingresado la contraseña", vbInformation, "Ingresar Contraseña"
'frm_login.txt_clave.SetFocus
'Exit Sub
'End If
'End If
'End Sub
'
'Sub cancelar()
'frm_login.txt_usuario.Text = ""
'frm_login.txt_clave.Text = ""
'End Sub
