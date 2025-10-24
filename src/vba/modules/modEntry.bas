Attribute VB_Name = "modEntry"
Option Explicit

' Punto de entrada rápido para abrir el formulario de registro.
' Crea un acceso directo sencillo para el usuario final.

Public Sub AbrirRegistroESV()
    On Error GoTo EH
    ' Asegura que las hojas/tablas/rangos existan
    SetupESVWorkbook

    ' Intenta mostrar el formulario (debe existir en el proyecto como frmRegistroESV)
    VBA.UserForms.Add("frmRegistroESV").Show
    Exit Sub
EH:
    MsgBox "No se pudo abrir el formulario 'frmRegistroESV'.\nAsegúrate de crearlo en el Editor de VBA.", vbExclamation
End Sub
