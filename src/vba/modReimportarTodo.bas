Attribute VB_Name = "modReimportarTodo"
Option Explicit

' Requiere referencia: Microsoft Visual Basic for Applications Extensibility 5.3
' Trust Center: habilitar "Trust access to the VBA project object model"

' Tipos de componente (evita depender de constantes de VBIDE)
Private Const CT_STD_MODULE As Long = 1      ' vbext_ct_StdModule
Private Const CT_CLASS_MODULE As Long = 2    ' vbext_ct_ClassModule
Private Const CT_MSFORM As Long = 3          ' vbext_ct_MSForm
Private Const CT_DOCUMENT As Long = 100      ' vbext_ct_Document

Public Sub ReimportarTodoElCodigo()
    Dim vbProj As Object
    Dim comp As Object
    Dim i As Long
    Dim base As String
    Dim modPath As String
    Dim clsPath As String
    Dim f As String
    
    Set vbProj = Application.VBE.ActiveVBProject
    base = "g:\My Drive\Trabajo\Staff-Fire\Seguridad Vehicular\Archivos de Trabajo\formulario_vba\src\vba"
    modPath = base & "\modules\"
    clsPath = base & "\classes\"
    
    ' Paso 1: Eliminar todos los módulos y clases (excepto este módulo y hojas/formularios)
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents(i)
        
        ' Solo eliminar módulos estándar y clases
        ' NO eliminar hojas (Document) ni formularios (MSForm)
        If comp.Type = CT_STD_MODULE Or comp.Type = CT_CLASS_MODULE Then
            ' No eliminar este módulo mientras se ejecuta
            If comp.name <> "modReimportarTodo" Then
                On Error Resume Next
                vbProj.VBComponents.Remove comp
                On Error GoTo 0
            End If
        End If
    Next i
    
    ' Paso 2: Importar módulos estándar (.bas)
    f = Dir(modPath & "*.bas")
    Do While Len(f) > 0
        vbProj.VBComponents.Import modPath & f
        f = Dir
    Loop
    
    ' Paso 3: Importar clases (.cls)
    f = Dir(clsPath & "*.cls")
    Do While Len(f) > 0
        vbProj.VBComponents.Import clsPath & f
        f = Dir
    Loop

    ' Paso 4: Limpiar encabezados inválidos en clases (si los hubiera)
    For i = 1 To vbProj.VBComponents.Count
        Set comp = vbProj.VBComponents(i)
        If comp.Type = CT_CLASS_MODULE Then
            CleanClassHeaders comp.CodeModule
        End If
    Next i
    
    MsgBox "Reimportación completada." & vbCrLf & _
           "Módulos y clases eliminados y reimportados desde:" & vbCrLf & _
           "- " & modPath & vbCrLf & _
           "- " & clsPath, vbInformation, "Reimportación"
End Sub

Private Sub CleanClassHeaders(cm As Object)
    ' Elimina, si existen, líneas de encabezado no válidas dentro del editor
    ' como: VERSION 1.0 CLASS / BEGIN / END al inicio del módulo
    Dim maxCheck As Long: maxCheck = Application.WorksheetFunction.Min(10, cm.CountOfLines)
    Dim i As Long
    Dim removed As Boolean
    ' Repetir hasta que ya no encuentre encabezados en las primeras líneas
    Do
        removed = False
        If cm.CountOfLines = 0 Then Exit Do
        ' Revisar primeras 3 líneas típicas del encabezado exportado
        For i = 1 To Application.WorksheetFunction.Min(3, cm.CountOfLines)
            Dim ln As String
            ln = Trim$(cm.Lines(i, 1))
            If ln Like "VERSION * CLASS" Or ln = "BEGIN" Or ln = "END" Then
                cm.DeleteLines i, 1
                removed = True
                Exit For
            ElseIf Left$(ln, 9) = "Attribute" Then
                ' Atributos son válidos, no tocarlos
            ElseIf Len(ln) = 0 Then
                cm.DeleteLines i, 1
                removed = True
                Exit For
            End If
        Next i
    Loop While removed
End Sub

Public Sub EliminarModReimportarTodo()
    ' Este procedimiento elimina el módulo modReimportarTodo después de usarlo
    ' Ejecutar SOLO después de haber ejecutado ReimportarTodoElCodigo
    Dim vbProj As Object
    Set vbProj = Application.VBE.ActiveVBProject
    
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents("modReimportarTodo")
    On Error GoTo 0
    
    MsgBox "Módulo modReimportarTodo eliminado.", vbInformation
End Sub


