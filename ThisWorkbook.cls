VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' =============================================================================
' CODIGO PARA AGREGAR AL MODULO ThisWorkbook
' FECHA: 2025-06-03 13:54:50 UTC
' DESCRIPCION: Evento Workbook_Open que ejecuta todos los procesos de inicio
' =============================================================================

' Agregar este codigo al evento Workbook_Open existente en ThisWorkbook
' Si ya existe un evento Workbook_Open, integrar las lineas marcadas con -->

Private Sub Workbook_Open()
    
    ' --> CODIGO NUEVO A AGREGAR - INICIO
    ' Ejecutar todos los procesos de inicio del libro
    Call Ejecutar_Procesos_Inicio_Libro
    ' --> CODIGO NUEVO A AGREGAR - FIN
    
    ' ... aqui va el resto del codigo existente del Workbook_Open ...
    
End Sub
