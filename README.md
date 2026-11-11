# bmar_summant_test

If COUNT() = 0 Then
    SetLog(0, 0, 0, "No hay registros filtrados para exportar.")
    Return M4_ERROR
EndIf


EXCEL!EXCEL.StartExcel()
EXCEL!EXCEL_WORKBOOKS.AddWorkbook()
EXCEL!EXCEL_WORKSHEETS.GetWorksheet(1)
EXCEL!EXCEL.Visible(0) /* Mantener oculto durante la carga para mayor velocidad */

vFila = 0

For i = 0 To COUNT() - 1
    MoveTo(i)
    vFila = i + 1
    
    vColumna = 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna,  BMAR_ID_HR)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna,  BMAR_GB_NAME)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna,  ID_OBRA_SOCIAL)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna,  REL_OS_PRE)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna,  CAPITAS)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna,  APORTES_APROX)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna,  ACT_ID_PREPAGA)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna,  ACT_ID_PLAN)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna,  ACT_COSTO_PLAN)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna, ACT_DIF_APORTES)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna, SIM_ID_PREPAGA)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna, SIM_ID_PLAN)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna, SIM_COSTO_PLAN)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna, SIM_DIF_APORTES)
    
    vColumna = vColumna + 1
    EXCEL!EXCEL_WORKSHEETS.SetCellValue(vFila, vColumna, COSTO_BANCO)
Next

EXCEL!EXCEL.Visible(1)
EXCEL!EXCEL.Yield()

Return M4_SUCCESS
