package allbegray.excel.style

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Workbook

interface ExcelCellStyle {

    fun apply(workbook: Workbook): CellStyle
}
