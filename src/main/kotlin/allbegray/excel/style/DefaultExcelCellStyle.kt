package allbegray.excel.style

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Workbook

class DefaultExcelCellStyle : ExcelCellStyle {

    override fun apply(workbook: Workbook): CellStyle {
        return workbook.createCellStyle()
    }
}
