package allbegray.excel.style

import allbegray.excel.extension.setBorder
import org.apache.poi.ss.usermodel.*

class DefaultHeaderExcelCellStyle : ExcelCellStyle {

    override fun apply(workbook: Workbook): CellStyle {
        return workbook.createCellStyle().apply {
            alignment = HorizontalAlignment.CENTER
            verticalAlignment = VerticalAlignment.CENTER
            setBorder(BorderStyle.THIN)
            fillPattern = FillPatternType.SOLID_FOREGROUND
            fillForegroundColor = IndexedColors.GREY_25_PERCENT.index

            val font = workbook.createFont().apply {
                bold = true
            }
            this.setFont(font)
        }
    }
}
