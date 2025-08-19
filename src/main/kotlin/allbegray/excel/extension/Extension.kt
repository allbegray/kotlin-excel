package allbegray.excel.extension

import allbegray.excel.style.color.RGBColor
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.XSSFCellStyle

fun CellStyle.setBorder(borderStyle: BorderStyle) {
    borderTop = borderStyle
    borderLeft = borderStyle
    borderRight = borderStyle
    borderBottom = borderStyle
}

fun CellStyle.setBorderColor(color: IndexedColors) {
    setBorderColor(color.index)
}

fun CellStyle.setBorderColor(color: Short) {
    topBorderColor = color
    leftBorderColor = color
    rightBorderColor = color
    bottomBorderColor = color
}

fun CellStyle.setFillForegroundColor(color: IndexedColors) {
    fillForegroundColor = color.index
}

fun CellStyle.setFillBackgroundColor(color: IndexedColors) {
    fillBackgroundColor = color.index
}

fun XSSFCellStyle.setFillForegroundColor(color: RGBColor) {
    setFillBackgroundColor(color.toXSSFColor())
}
