package allbegray.excel

import java.io.OutputStream

interface ExcelGenerator<T> {

    fun excelType(): ExcelType

    fun fileExtension(): String

    fun mediaType(): String

    fun addRow(item: T)

    fun addRows(items: List<T>)

    fun write(os: OutputStream)
}
