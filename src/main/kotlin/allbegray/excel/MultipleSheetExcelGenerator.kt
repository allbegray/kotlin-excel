package allbegray.excel

import allbegray.excel.dataformat.DataFormatStrategy
import allbegray.excel.dataformat.DefaultDataFormatStrategy
import org.apache.poi.ss.usermodel.Sheet
import java.io.OutputStream

class MultipleSheetExcelGenerator<T> @JvmOverloads constructor(
    clazz: Class<T>,
    chunk: Int? = null,
    dataFormatStrategy: DataFormatStrategy = DefaultDataFormatStrategy(),
    excelType: ExcelType = ExcelType.SXSSF
) : AbstractExcelGenerator<T>(clazz, excelType, dataFormatStrategy) {

    protected val sheets: MutableList<Sheet> = mutableListOf()
    protected lateinit var currentSheet: Sheet
    protected var currentRowIndex: Int = 0
    protected var rownum = 0
    protected val chunk: Int = if (chunk == null || chunk > maxRows) maxRows else chunk

    init {
        createSheetWithHeader()
    }

    protected fun createSheetWithHeader() {
        currentSheet = createSheet("_${(sheets.size + 1)}").also {
            sheets.add(it)
        }
        currentRowIndex = 0
        renderHeader()
    }

    protected fun renderHeader() {
        renderHeader(currentSheet, currentRowIndex++)
    }

    protected fun renderBody(item: T) {
        renderBody(item, currentSheet, currentRowIndex++)
    }

    override fun addRow(item: T) {
        if (rownum > 0 && rownum % chunk == 0) {
            autoSizeColumns(currentSheet)
            createSheetWithHeader()
        }
        renderBody(item)
        rownum++
    }

    override fun addRows(items: List<T>) {
        items.forEach(::addRow)
    }

    override fun write(os: OutputStream) {
        autoSizeColumns(currentSheet)
        super.write(os)
    }
}
