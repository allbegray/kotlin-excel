package allbegray.excel

import io.github.oshai.kotlinlogging.KotlinLogging
import org.apache.poi.ss.formula.eval.ErrorEval
import org.apache.poi.ss.usermodel.*
import java.io.InputStream

/**
 * 엑셀 파싱 유틸리티.
 *
 * 억지로 Row 데이터를 POJO 에 집어 넣어서 리턴 하려는 시도 하지 말 것.
 *
 * 특정한 변환 라이브러리를 강제 하지 마라.
 */
object ExcelParser {

    private val logger = KotlinLogging.logger {}

    private fun cellValue(cell: Cell?, formulaEvaluator: FormulaEvaluator): Any? {
        if (cell == null) return null
        val cellType = cell.cellType ?: throw NullPointerException("셀 유형 값이 존재 하지 않습니다.")
        return when (cellType) {
            CellType._NONE -> null
            CellType.NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    cell.dateCellValue
                } else {
                    cell.numericCellValue
                }
            }

            CellType.STRING -> cell.stringCellValue
            CellType.FORMULA -> {
                val evaluate: CellValue = formulaEvaluator.evaluate(cell)
                val evaluateCellType = evaluate.cellType ?: throw NullPointerException("셀 유형 값이 존재 하지 않습니다.")
                when (evaluateCellType) {
                    CellType._NONE -> null
                    CellType.NUMERIC -> evaluate.numberValue
                    CellType.STRING -> evaluate.stringValue
                    CellType.FORMULA -> throw UnsupportedOperationException("FORMULA 내 FORMULA 는 지원하지 않습니다.")
                    CellType.BLANK -> null
                    CellType.BOOLEAN -> evaluate.booleanValue
                    CellType.ERROR -> {
                        val errorMessage = ErrorEval.getText(evaluate.errorValue.toInt())
                        throw IllegalStateException("수식에 오류가 존재 합니다. $errorMessage")
                    }
                }
            }

            CellType.BLANK -> null
            CellType.BOOLEAN -> cell.booleanCellValue
            CellType.ERROR -> {
                val errorMessage = ErrorEval.getText(cell.errorCellValue.toInt())
                throw IllegalStateException("수식에 오류가 존재 합니다. $errorMessage")
            }
        }
    }

    private fun cellValue(row: Row, cellnum: Int, formulaEvaluator: FormulaEvaluator): Any? {
        val cell = row.getCell(cellnum)
        return allbegray.excel.ExcelParser.cellValue(cell, formulaEvaluator)
    }

    fun parseWithoutHeader(workbook: Workbook, sheetIndex: Int = 0, block: (row: List<Any?>) -> Unit) {
        val sheet = workbook.getSheetAt(sheetIndex)
        val creationHelper = workbook.creationHelper
        val formulaEvaluator = creationHelper.createFormulaEvaluator()

        for (rownum in 0..sheet.lastRowNum) {
            val row = sheet.getRow(rownum)
            val values = row.map { cell ->
                allbegray.excel.ExcelParser.cellValue(cell, formulaEvaluator)
            }
            block(values)
        }
    }

    /**
     * @param is 입력스트림. 자동으로 close 하지 않습니다. 반드시 close 핸들링 하세요.
     */
    fun parseWithoutHeader(`is`: InputStream, sheetIndex: Int = 0, block: (row: List<Any?>) -> Unit) {
        val workbook: Workbook = WorkbookFactory.create(`is`)
        parseWithoutHeader(workbook, sheetIndex, block)
    }

    fun parse(workbook: Workbook, headerRownum: Int = 0, sheetIndex: Int = 0, block: (row: Map<String, Any?>) -> Unit) {
        val sheet = workbook.getSheetAt(sheetIndex)
        val creationHelper = workbook.creationHelper
        val formulaEvaluator = creationHelper.createFormulaEvaluator()

        val headers: List<String> = sheet.getRow(headerRownum).map { it.stringCellValue }

        if (headers.size != headers.filter { it.isNotBlank() }.size) {
            throw IllegalArgumentException("공백의 헤더 이름이 존재 합니다.")
        }
        if (headers.size != headers.toSet().size) {
            throw IllegalArgumentException("중복된 헤더 이름이 존재 합니다.")
        }

        for (rownum in (headerRownum + 1)..sheet.lastRowNum) {
            val row = sheet.getRow(rownum)
            val map = headers.mapIndexed { cellnum, name ->
                val value = allbegray.excel.ExcelParser.cellValue(row, cellnum, formulaEvaluator)
                name to value
            }.toMap()

            block(map)
        }
    }

    /**
     * @param is 입력스트림. 자동으로 close 하지 않습니다. 반드시 close 핸들링 하세요.
     */
    fun parse(`is`: InputStream, headerRownum: Int = 0, sheetIndex: Int = 0, block: (row: Map<String, Any?>) -> Unit) {
        val workbook: Workbook = WorkbookFactory.create(`is`)
        parse(workbook, headerRownum, sheetIndex, block)
    }
}
