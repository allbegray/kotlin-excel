package allbegray.excel.annotation

import org.apache.poi.ss.usermodel.Font

enum class Underline {
    NONE, SINGLE, DOUBLE, SINGLE_ACCOUNTING, DOUBLE_ACCOUNTING;

    fun toByte(): Byte {
        return when (this) {
            NONE -> Font.U_NONE
            SINGLE -> Font.U_SINGLE
            DOUBLE -> Font.U_DOUBLE
            SINGLE_ACCOUNTING -> Font.U_SINGLE_ACCOUNTING
            DOUBLE_ACCOUNTING -> Font.U_DOUBLE_ACCOUNTING
        }
    }
}
