package allbegray.excel.dataformat

import org.apache.poi.ss.usermodel.DataFormat

interface DataFormatStrategy {

    fun apply(dataFormat: DataFormat, type: Class<*>): Short

    companion object {
        @JvmStatic
        fun default(): DataFormatStrategy = DefaultDataFormatStrategy()

        @JvmStatic
        fun empty(): DataFormatStrategy = EmptyDataFormatStrategy()
    }
}
