package allbegray.excel.dataformat

import org.apache.poi.ss.usermodel.DataFormat

class EmptyDataFormatStrategy : DataFormatStrategy {

    override fun apply(dataFormat: DataFormat, type: Class<*>): Short {
        return 0
    }
}
