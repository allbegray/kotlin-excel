package allbegray.excel.dataformat

import org.apache.poi.ss.usermodel.DataFormat
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.*

class DefaultDataFormatStrategy : DataFormatStrategy {

    override fun apply(dataFormat: DataFormat, type: Class<*>): Short {
        val kClass = type.kotlin

        val integerTypes = listOf(Byte::class, Short::class, Int::class, Long::class)
        val realTypes = listOf(Float::class, Double::class)
        val dateTimeTypes = listOf(Date::class, Calendar::class, LocalDateTime::class)

        val format = when {
            integerTypes.contains(kClass) -> "#,##0"
            realTypes.contains(kClass) -> "#,##0.00"
            kClass == LocalDate::class -> "yyyy-mm-dd"
            dateTimeTypes.contains(kClass) -> "yyyy-mm-dd hh:mm:ss"
            else -> null
        }

        return format?.let { dataFormat.getFormat(it) } ?: 0
    }
}
