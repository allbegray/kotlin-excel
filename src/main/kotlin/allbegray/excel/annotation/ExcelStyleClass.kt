package allbegray.excel.annotation

import allbegray.excel.style.ExcelCellStyle
import kotlin.reflect.KClass

@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
@MustBeDocumented
annotation class ExcelStyleClass(val value: KClass<out ExcelCellStyle>)
