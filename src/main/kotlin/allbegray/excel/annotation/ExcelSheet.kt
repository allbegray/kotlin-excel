package allbegray.excel.annotation

import allbegray.excel.style.DefaultExcelCellStyle
import allbegray.excel.style.ExcelCellStyle
import java.lang.annotation.Inherited
import kotlin.reflect.KClass

@Target(AnnotationTarget.CLASS)
@Retention(AnnotationRetention.RUNTIME)
@MustBeDocumented
@Inherited
annotation class ExcelSheet(
    val value: String = "Sheet",
    val columnWidth: Int = 8,
    val rowHeight: Short = 300,
    val headerStyle: ExcelStyle = ExcelStyle(),
    val headerStyleClass: KClass<out ExcelCellStyle> = DefaultExcelCellStyle::class,
    val bodyStyle: ExcelStyle = ExcelStyle(),
    val bodyStyleClass: KClass<out ExcelCellStyle> = DefaultExcelCellStyle::class,
    /**
     * 필드 수동 정렬
     */
    val fieldOrder: Array<String> = [],
    /**
     * 필드 정렬 방식
     * - NONE: 정렬하지 않음
     * - NAME: 필드 이름으로 정렬
     * - ORDER: 필드 순서로 정렬
     */
    val fieldSort: Sort = Sort.NONE,
    /**
     * 틀 고정 (헤더)
     */
    val freezeHeaderPane: Boolean = false
)
