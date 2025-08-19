package allbegray.excel.annotation

/**
 * function 에 사용 시 리턴 타입이 있는 getter 만 등록 된다.
 */
@Target(AnnotationTarget.FIELD, AnnotationTarget.FUNCTION, AnnotationTarget.PROPERTY_GETTER)
@Retention(AnnotationRetention.RUNTIME)
@MustBeDocumented
annotation class ExcelColumn(
    val value: String = "",
    val order: Int = 0,
    /**
     * 1 글자가 256 사이즈 이기 때문에 256 배수로 설정
     */
    val width: Int = 15 * 256,
    val autoSize: Boolean = false
)
