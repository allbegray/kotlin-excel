# kotlin-excel

kotlin apache poi 를 사용하여 DTO 기반으로 excel 다운로드 할 수 있는 라이브러리

### Example
```kotlin
class SingleSheetExcelGeneratorTest {

    class TestHeaderExcelCellStyle : ExcelCellStyle {

        override fun apply(workbook: Workbook): CellStyle {
            val baseColor = RGBColor(10, 180, 173)

            return workbook.createCellStyle().apply {
                alignment = HorizontalAlignment.CENTER
                verticalAlignment = VerticalAlignment.CENTER
                setBorder(BorderStyle.THIN)
                setBorderColor(IndexedColors.WHITE.index)
                fillPattern = FillPatternType.SOLID_FOREGROUND

                (this as XSSFCellStyle).setFillForegroundColor(baseColor.toXSSFColor())

                val font = workbook.createFont().apply {
                    bold = true
                    color = IndexedColors.WHITE.index
                }
                this.setFont(font)
            }
        }
    }

    class TestBodyExcelCellStyle : ExcelCellStyle {

        override fun apply(workbook: Workbook): CellStyle {
            val baseColor = RGBColor(10, 180, 173)

            return workbook.createCellStyle().apply {
                verticalAlignment = VerticalAlignment.CENTER
                setBorder(BorderStyle.THIN)

                (this as XSSFCellStyle).apply {
                    val color = baseColor.toXSSFColor()
                    setTopBorderColor(color)
                    setLeftBorderColor(color)
                    setRightBorderColor(color)
                    setBottomBorderColor(color)
                }
            }
        }
    }

    @ExcelSheet(
        value = "테스트 시트",
        columnWidth = 8,
        rowHeight = 300,
        // 헤더 스타일 커스텀. 기본값은 poi 에서 제공하는 기본 스타일
        headerStyleClass = TestHeaderExcelCellStyle::class,
        // 본문 스타일 커스텀. 기본값은 poi 에서 제공하는 기본 스타일
        bodyStyleClass = TestBodyExcelCellStyle::class,
        // 필드 순서를 수동으로 지정하고 싶을 때 사용. 기본값은 jvm 필드 순서에 따른다.
        fieldOrder = [
            "컬럼1",
            "컬럼2",
            "컬럼3",
            "컬럼4",
            "컬럼5",
            "컬럼6"
        ],
        // 헤더 고정
        freezeHeaderPane = true
    )
    data class Pojo(
        @ExcelColumn("컬럼1")
        val foo: String,
        @ExcelColumn("컬럼2")
        val bar: Int,
        val bool: Boolean = Random.nextBoolean()
    ) {
        @ExcelColumn("컬럼4")
        val ldt: LocalDateTime = LocalDateTime.now()

        @ExcelColumn("컬럼5")
        val year: Year = Year.now()

        @ExcelColumn("컬럼3")
        fun zoo(): String = foo.repeat(2)

        // boolean 전용 데이터 포멧팅은 없기 때문에 수동으로 포멧팅
        @ExcelColumn("컬럼6")
        fun excelBool(): String = if (bool) "참" else "거짓"
    }

    @Test
    fun test() {
        val rows = listOf(
            Pojo("푸", 1),
            Pojo("바", 2)
        )
        val generator = SingleSheetExcelGenerator(Pojo::class.java)
        for (row in rows) {
            generator.addRow(row)
        }
        generator.write(FileOutputStream("build/test.xlsx"))
    }
}

```

### Coming soon next
컬럼에 @ExcelStyle 어노테이션 사용 시 @ExcelSheet 에서 지정한 본문 스타일 상속 처리