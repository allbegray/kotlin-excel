# kotlin-excel

kotlin apache poi 를 사용하여 DTO 기반으로 excel 다운로드 할 수 있는 라이브러리

### Example
```
    @ExcelSheet(
        value = "테스트 시트",
        // headerStyleClass = TestHeaderExcelCellStyle::class,
        // bodyStyleClass = TestBodyExcelCellStyle::class,
        
        /* *
        // fieldOrder = [
        //     "컬럼1",
        //     "컬럼2",
        //     "컬럼3"
        // ]
    )
    data class Pojo(
        @ExcelColumn("컬럼1")
        val foo: String,
        @ExcelColumn("컬럼2")
        val bar: Int,
    ) {
        @ExcelColumn("컬럼3")
        fun zoo(): String {
            return foo.repeat(2)
        }
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
```