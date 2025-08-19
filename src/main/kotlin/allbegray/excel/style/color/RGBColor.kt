package allbegray.excel.style.color

import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap
import org.apache.poi.xssf.usermodel.XSSFColor
import java.awt.Color

data class RGBColor(val red: Int, val green: Int, val blue: Int) {

    init {
        if (listOf(red, green, blue).any { it !in 0..255 }) {
            throw IllegalArgumentException(String.format("Wrong RGB(%s, %s, %s)", red, green, blue))
        }
    }

    fun toXSSFColor(): XSSFColor {
        val color = Color(red, green, blue)
        return XSSFColor(color, DefaultIndexedColorMap())
    }
}
