package allbegray.excel

import java.lang.reflect.AccessibleObject
import java.lang.reflect.Field
import java.lang.reflect.Method

data class CellInfo(
    val target: AccessibleObject,
    private val name: String,
    val order: Int,
    val width: Int,
    val autoSize: Boolean
) {
    fun <T> invoke(obj: T): Any? {
        return when (target) {
            is Method -> target.invoke(obj)
            is Field -> target.get(obj)
            else -> throw UnsupportedOperationException()
        }
    }

    fun type(): Class<out Any> {
        return when (target) {
            is Method -> target.returnType
            is Field -> target.type
            else -> throw UnsupportedOperationException()
        }
    }

    fun name(): String {
        return name.ifBlank {
            when (target) {
                is Method -> target.name
                is Field -> target.name
                else -> throw UnsupportedOperationException()
            }
        }
    }

    fun styleName(): String {
        return "${name()}Style"
    }
}
