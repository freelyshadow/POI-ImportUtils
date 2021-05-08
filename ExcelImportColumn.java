import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel导入-列注解
 *
 * @author liuxu
 * @date 2021/4/30 11:55
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelImportColumn {

    /**
     * 列号
     */
    int column() default 0;

}
