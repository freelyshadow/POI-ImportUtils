import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel导入-类注解
 *
 * @author liuxu
 * @date 2021/5/6 14:19
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelImportClass {

    /**
     * 起始行号
     */
    int startRowNum() default 1;

}
