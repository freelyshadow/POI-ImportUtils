import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.lang.reflect.Field;
import java.util.*;

/**
 * Excel导入工具
 *
 * @author liuxu
 * @date 2021/4/30 11:14
 */
public class ExcelImportUtils {

    /**
     * 日志
     */
    private static final Logger log = LoggerFactory.getLogger(ExcelImportUtils.class);

    /**
     * 获取Workbook对象
     *
     * @param multipartFile Spring文件对象
     * @return org.apache.poi.ss.usermodel.Workbook
     * @author liuxu
     * @date 2021/4/30 11:24
     */
    private static Workbook getWorkbook(MultipartFile multipartFile) throws Exception {

        // 原文件名
        String fileName = multipartFile.getOriginalFilename();

        log.info("Excel file {} parsing...,", fileName);

        // 截取文件后缀
        String fileType = fileName.substring(fileName.lastIndexOf("."));

        Workbook workbook;

        // 识别Excel版本
        if (".xls".equals(fileType)) {

            log.info("Excel file parser HSSFWorkbook!");

            workbook = new HSSFWorkbook(multipartFile.getInputStream());

        } else if (".xlsx".equals(fileType)) {

            log.info("Excel file parser XSSFWorkbook!");

            workbook = new XSSFWorkbook(multipartFile.getInputStream());

        } else {

            throw new Exception("请上传excel文件！");

        }

        return workbook;

    }

    /**
     * Excel转Bean集合
     *
     * @param multipartFile Spring文件对象
     * @return java.lang.String
     * @author liuxu
     * @date 2021/4/30 11:29
     */
    public static List<List<?>> excel2BeanList(MultipartFile multipartFile, Class<?>... beanClasses) throws Exception {

        log.info("Start - Excel file to bean list");

        // Sheet页数据
        List<List<?>> sheetDataset = new ArrayList<>();

        //创建Excel工作薄
        Workbook work = getWorkbook(multipartFile);

        // 遍历Sheet
        for (int i = 0; i < work.getNumberOfSheets(); i++) {

            // Sheet对象
            Sheet sheet = work.getSheetAt(i);
            if (sheet == null) {

                log.info("Sheet is null!");

                continue;

            } else if (i >= beanClasses.length) {

                log.info("Sheep number > Bean classes length!");

                continue;

            }

            Map<Integer, String> fieldMap = new HashMap<>();

            Class<?> beanClass = beanClasses[i];

            int startRowNum = beanClass.getAnnotation(ExcelImportClass.class).startRowNum();

            Field[] fields = beanClass.getDeclaredFields();
            for (Field field : fields) {

                // 行号-字段映射
                ExcelImportColumn excelImportColumn = field.getAnnotation(ExcelImportColumn.class);
                if (excelImportColumn != null) {

                    fieldMap.put(excelImportColumn.column(), field.getName());

                }

            }

            // 行数据集合
            List<Object> rowList = new ArrayList<>();

            // 遍历Row
            for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {

                Row row = sheet.getRow(j);
                if (row == null || startRowNum > j) {

                    continue;

                }

                // 遍历Cell
                JSONObject rowJsonData = new JSONObject();
                for (int k = row.getFirstCellNum(); k < row.getLastCellNum(); k++) {

                    Cell cell = row.getCell(k);
                    if (cell != null) {

                        Optional.ofNullable(fieldMap.get(k)).ifPresent(fieldName -> rowJsonData.put(fieldName, cell.toString()));

                    }

                }

                // 行数据
                rowList.add(JSON.toJavaObject(rowJsonData, beanClass));

            }

            // 行数据加入Sheet集合
            sheetDataset.add(rowList);

        }

        log.info("End - Excel file to bean list");

        return sheetDataset;

    }

}
