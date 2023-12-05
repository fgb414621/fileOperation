import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author: fgb
 * @Date: 2023/12/04/18:56
 * @Description:
 */
@Slf4j
public class ExcelUtil {

    /**
     * 从Excel读取数据返回对象集合
     *
     * @param filePath         文件路径
     * @param t                java对象
     * @param filterRowNumbers 过滤行数（比如数据的头部一行就不是数据）
     * @param <T>              泛型
     * @return List<T>
     */
    public static <T> List<T> readDataFromExcel(String filePath, Class<T> t, int filterRowNumbers) {
        FileInputStream fis = null;
        ArrayList<T> list = new ArrayList<>();
        try {
            fis = new FileInputStream(filePath);
            Workbook sheets = WorkbookFactory.create(fis);
            // 获取sheet第一页（根据自己需要）
            Sheet sheet = sheets.getSheetAt(0);
            // 获取表格的行数
            int totalRowNumber = sheet.getPhysicalNumberOfRows();
            // 获取对象的字段列表
            Field[] fields = t.getDeclaredFields();
            for (int i = filterRowNumbers; i < totalRowNumber; i++) {
                // 获取一行数据
                Row row = sheet.getRow(i);
                // 实例对象放到循环内（这个bug是网友【starwenran】提到的，网址【https://blog.csdn.net/weixin_42345741】）
                T obj = t.newInstance();
//              //变量一行数据的每个单元格，row.getPhysicalNumberOfCells()是单元格的数量
                for (int j = 0, jLen = row.getPhysicalNumberOfCells(); j < jLen; j++) {
                    Cell cell = row.getCell(j);
                    // 获取字段
                    Field dataField = fields[j];
                    String startWord = dataField.getName().substring(0, 1);
                    // 需要注意的是如果变量是is开头的Boolean类型,它的set方法不能用下面的，因为它的set方法是去掉is的方法
                    // 比如isEnable，set方法是setEnable
                    String methodName = "set" + dataField.getName().replaceFirst(startWord, startWord.toUpperCase());
                    // 获取字段的set方法，dataField.getType()是参数的类型
                    Method method = t.getMethod(methodName, dataField.getType());
                    // 反射调用set方法，getValueFromCell是把表格的值转成对应的类型
                    method.invoke(obj, getValueFromCell(dataField, cell));
                }
                list.add(obj);
            }
            return list;
        } catch (Exception e) {
            log.error("从Excel读取数据异常", e);
            return Collections.emptyList();
        } finally {
            try {
                if (fis != null) {
                    fis.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 把数据导入到Excel（支持HSSFWorkbook、XSSFWorkbook、SXSSFWorkbook）
     *
     * @param workbook    工作薄
     * @param sheetName   sheet名称
     * @param headerArray 头部标签数据
     * @param list        数据list
     * @param filePath    数据导出路径
     * @param <T>         泛型
     */
    public static <T> void exportDataToExcel(Workbook workbook, String sheetName, String[] headerArray, List<T> list, String filePath) {
        // 生成一个表格,并命名
        Sheet sheet = workbook.createSheet(sheetName);
        // 设置表格默认列宽15个字节
        sheet.setDefaultColumnWidth(15);
        // 生成一个头部样式
        CellStyle headerStyle = getCellStyle(workbook, true);
        // 生成表格标题
        Row headerRow = sheet.createRow(0);
        headerRow.setHeight((short) 300);

        for (int i = 0, len = headerArray.length; i < len; i++) {
            // 创建头部行的一个小单元格
            Cell headerRowCell = headerRow.createCell(i);
            // 设置头部单元格的样式
            headerRowCell.setCellStyle(headerStyle);
            // 设置头部单元格的值
            headerRowCell.setCellValue(headerArray[i]);
        }

        // 获取数据域样式
        CellStyle bodyStyle = getCellStyle(workbook, false);
        FileOutputStream os = null;
        try {
            // 将数据放入sheet中
            for (int i = 0, iLen = list.size(); i < iLen; i++) {
                // 创建一行，因为头部已经占用一行故需要加1
                Row dataRow = sheet.createRow(i + 1);
                T t = list.get(i);
                // 利用反射，根据JavaBean属性的先后顺序，动态调用get方法得到属性的值
                Field[] fields = t.getClass().getDeclaredFields();
                try {
                    for (int j = 0, jLen = fields.length; j < jLen; j++) {
                        // 获取单元格第值
                        Cell dataRowCell = dataRow.createCell(j);
                        // 获取字段
                        Field dataField = fields[j];
                        String startWord = dataField.getName().substring(0, 1);
                        // 需要注意的是如果变量是is开头的Boolean类型,它的get方法不能用下面的，因为它的get方法是去掉is的方法
                        // 比如isEnable，get方法是getEnable
                        String methodName = "get" + dataField.getName().replaceFirst(startWord, startWord.toUpperCase());
                        // 获取对象的get方法
                        Method getMethod = t.getClass().getMethod(methodName);
                        // 反射调用get方法
                        Object value = getMethod.invoke(t);
                        // 单元格值为String
                        dataRowCell.setCellValue(null == value ? "" : value.toString());
                        dataRowCell.setCellStyle(bodyStyle);
                    }
                } catch (Exception e) {
                    log.error("第【{}】行数据生成异常(下标0开始)", i, e);
                }
            }
            os = new FileOutputStream(filePath);
            workbook.write(os);
            os.flush();
        } catch (Exception e) {
            log.error("生成数据异常", e);
        } finally {
            try {
                if (os != null) {
                    os.close();
                }
            } catch (IOException e) {
                log.error("关闭文件异常", e);
            }
        }
    }

    /**
     * 获取单元格样式
     *
     * @param workbook 工作薄
     * @param isHeader 是否是头部标签
     * @return CellStyle
     */
    public static CellStyle getCellStyle(Workbook workbook, boolean isHeader) {
        CellStyle style = workbook.createCellStyle();
        // 设置边框
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        // 设置边框颜色
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        // 水平对齐方式
        style.setAlignment(HorizontalAlignment.CENTER);
        // 垂直对齐方式
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // 设置字体样式
        Font font = workbook.createFont();
        font.setColor(IndexedColors.BLACK.getIndex());
        font.setFontHeightInPoints((short) 12);
        if (isHeader) {
            // 设置背景色
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            font.setFontHeightInPoints((short) 14);
            font.setBold(true);
        }
        // 把字体应用到当前样式
        style.setFont(font);
        return style;
    }

    /**
     * 把数据转成对象字段相应的类型（待完善：其他类型的判断，默认值及为空判断）
     *
     * @param dataField
     * @param cell
     * @return
     */
    public static Object getValueFromCell(Field dataField, Cell cell) {
        String fieldTypeStr = dataField.getType().toString();
        if (fieldTypeStr.contains("String")) {
            return cell.getStringCellValue();
        } else if (fieldTypeStr.contains("Integer") || fieldTypeStr.contains("int")) {
            return Integer.parseInt(cell.getStringCellValue());
        } else if (fieldTypeStr.contains("Boolean") || fieldTypeStr.contains("boolean")) {
            return Boolean.getBoolean(cell.getStringCellValue());
        } else if (fieldTypeStr.contains("Double") || fieldTypeStr.contains("double")) {
            return Double.parseDouble(cell.getStringCellValue());
        } else if (fieldTypeStr.contains("float")) {
            return Float.parseFloat(cell.getStringCellValue());
        } else if (fieldTypeStr.contains("Long") || fieldTypeStr.contains("long")) {
            return Long.parseLong(cell.getStringCellValue());
        } else if (fieldTypeStr.contains("char")) {
            return cell.getStringCellValue().charAt(0);
        } else if (fieldTypeStr.contains("LocalTime")) {
            return LocalTime.parse(cell.getStringCellValue());
        } else if (fieldTypeStr.contains("LocalDate")) {
            return LocalDate.parse(cell.getStringCellValue());
        } else if (fieldTypeStr.contains("LocalDateTime")) {
            return LocalDateTime.parse(cell.getStringCellValue());
        } else if (fieldTypeStr.contains("Date")) {
            try {
                return new SimpleDateFormat("yyyy-MM-dd").parse(cell.getStringCellValue());
            } catch (ParseException e) {
                log.error("时间转化异常", e);
            }
        }
        return null;
    }

}
