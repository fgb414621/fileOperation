import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author: fgb
 * @Date: 2023/12/04/19:03
 * @Description:
 */
public class main {
    public static void main(String[] args) {
        String filePath = "G:\\fileOperation\\src\\main\\resources\\Excel文件测试.xls";
        List<EmployeeVo> dataList1 = new ArrayList<>();
        List<EmployeeVo> dataList2 = new ArrayList<>();
        List<EmployeeVo> employeeVos = ExcelUtil.readDataFromExcel(filePath, EmployeeVo.class, 1);
        System.out.print("读取的数据行数：" + employeeVos.size());
        int i = 1;
        for (EmployeeVo employeeVo : employeeVos) {
            if (i <= 50) {
                dataList1.add(employeeVo);
            } else {
                dataList2.add(employeeVo);
            }
            i++;
            System.out.println(employeeVo);
        }

        System.out.println("==================================");

        String sheetName = "员工信息";
        String[] headerArray = new String[]{"员工编号", "员工姓名", "员工年龄", "工资", "部门", "入职时间"};
        String filePath1 = "G:\\fileOperation\\src\\main\\resources\\output\\Excel文件测试1.xls";
        ExcelUtil.exportDataToExcel(new HSSFWorkbook(), sheetName, headerArray, dataList1, filePath1);
        String filePath2 = "G:\\fileOperation\\src\\main\\resources\\output\\Excel文件测试2.xls";
        ExcelUtil.exportDataToExcel(new HSSFWorkbook(), sheetName, headerArray, dataList2, filePath2);


    }

}


