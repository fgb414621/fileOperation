import lombok.Data;

import java.time.LocalDate;

/**
 * Created with IntelliJ IDEA.
 *
 * @Author: fgb
 * @Date: 2023/12/04/18:59
 * @Description:
 */
@Data
public class EmployeeVo {

    /**
     * 员工编号
     */
    private String id;

    /**
     * 员工姓名
     */
    private String name;

    /**
     * 员工年龄
     */
    private int age;

    /**
     * 工资
     */
    private double salary;

    /**
     * 部门
     */
    private String department;

    /**
     * 入职时间
     */
    private LocalDate hireDate;

    /**
     * 无参构造函数不能少
     */
    public EmployeeVo() {

    }

    public EmployeeVo(String id, String name, int age, double salary, String department, LocalDate hireDate) {
        this.id = id;
        this.name = name;
        this.age = age;
        this.salary = salary;
        this.department = department;
        this.hireDate = hireDate;
    }
}
