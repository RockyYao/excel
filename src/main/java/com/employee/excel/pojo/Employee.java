package com.employee.excel.pojo;

import com.employee.excel.easypoi.ExceVo;
import lombok.Data;
import lombok.ToString;


@Data
@ToString
public class Employee {

    @ExceVo(name = "工号",sort = 1)
    private String 工号;
    @ExceVo(name = "姓名",sort = 2)
    private String 姓名;
    @ExceVo(name = "职级",sort = 4)
    private String 职级;
    @ExceVo(name = "部门",sort = 3)
    private String 部门;

    public Employee(String 工号, String 姓名, String 职级, String 部门) {
        this.工号 = 工号;
        this.姓名 = 姓名;
        this.职级 = 职级;
        this.部门 = 部门;
    }

    public Employee() {
    }
}
