package com.employee.excel.controller;

import com.employee.excel.pojo.Employee;
import com.employee.excel.util.ExcelExport;
import com.employee.excel.util.ExcelImport;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Controller
@RequestMapping("/employee")
public class EmployeeController {


    @RequestMapping("/index")
    public String index(){
        return "thymeleaf";
    }




    @RequestMapping("/test")
    public void excelExport(HttpServletResponse response ) throws IOException, IllegalAccessException {
        String fileName="Employeedata"+".xlsx";
        Employee employee=new Employee("111111","ceshi","B","PR");
        List<Employee> list=new ArrayList<>();
        for (int i=0;i<10;i++) {
            list.add(employee);
        }

        new ExcelExport(fileName,Employee.class).writeToFile(list).write(response,fileName).close();

    }


    @ResponseBody
    @RequestMapping("/test1")
    public void excelImport(@RequestParam("file") MultipartFile file,
                            HttpServletRequest request) throws IOException, InstantiationException, IllegalAccessException {

        ExcelImport excelImport=new ExcelImport(file,1,0);

        List<Employee> list=  excelImport.getDateList(Employee.class);


        System.out.println(list.toString());




    }









}
