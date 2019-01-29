package com.employee.excel.util;

import com.employee.excel.easypoi.ExceVo;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

public class ExcelImport {

    private static Logger log = LoggerFactory.getLogger(ExcelImport.class);
    /**
     * 工作薄对象
     */
    private Workbook wb;
    /**
     * 工作表对象
     */
    private Sheet sheet;
    /**
     * 标题行号
     */
    private int headerNum;
    /**
     * 获取数据行号
     *
     * @return
     */
    public int getDataRowNum() {
        return headerNum + 1;
    }

    public int getLastDataRowNum() {
        return this.sheet.getLastRowNum() + headerNum;
    }

    /**
     * 获取行对象
     *
     * @param rownum
     * @return
     */
    public Row getRow(int rownum) {
        return this.sheet.getRow(rownum);
    }


    public ExcelImport(MultipartFile file, int headerNum, int sheetNum) throws IOException {

        this(file.getOriginalFilename(),file.getInputStream(),headerNum,sheetNum);
    }
    public ExcelImport(String fileName, InputStream in,int headerNum,int sheetNum) throws IOException {

        if (StringUtils.isBlank(fileName)){

            throw new RuntimeException("Import file is empty!");

        }else if (fileName.toLowerCase().endsWith("xls")) {
            this.wb = new HSSFWorkbook(in);
        } else if (fileName.toLowerCase().endsWith("xlsx")) {
            this.wb = new XSSFWorkbook(in);
        } else {
            throw new RuntimeException("Invalid import file type!");
        }
        if (this.wb.getNumberOfSheets() < sheetNum) {
            throw new RuntimeException("No sheet in Import file!");
        }
        this.sheet = this.wb.getSheetAt(sheetNum);
        this.headerNum = headerNum;
        log.debug("Initialize success.");
    }


    public <E> List<E> getDateList(Class<E> cls) throws IllegalAccessException, InstantiationException {

        List<Object[]> annotationList = new ArrayList<>();

        List<E> dataList = new ArrayList<>();

        Field[] fields = cls.getDeclaredFields();

        for (Field field:fields){
            ExceVo exceVo = field.getAnnotation(ExceVo.class);

            annotationList.add(new Object[]{exceVo,field});
        }

        /**
         * 根据sort排序
         */
        Collections.sort(annotationList, new Comparator<Object[]>() {
            @Override
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((ExceVo)o1[0]).sort()).compareTo(new Integer(((ExceVo)o2[0]).sort()));
            }
        });

        /**
         * 遍历数据存进list
         */
        for (int i=this.getDataRowNum();i<this.getLastDataRowNum();i++){
            int comNum=0;
            E e=(E)cls.newInstance();
            Row row = this.getRow(i);

            for (int g=comNum;g<fields.length;g++){
                Cell cell = row.getCell(g);
                Field field=(Field)(annotationList.get(g))[1];
                field.setAccessible(true);
                field.set(e,cell.getStringCellValue());
            }

            dataList.add(e);
        }

        return dataList;





    }





}
