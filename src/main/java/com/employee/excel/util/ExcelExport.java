package com.employee.excel.util;

import com.employee.excel.easypoi.ExceVo;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;

public class ExcelExport {


    private static Logger log = LoggerFactory.getLogger(ExcelExport.class);

    /**
     * 工作薄对象
     */
    private SXSSFWorkbook wb;

    /**
     * 工作表对象
     */
    private Sheet sheet;

    /**
     * 样式列表
     */
    private Map<String, CellStyle> styles;

    /**
     * 当前行号
     */
    private int rownum;

    List<Object[]> annotationList =new ArrayList<>();

    public ExcelExport(String title, Class<?> cls){

        Field[] fields = cls.getDeclaredFields();

        for (Field field:fields){
            ExceVo exceVo=field.getAnnotation(ExceVo.class);
            if (exceVo!=null){
                annotationList.add(new Object[]{exceVo,field});
            }

        }

        Collections.sort(annotationList, new Comparator<Object[]>() {
            @Override
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((ExceVo)o1[0]).sort()).compareTo(new Integer(((ExceVo)o2[0]).sort()));
            }
        });


        List<String > headList=new ArrayList<>();


        for (Object[] os :annotationList){

            String t=((ExceVo)os[0]).name();
            headList.add(t);
        }


        initialize(title,headList);


    }
    private Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<String, CellStyle>();

        CellStyle style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        Font titleFont = wb.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style.setFont(titleFont);
        styles.put("title", style);

        style = wb.createCellStyle();
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font dataFont = wb.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        style.setFont(dataFont);
        styles.put("data", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(CellStyle.ALIGN_LEFT);
        styles.put("data1", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(CellStyle.ALIGN_CENTER);
        styles.put("data2", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
        style.setAlignment(CellStyle.ALIGN_RIGHT);
        styles.put("data3", style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get("data"));
//		style.setWrapText(true);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        Font headerFont = wb.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        style.setFont(headerFont);
        styles.put("header", style);

        return styles;
    }

    private void initialize(String title, List<String> headList) {

        this.wb=new SXSSFWorkbook(500);

        this.sheet = wb.createSheet("Export");
        this.styles = createStyles(wb);

        // Create title
        if (StringUtils.isNotBlank(title)){
            Row titleRow = sheet.createRow(rownum++);
            titleRow.setHeightInPoints(50);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellStyle(styles.get("title"));
            titleCell.setCellValue(title);
            sheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(),
                    titleRow.getRowNum(), titleRow.getRowNum(), headList.size()-1));
        }

        // Create header
        if (headList == null){
            throw new RuntimeException("headerList not null!");
        }
        Row headerRow = sheet.createRow(rownum++);
        headerRow.setHeightInPoints(20);

        for (int i=0;i<headList.size();i++){
            Cell cell=headerRow.createCell(i);
            cell.setCellStyle(styles.get("header"));
            cell.setCellValue(headList.get(i));
         //   sheet.autoSizeColumn(i);
        }


        log.debug("Initialize success.");


    }


    public <E> ExcelExport writeToFile(List<E> e) throws IllegalAccessException {
        int size = 0;
        Class eClass = e.get(0).getClass();
        Field[] fields = eClass.getDeclaredFields();


        List<Object[]> list=new ArrayList<>();
        for (Field field:fields){
            ExceVo exceVo=field.getAnnotation(ExceVo.class);
            if (exceVo!=null){
                list.add(new Object[]{exceVo,field});
            }

        }

        Collections.sort(list, new Comparator<Object[]>() {
            @Override
            public int compare(Object[] o1, Object[] o2) {
                return new Integer(((ExceVo)o1[0]).sort()).compareTo(new Integer(((ExceVo)o2[0]).sort()));
            }
        });


        if (e != null) {
            size = eClass.getDeclaredFields().length;
        }
        for (int i = 0; i < e.size(); i++) {
            Row row = this.sheet.createRow(rownum++);

            for (int g = 0; g < size; g++) {



                Field field=(Field) (list.get(g))[1];
                field.setAccessible(true);
                Cell cell = row.createCell(g);
                cell.setCellStyle(styles.get("data"));
                Object value = field.get(e.get(0));
                cell.setCellValue(value.toString());

            }
        }
        return this;
    }

    public ExcelExport write(HttpServletResponse response, String filename) throws IOException {
        response.reset();
        response.setContentType("application/octet-stream; charset=utf-8");
        response.setHeader("Content-Disposition", "attachment; filename="+Encodes.urlEncode(filename));
        write(response.getOutputStream());

        return this;


    }

    public ExcelExport write(OutputStream outputStream) throws IOException {

        wb.write(outputStream);

        return this;

    }

    public ExcelExport close(){

        wb.dispose();
        return this;

    }


}
