package cn.h.jar.reportexcel.utils;

import org.apache.poi.hssf.usermodel.*;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.List;
import java.util.Map;

/**
 * excel工具
 */
public class ExcelUtils {

    public static void generateWrite(List<Map<String,Object>> mapList,String fileName,String sheetName,HttpServletResponse response) throws Exception {
        if(sheetName==null||sheetName.length()<1) throw new Exception("sheet名称不能为空");
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();

        /**
         * 创建一个sheet
         */
        HSSFSheet sheet = hssfWorkbook.createSheet(sheetName);
        sheet.setColumnWidth(0, 20 * 256);
        sheet.setDefaultColumnWidth(20);
        sheet.setDefaultRowHeightInPoints(20);
        /**
         * 创建单元格，并设置值表头 设置表头居中
         */
        HSSFCellStyle style = hssfWorkbook.createCellStyle();
        /**
         * 创建一个居中格式
         */
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        HSSFFont font = hssfWorkbook.createFont();
        //font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        style.setFont(font);

        /**
         * row样式1
         */
        HSSFCellStyle style1 = hssfWorkbook.createCellStyle();
        HSSFFont font1 = hssfWorkbook.createFont();
        font1.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        /**
         * row样式2
         */
        HSSFCellStyle style2 = hssfWorkbook.createCellStyle();
        style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        HSSFFont font2 = hssfWorkbook.createFont();
        font2.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        style2.setFont(font2);

        HSSFRow row;HSSFCell cell;int cellIndex;int rowIndex=0;
        for (Map<String,Object> mapRow:mapList)
        {
            row = sheet.createRow(rowIndex);
            cellIndex = 0;
            for(Map.Entry<String, Object> entry : mapRow.entrySet()){
                System.out.println(entry.getKey()+":"+entry.getValue());
                cell = row.createCell(cellIndex);
                cell.setCellValue(entry.getValue()+"");
                if(rowIndex==0)
                {
                    cell.setCellStyle(style2);
                }
                else
                {
                    cell.setCellStyle(style1);
                }
                cellIndex++;
            }
            rowIndex++;
        }

        write(hssfWorkbook,fileName,null,response);
    }


    /**
     * 输出流
     * @param hssfWorkbook
     * @param name
     * @param suffix  后缀
     * @param response
     * @throws IOException
     */
    public static void write(HSSFWorkbook hssfWorkbook,String name,String suffix, HttpServletResponse response) throws IOException {

        if(suffix==null) suffix="xls";

        OutputStream os = null;
        //将excel的数据写入文件
        response.setContentType("octets/stream");
        String excelName = name;//文件名字
        //转码防止乱码
        try {
            response.addHeader("Content-Disposition", "attachment;filename=" + new String(excelName.getBytes("gb2312"), "ISO8859-1") + "."+suffix);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }

        try {
            os = response.getOutputStream();
            hssfWorkbook.write(os);
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("生成表格出错");
        } finally {
            if (os != null) {
                os.flush();
                os.close();
            }
        }

    }

}
