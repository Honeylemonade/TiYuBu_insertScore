

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.net.URLEncoder;

/**
 * @description
 * 1.用于将汉字转换为URL编码(GB2312)
 * 2.读写Excel内容
 *
 * @author XvYanPeng
 * @date   2019/4/22
 */
public class GetGB2312 {
    /**
     * @params []
     * @return void
     * @description 字符串转为url编码
     */
    @Test
    public void getURLGB2312() throws Exception {
        String mytext = java.net.URLEncoder.encode("1819_2_00001755_7_1_周二_3/4节","gb2312");
        System.out.println(mytext);
    }
/**
 * @params
 * @return void
 * @description 写入Excel
 */
    @Test
    public void writeExcel() throws Exception {
        //创建工作簿
        XSSFWorkbook workbook=new XSSFWorkbook();
        //创建工作表
        XSSFSheet sheet = workbook.createSheet();
        //创建行
        XSSFRow row = sheet.createRow(2);
        //创建单元格,操作第三行第三列
        XSSFCell cell = row.createCell(2, CellType.STRING);
        cell.setCellValue("hellword");
        //
        FileOutputStream outputStream = new FileOutputStream(new File("test.xlsx"));
        workbook.write(outputStream);
        //关闭工作簿
        workbook.close();
    }
    /**
     * @description 读取Excel
     *
     * @params []
     * @return void
     * @date   2019/4/22
     */
    @Test
    public void readExcel() throws Exception {
        //打开需要读取的文件
        FileInputStream inputStream = new FileInputStream(new File("allmap.xlsx"));
        //读取工作簿
        XSSFWorkbook wordBook = new XSSFWorkbook(inputStream);
        //读取工作表,从0开始
        XSSFSheet sheet = wordBook.getSheetAt(0);
        //读取第三行
        XSSFRow row = sheet.getRow(1);
        //读取单元格
        XSSFCell cell = row.getCell(1);//获取单元格对象
        String value = cell.getStringCellValue();
        System.out.println(value);
        //关闭输入流
        inputStream.close();
        //关闭工作簿
        wordBook.close();
    }
/**
 * @description 读取一个纯中文的Excel，转换为URl编码(GB2312)，并导出另一个Excel
 *
 * @params []
 * @return void
 * @date   2019/4/22
 */
    @Test
    public  void transfer() throws IOException {
        File file=new File("allmap.xlsx");
        //打开需要读取的文件
        FileInputStream inputStream = new FileInputStream(file);
        //读取工作簿
        XSSFWorkbook wordBook = new XSSFWorkbook(inputStream);
        FileOutputStream outputStream = new FileOutputStream(file);
        //读取工作表,从0开始
        XSSFSheet sheet = wordBook.getSheetAt(0);
        XSSFRow row;
        XSSFCell cell;
        String ZHStr;
        String URLStr;
        //循环写入wordBook
        for(int i=0;i<8731;i++){
            //读取第i+1行
            row = sheet.getRow(i+1);
            //读取第一列单元格
            cell = row.getCell(1);
            //获取中文内容
            ZHStr = cell.getStringCellValue();
            //转为URL
            URLStr = java.net.URLEncoder.encode(ZHStr,"gb2312");
            //写入相应位置
            cell.setCellValue(URLStr);
            System.out.println("完成第"+(i+1)+"数据的替换");
        }
        //将wordBook写入excel
        wordBook.write(outputStream);
        //关闭操作流
        inputStream.close();
        outputStream.close();
        //关闭工作簿
        wordBook.close();
    }
}
