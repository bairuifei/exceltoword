import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class MainTest {
    private Configuration configuration = null;

    public MainTest(){
        configuration = new Configuration();
        configuration.setDefaultEncoding("UTF-8");
    }

    public static void main(String[] args) {
        MainTest test = new MainTest();
        //解析excel
        Workbook wb =null;
        Sheet sheet = null;
        Row row = null;
        List<Map<String,String>> list = null;
        String cellData = null;
        String filePath = "C:\\Users\\Administrator\\Desktop\\文档\\广东大数据\\验收-测试用例\\验收-测试用例\\test-ggfw.xlsx";
        String columns[] = {"testNo","testX","testZx","testMd","testTj","testBz","testJg"};
        wb = readExcel(filePath);
        if(wb != null){
            int count = wb.getNumberOfSheets();
            for (int c=0;c<count;c++){
                //----解析excel数据
                list = new ArrayList<Map<String,String>>();
                sheet = wb.getSheetAt(c);
                int rownum = sheet.getPhysicalNumberOfRows();
                row = sheet.getRow(0);
                int colnum = row.getPhysicalNumberOfCells();
                for (int i = 1; i<rownum; i++) {
                    Map<String,String> map = new LinkedHashMap<String,String>();
                    row = sheet.getRow(i);
                    if(row !=null){
                        for (int j=0;j<colnum;j++){
                            cellData = (String) getCellFormatValue(row.getCell(j));
                            map.put(columns[j], cellData);
                        }
                    }else{
                        break;
                    }
                    list.add(map);
                }
                //----数据封装
                Map<String,Object> map = new HashMap<String, Object>();
                List<Map<String,Object>> wordlist = new ArrayList<Map<String,Object>>();
                map.put("title",sheet.getSheetName());
                //行循环
                for (Map<String,String> excelmap : list) {
                    Map<String,Object> dataMap = new HashMap<String, Object>();
                    for (Map.Entry<String,String> entry : excelmap.entrySet()) {
                        //单行列循环
                        dataMap.put(entry.getKey(), entry.getValue());
                    }
                    dataMap.put("testSj", "2018/9/15");
                    wordlist.add(dataMap);
                }
                map.put("list",wordlist);
                //----创建word
                test.createWord(map,sheet.getSheetName());
            }
        }
    }

    public void createWord(Map<String,Object> dataMap,String fileName){
        configuration.setClassForTemplateLoading(this.getClass(), "");//模板文件所在路径
        Template t=null;
        try {
            t = configuration.getTemplate("testyl.ftl"); //获取模板文件
        } catch (IOException e) {
            e.printStackTrace();
        }
        File outFile = new File("D:/outfile/公共服务/"+fileName+".doc"); //导出文件
        Writer out = null;
        try {
            out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile)));
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }

        try {
            t.process(dataMap, out); //将填充数据填入模板文件并输出到目标文件
        } catch (TemplateException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //读取excel
    public static Workbook readExcel(String filePath){
        Workbook wb = null;
        if(filePath==null){
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if(".xls".equals(extString)){
                return wb = new HSSFWorkbook(is);
            }else if(".xlsx".equals(extString)){
                return wb = new XSSFWorkbook(is);
            }else{
                return wb = null;
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }
    public static Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        if(cell!=null){
            //判断cell类型
            switch(cell.getCellType()){
                case Cell.CELL_TYPE_NUMERIC:{
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case Cell.CELL_TYPE_FORMULA:{
                    //判断cell是否为日期格式
                    if(DateUtil.isCellDateFormatted(cell)){
                        //转换为日期格式YYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                    }else{
                        //数字
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING:{
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }
}
