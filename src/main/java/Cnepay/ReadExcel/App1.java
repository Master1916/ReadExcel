package Cnepay.ReadExcel;

import java.io.DataOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App1 {
	public static void main(String[] args) {
		getDataFromExcel("D:\\work\\去激活码\\线上数据\\3台去激活码设备列表.xlsx","D:\\888.sql");
	}
	
	 /**
	  * 
     * 读取出filePath中的所有数据信息
     * 
     * @param filePath excel文件的绝对路径
     * @param outfilepath sql文件路径
     * 
     */
    public static void getDataFromExcel(String filePath,String outfilepath)
    {
      
        //判断是否为excel类型文件
        if(!filePath.endsWith(".xls")&&!filePath.endsWith(".xlsx"))
        {
            System.out.println("文件不是excel类型");
        }
        
        FileInputStream fis =null;
        FileOutputStream out =null;
        File file=new File(outfilepath);//抽象
        Workbook wookbook = null;
        DataOutputStream dos=null;
        try
        {
            //获取一个绝对地址的流
              fis = new FileInputStream(filePath);
              out = new FileOutputStream(file);
              dos=new DataOutputStream(out);
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
       
        try 
        {
            //2003版本的excel，用.xls结尾
        	 wookbook = new XSSFWorkbook(fis);//得到工作簿
             
        } 
        catch (Exception ex) 
        {
             
                ex.printStackTrace();
             
        }
        
        //得到一个工作表
        Sheet sheet = wookbook.getSheetAt(0);
        
        //获得表头
        Row rowHead = sheet.getRow(0);
        
        //判断表头是否正确
        if(rowHead.getPhysicalNumberOfCells() != 4)
        {
            System.out.println("表头的数量不对!");
        }
        
        //获得数据的总行数
        int totalRowNum = sheet.getLastRowNum();
        System.out.println("totalRowNum："+totalRowNum);
        
        //要获得属性
        int number = 0;
        String code = "";
        String ksn = "";
        String Server = "";
        int flag=0;
        StringBuffer sb=new StringBuffer();
       //获得所有数据
        for(int i = 1 ; i <= totalRowNum ; i++)
        {
            //获得第i行对象
            Row row = sheet.getRow(i);
            
            //序号
            Cell cell = row.getCell((short)0);
            
             number = (int) cell.getNumericCellValue();
            //激活码
            cell = row.getCell((short)1);
            code = cell.getStringCellValue().toString();
            
            //ksn
            cell = row.getCell((short)2);
            ksn = cell.getStringCellValue().toString();
            
            //服务商
            cell = row.getCell((short)3);
            Server = cell.getStringCellValue().toString();
           
            if(number!=0){
            	flag++;
            	sb.append("insert into TMP_SERIAL_NUMBER_TERMINAL(id, TRACE_NO, serial_number, ksn_no, agency_name) values (SEQ_TMPSERIALNUMBERTERMINAL.nextval,'"+number+"','"+code+"','"+ksn+"','"+Server+"');");
            	sb.append("\n");
            	//System.out.println("insert into TMP_SERIAL_NUMBER_TERMINAL(id, TRACE_NO, serial_number, ksn_no, agency_name) values (SEQ_TMPSERIALNUMBERTERMINAL.nextval,'"+number+"','"+code+"','"+ksn+"','"+Server+"')");
            	
            }
        }
        
        try {
			dos.writeChars(sb.toString());
		} catch (IOException e) {
			e.printStackTrace();
		}
        System.out.println("文件一共存在数据："+totalRowNum+"条。实际读出数据："+flag);
    }
}
