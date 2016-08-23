package file;

import j.u.XDate;
import j.u.XList;
import j.u.XMap;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

import net.sf.jxls.transformer.XLSTransformer;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
/**
 * excel内容模板：本质就是个JSTL的循环
 * 
 * 	订单号				结算日期				订单可结算时间				卖家ID
<jx:forEach items="${ordersList}" var="row">	
		
${row.order_id}		${row.settle_time}	${row.order_settle_time}	${row.merchant_id}

</jx:forEach>		
 */
public class ExeceltReadWrite {

	public static void writeExcel(){
		String destFilePath = "D:/ee/结算单明细%S.xls";
		String templatePath = "D:/BB/结算单明细模板.xls";
		
		XLSTransformer transformer = new XLSTransformer();
        Map<String, Object> dataMap = new HashMap<String, Object>();
        dataMap.put("startTime","2015-08-01");
        dataMap.put("endTime","2015-012-01");
        dataMap.put("nowTime", "2015-02-01");
        XList dataList = new XList<>();
        XMap one = new XMap();
        one.put("name","张三");
        XMap two = new XMap();
        two.put("name", "李四");
        dataList.add(one);
        dataList.add(two);
        dataMap.put("ordersList", dataList);
		try {
	        String outpurPath = String.format(destFilePath, "_(" + XDate.getNow().getTime() + ")");
			InputStream is = new BufferedInputStream(new FileInputStream(templatePath));
			HSSFWorkbook workbook = transformer.transformXLS(is, dataMap);
		    OutputStream os = new FileOutputStream(outpurPath);
		    workbook.write(os);
		    is.close();
		    os.flush();
		    os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
	public static void readFile(){
		try {
			String srcFile = "/Users/yourmall/Documents/workspace/Buss_Eclipse_work/web.settle.platform.liequ/src/main/resources/template/结算单汇总.xls";
			// 把一张xls的数据表读到wb里
			HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(new File(srcFile)));
			HSSFSheet sheet = wb.getSheetAt(0);
			// 读取第一页,一般一个excel文件会有三个工作表，这里获取第一个工作表来进行操作
			// 循环遍历表sheet.getLastRowNum()是获取一个表最后一条记录的记录号，
			// 如果总共有3条记录，那获取到的最后记录号就为2，因为是从0开始的
			for (int j = 0; j < sheet.getLastRowNum() + 1; j++) {
				// 创建一个行对象
				HSSFRow row = sheet.getRow(j);
				if (row==null)continue;
				// 把一行里的每一个字段遍历出来
				for (short i = 0; i < row.getLastCellNum(); i++) {
					// 创建一个行里的一个字段的对象，也就是获取到的一个单元格中的值
					HSSFCell cell = row.getCell(i);
					System.out.println(cell.getStringCellValue());
					// 在这里我们就可以做很多自己想做的操作了，比如往数据库中添加数据等
					// System.out.println(cell.getRichStringCellValue());
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void main(String[] args) {
		readFile();
	}
}
