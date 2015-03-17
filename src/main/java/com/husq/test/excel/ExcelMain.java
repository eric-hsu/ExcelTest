package com.husq.test.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.write.Boolean;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ExcelMain {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		//createExcel();
		getExcel();
	}
	
	
	public static void createExcel(){
		try {
			//1 创建工作簿
			File file = new File("e:excel.excel");
			FileOutputStream os = new FileOutputStream(file);
			WritableWorkbook wwb = Workbook.createWorkbook(os);
			
			//2 创建工作表
			WritableSheet sheet = wwb.createSheet("sheet1", 0);
			
			//3 创建标签

			//字符串
			Label label = new Label(10,10,"area1");
			sheet.addCell(label);
			//数字
			Number number = new Number(10,11,22);
			sheet.addCell(number);
			//boolean类型
			jxl.write.Boolean bool = new Boolean(10,12,true);
			sheet.addCell(bool);
			//填充日期
			SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			String now = format.format(new Date());
			Label label1 = new Label(10,13,now);
			sheet.addCell(label1);
			
			//4 合并单元格
			sheet.mergeCells(0, 24, 1, 24);
			
			//5 添加单元格样式
			WritableCellFormat wc = new WritableCellFormat();
			//设置居中
			wc.setAlignment(Alignment.CENTRE);
			//设置边框线
			wc.setBorder(Border.ALL, BorderLineStyle.THIN);
			//设置单元格的背景颜色
			wc.setBackground(Colour.RED);
			Label label2 = new Label(10,14,"添加单元格样式",wc);
			sheet.addCell(label2);
			
			//6 设置单元格字体
			WritableFont font = new WritableFont(WritableFont.createFont("楷体"),20);
			WritableCellFormat wc1 = new WritableCellFormat(font);
			Label label3 = new Label(10,15,"楷体",wc1);
			sheet.addCell(label3);
			
			wwb.write();
			wwb.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public static void getExcel(){
		try {
			File file = new File("e:excel.excel");
			FileInputStream fis = new FileInputStream(file);
			Workbook wb = Workbook.getWorkbook(fis);
			Sheet sheet = wb.getSheet(0);
			List list = new ArrayList();
			int rows = sheet.getRows();
			int columns = sheet.getColumns();
			System.out.println("表格tag1标签有"+rows+"行"+columns+"列！");
			for(int i=0;i<rows;i++){
				for(int j=0;j<columns;j++){
					
					Cell cell = sheet.getCell(j,i);
					String str = cell.getContents();
					if(!"".equals(str)){
						System.out.println("第"+i+"行，第"+j+"列："+str);
						list.add(str);
					}
				}
			}
			
			/*Cell cell = sheet.getCell(10,10);
			String values = cell.getContents();
			System.out.println(values);*/
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}

}
