package src;
/**
 * Title:	    Export Excel file from Database using XML input.
 * Class:	    ExportExcelV01
 * Description: Export Excel file from Database using XML input.
 * @author  	ArunDavid
 * @version  	1.0
 *
 */
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.List;
import java.util.Map;
import java.util.HashMap;
import java.util.StringTokenizer;

import org.jdom.Document;
import org.jdom.Element;
import org.jdom.JDOMException;
import org.jdom.input.SAXBuilder;

import org.apache.poi.hssf.dev.HSSF;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;


class ExportTableV1{
	Connection con;
	Statement st;
	int topPadding=1,leftPadding=1;
	int rowCursor,columnCursor;
	List constants;
	Map style;
	
	ExportTableV1(){}
	ExportTableV1(List constants){this.constants=constants;}
	
	String replaceConstants(String name){ 
		for (int i = 0; i <	this.constants.size(); i++) { 
			Element constant = (Element)constants.get(i);
			name=name.replace(constant.getAttributeValue("name"),constant.getText());
		} 
		return name; 
	}
	
	void connectDb(Element database) throws ClassNotFoundException, SQLException{
		System.out.println("Connecting to Database...");
		String dbPath=database.getChild("path").getText();
		String username=database.getChild("username").getText();
		String password=database.getChild("password").getText();
		String dbDriver=database.getChild("driver").getText();
		Class.forName(dbDriver);
		con = DriverManager.getConnection(dbPath,username,password);
		st=con.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE,ResultSet.CONCUR_READ_ONLY);
		System.out.println("Database connected...");
	}
	void closeDb() throws SQLException{
		st.close();
		con.close();
		System.out.println("Database connection closed...");
	}
	void setStyle(Element styleElement){
		System.out.println("Setting up styles...");
		style=new HashMap();
		style.put("headTextColor",styleElement.getChild("headTextColor").getText());
		style.put("titleTextColor",styleElement.getChild("titleTextColor").getText());
		style.put("titleBackgroundColor",styleElement.getChild("titleBackgroundColor").getText());
		style.put("fieldTextColor",styleElement.getChild("fieldTextColor").getText());
		style.put("fieldBackgroundColor",styleElement.getChild("fieldBackgroundColor").getText());
		topPadding=Integer.parseInt(styleElement.getChild("topPadding").getText());
		leftPadding=Integer.parseInt(styleElement.getChild("leftPadding").getText());
		System.out.println("Setting styles completed.");
	}
	void export(String inputFile){
		try {			
			System.out.println("Reading input file '"+inputFile+"'...");
			KeyMethodMapV1 keyMethod=new KeyMethodMapV1();
			SAXBuilder builder = new SAXBuilder();
			File xmlFile = new File(inputFile);
			Document document = (Document) builder.build(xmlFile);
			Element rootNode = document.getRootElement();
			List excels = rootNode.getChildren("excel");
			for (int i = 0; i < excels.size(); i++) {
			   Element excel = (Element) excels.get(i);
			   //System.out.println("Excel : " + excel.getAttributeValue("name"));
			   File file=new File(replaceConstants(excel.getAttributeValue("name")));
			   FileOutputStream fos = new FileOutputStream(file);
			   HSSFWorkbook workbook = new HSSFWorkbook();
				HSSFPalette palette = workbook.getCustomPalette();
				StringTokenizer strTkn = null;
				strTkn=new StringTokenizer(style.get("fieldBackgroundColor").toString(),",");
				palette.setColorAtIndex(HSSFColor.GREY_25_PERCENT.index, (byte)Integer.parseInt(strTkn.nextToken()), (byte)Integer.parseInt(strTkn.nextToken()), (byte)Integer.parseInt(strTkn.nextToken()));
				strTkn=new StringTokenizer(style.get("titleBackgroundColor").toString(),",");
				palette.setColorAtIndex(HSSFColor.GREY_40_PERCENT.index, (byte)Integer.parseInt(strTkn.nextToken()), (byte)Integer.parseInt(strTkn.nextToken()), (byte)Integer.parseInt(strTkn.nextToken()));
				strTkn=new StringTokenizer(style.get("titleTextColor").toString(),",");
				palette.setColorAtIndex(HSSFColor.GREY_50_PERCENT.index, (byte)Integer.parseInt(strTkn.nextToken()), (byte)Integer.parseInt(strTkn.nextToken()), (byte)Integer.parseInt(strTkn.nextToken()));
				strTkn=new StringTokenizer(style.get("fieldTextColor").toString(),",");
				palette.setColorAtIndex(HSSFColor.GREY_80_PERCENT.index, (byte)Integer.parseInt(strTkn.nextToken()), (byte)Integer.parseInt(strTkn.nextToken()), (byte)Integer.parseInt(strTkn.nextToken()));
				strTkn=new StringTokenizer(style.get("headTextColor").toString(),",");
				palette.setColorAtIndex(HSSFColor.RED.index, (byte)Integer.parseInt(strTkn.nextToken()), (byte)Integer.parseInt(strTkn.nextToken()), (byte)Integer.parseInt(strTkn.nextToken()));
				
				HSSFCellStyle headStyle = workbook.createCellStyle();
				HSSFFont headFont=workbook.createFont();
				headFont.setFontName("Calibri");
				headFont.setFontHeightInPoints((short)11);
				headFont.setColor(HSSFColor.RED.index);
				headFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				headStyle.setFont(headFont);		
			 
				HSSFCellStyle titleStyle = workbook.createCellStyle();
				titleStyle.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
				titleStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				HSSFFont titleFont=workbook.createFont();
				titleFont.setFontName("Calibri");
				titleFont.setFontHeightInPoints((short)11);
				titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				titleFont.setColor(HSSFColor.GREY_50_PERCENT.index);
				titleStyle.setFont(titleFont);
				//titleStyle.setFillBackgroundColor(HSSFColor.BLACK.index);
				titleStyle.setBorderLeft((short)1);
				titleStyle.setBorderTop((short)1);
				titleStyle.setBorderRight((short)1);
				titleStyle.setBorderBottom((short)1);
				titleStyle.setBottomBorderColor(HSSFColor.BLACK.index);
				titleStyle.setLeftBorderColor(HSSFColor.BLACK.index);
				titleStyle.setRightBorderColor(HSSFColor.BLACK.index);
				titleStyle.setTopBorderColor(HSSFColor.BLACK.index);
				
				HSSFCellStyle fieldStyle = workbook.createCellStyle();
				fieldStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
				fieldStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
				HSSFFont fieldFont=workbook.createFont();
				fieldFont.setFontName("Calibri");
				fieldFont.setFontHeightInPoints((short)11);
				//titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				fieldFont.setColor(HSSFColor.GREY_80_PERCENT.index);
				fieldStyle.setFont(fieldFont);
				//fieldStyle.setFillBackgroundColor(HSSFColor.BLACK.index);
				fieldStyle.setBorderLeft((short)1);
				fieldStyle.setBorderTop((short)1);
				fieldStyle.setBorderRight((short)1);
				fieldStyle.setBorderBottom((short)1);
				fieldStyle.setBottomBorderColor(HSSFColor.BLACK.index);
				fieldStyle.setLeftBorderColor(HSSFColor.BLACK.index);
				fieldStyle.setRightBorderColor(HSSFColor.BLACK.index);
				fieldStyle.setTopBorderColor(HSSFColor.BLACK.index);
					
					
				List sheets = excel.getChildren("sheet");
				for (int j = 0; j < sheets .size(); j++) {
				   Element sheet= (Element) sheets .get(j);
				   //System.out.println("Sheet  : " + sheet.getAttributeValue("name"));
				   HSSFSheet worksheet = workbook.createSheet(replaceConstants(sheet.getAttributeValue("name")));
				   //rowCursor=topPadding;columnCursor=leftPadding;	   
				   List reports= sheet.getChildren("report");
				   for (int k = 0; k < reports.size(); k++) {
					   Element report= (Element)reports.get(k);
					   if(report.getAttributeValue("name")!=null){
							System.out.println("Generating...\"" + replaceConstants(report.getAttributeValue("name"))+"\"");
					   }
					   if(report.getAttributeValue("type")!=null&&report.getAttributeValue("type").equals("label")){
						   rowCursor+=topPadding;columnCursor=leftPadding;
						   HSSFRow row = worksheet.createRow((short)rowCursor);rowCursor++;
						   HSSFCell cell = row.createCell((short)columnCursor);columnCursor=leftPadding;
						   cell.setCellValue(replaceConstants(report.getAttributeValue("name")));
						   cell.setCellStyle(headStyle);
						   continue;
					   }
					   String query= report.getChild("query").getChildText("sql");
					   List fields= report.getChild("fields").getChildren("field");
					   //System.out.println("SQL : "+query);
					   ResultSet rs=st.executeQuery(replaceConstants(query));
					   rs.last();int numRows = rs.getRow();rs.beforeFirst();
					   if(numRows==0){
						   if(report.getAttributeValue("emptytext")!=null){
							   if(report.getAttributeValue("emptytext").equalsIgnoreCase("HIDDEN")){
								   continue;
							   }else{
								   rowCursor+=topPadding;columnCursor=leftPadding;
								   HSSFRow row = worksheet.createRow((short)rowCursor);rowCursor++;
								   HSSFCell cell = row.createCell((short)columnCursor);columnCursor=leftPadding;
								   cell.setCellValue(report.getAttributeValue("emptytext"));
								   cell.setCellStyle(titleStyle);
								   continue;
							   }
						   }
					   }
					   rowCursor+=topPadding;columnCursor=leftPadding;
					   if(report.getAttributeValue("name")!=null){
						   HSSFRow row = worksheet.createRow((short)rowCursor);rowCursor++;
						   HSSFCell cell = row.createCell((short)columnCursor);columnCursor=leftPadding;
						   cell.setCellValue(replaceConstants(report.getAttributeValue("name")));
						   cell.setCellStyle(headStyle);
					   }
					   HSSFRow rowT = worksheet.createRow((short)rowCursor);rowCursor++;
					   for (int l = 0; l < fields.size(); l++) {
						   Element field= (Element)fields.get(l);
						   //System.out.println("Field : " + field.getAttributeValue("name")+"="+field.getText());
						   HSSFCell cellT = rowT.createCell((short)columnCursor);
						   if(field.getAttributeValue("width")!=null){
							   worksheet.setColumnWidth((short)columnCursor,(short)(256*Integer.parseInt(field.getAttributeValue("width"))));
						   }columnCursor++;
						   cellT.setCellValue(replaceConstants(field.getAttributeValue("name")));   
						   cellT.setCellStyle(titleStyle);
					   }columnCursor=leftPadding;   
					   while(rs.next()){
							HSSFRow rowV = worksheet.createRow((short)rowCursor);rowCursor++;
							for (int l = 0; l < fields.size(); l++) {
								Element field= (Element)fields.get(l);
								//System.out.println("Field : " + field.getAttributeValue("name")+"="+field.getText())
								HSSFCell cellV = rowV.createCell((short)columnCursor);
								if(field.getAttributeValue("type")!=null&&field.getAttributeValue("type").equals("method")){
									cellV.setCellValue(keyMethod.getMethodValue(field.getText(),rs));
								}else{
									if(rs.getString(field.getText())!=null){
										if(field.getAttributeValue("datatype")!=null&&field.getAttributeValue("datatype").equals("number")){
											cellV.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
											cellV.setCellValue(new Double(rs.getString(field.getText())));			   
										}else{
											String cellVal="";
											if(field.getAttributeValue("prefix")!=null&&!field.getAttributeValue	("prefix").equals("")){
												cellVal+=field.getAttributeValue("prefix");
											}
											cellVal+=rs.getString(field.getText());
											if(field.getAttributeValue("suffix")!=null&&!field.getAttributeValue	("suffix").equals("")){
												cellVal+=field.getAttributeValue("suffix");
											}
										   cellV.setCellValue(cellVal);
										}
									}else{
										if(field.getAttributeValue("default")!=null&&!field.getAttributeValue("default").equals("")){
											cellV.setCellValue(field.getAttributeValue("default"));
									    }
									}
								}cellV.setCellStyle(fieldStyle);
								columnCursor++;
							}columnCursor=leftPadding;
						}
					}
			   }
			   workbook.write(fos);
			   fos.flush();
			   fos.close();
			   System.out.println("File is written...");
			}
		} catch (IOException e) {
			System.out.println(e);
		} catch (JDOMException e) {
			System.out.println(e);
		} catch (SQLException e) {
			System.out.println(e);
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}