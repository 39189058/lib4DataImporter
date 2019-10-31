package com.qlmsoft.lib4.dataimporter;

import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.google.common.collect.Lists;
import com.qlmsoft.lib4.common.Util;

public class ParseTablesFromWord {

	public static void main(String[] args) {
		handle();

	}
	
	public static void handle(){
		String realPath = "g:/lib4.docx";

		try {
			FileInputStream in = new FileInputStream(realPath);

			if (realPath.toLowerCase().endsWith("docx")) {
				// word 2007 图片不会被读取， 表格中的数据会被放在字符串的最后
				XWPFDocument xwpf = new XWPFDocument(in);// 得到word文档的信息
				List<XWPFParagraph> listParagraphs = xwpf.getParagraphs();//得到段落信息
				
				System.out.println("listParagraphs.size():" + listParagraphs.size());
				
				List<String> tableNames = Lists.newArrayList();
				List<String> tableHints = Lists.newArrayList();
				
				int doFlag = 0;
				for(XWPFParagraph title: listParagraphs){
					String text = title.getParagraphText();
					if(text.startsWith("5.1.1")){
						doFlag = 1;
					}
					
					if(doFlag == 0){
						continue;
					}
					
					tableNames.add(parseTableName(text));
					tableHints.add(parseTablehint(text));
					
					
					if(text.startsWith("5.8.2")){
						doFlag = 2;
					}
					
					if(doFlag == 2){
						break;
					}
				}
				
				StringBuffer buffer = new StringBuffer();
				
				Iterator<XWPFTable> it = xwpf.getTablesIterator();// 得到word中的表格
				
				int tableIndex = 0;

				while (it.hasNext()) {
					
					buffer = new StringBuffer();

					XWPFTable table = it.next();
					
					List<XWPFTableRow> rows = table.getRows();
					
					//System.out.println(" table.getRows()"+ table.getRows());
					
					int cols = rows.get(0).getTableCells().size();
					
					if(cols!=7){
						//不是需要的表，跳过
						continue;
					}
					
					System.out.println("/*"+tableNames.get(tableIndex)+"*/");
					
					String tablename = tableNames.get(tableIndex);
					String tablehint = tableHints.get(tableIndex);
					
					//buffer.append("IF OBJECT_ID(N'"+tablename+"', N'U') IS  NOT  NULL DROP TABLE [dbo].["+tablename+"];").append("\n");
					buffer.append("/*"+tablehint+"*/\n");
					buffer.append("CREATE TABLE [dbo].["+tablename+"] (").append("\n");

					// 读取每一行数据, 从1开始，过滤第一行
					for (int i = 1; i < rows.size(); i++) {
						XWPFTableRow row = rows.get(i);
						// 读取每一列数据
						List<XWPFTableCell> cells = row.getTableCells();
						
						String fieldDefine = genClause(cells, (i==(rows.size()-1)));	
						
						buffer.append(fieldDefine).append("\n");

					}
					
					buffer.append(");").append("\n");
					
					System.out.println(buffer.toString());
					System.out.println("");
					
					tableIndex ++;
					
					if(tableIndex == tableNames.size()){
						break;
					}
					
				}
				
			} else {
				//
				System.out.println("格式错误");
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static String parseTableName(String line){
		int index1 = line.indexOf("TB");
		int index2 = line.indexOf("）");
		return line.substring(index1, index2);
	}
	
	private static String parseTablehint(String line){
		int index1 = line.indexOf(" ");
		int index2 = line.indexOf("）");
		return line.substring(index1, index2+1);
	}
	
	private static String genClause(List<XWPFTableCell> cells, boolean lastOne){
		StringBuffer buffer = new StringBuffer();
		String fieldComment = handleStr(cells.get(0).getText());
		String fieldName = handleStr(cells.get(1).getText());
		String fieldType = handleStr(cells.get(2).getText());
		String fieldLen = handleStr(cells.get(3).getText());
		int digitals = Util.getInteger(cells.get(4).getText().trim());
		String mandatory = handleStr(cells.get(5).getText());
		
		if("VVarchar".equals(fieldType)){
			fieldType = "Varchar";
		}
		
		buffer.append("["+fieldName+"] ");
		if("Int".equals(fieldType) || "int".equals(fieldType) 
		|| "Date".equals(fieldType)  || "date".equals(fieldType)
		|| "Datetime".equals(fieldType)  || "datetime".equals(fieldType)){
			buffer.append(fieldType +" ");
		}else if("Numeric".equals(fieldType) || "Float".equals(fieldType) 
				|| "decimal".equals(fieldType) || "Decimal".equals(fieldType)){	
			if(fieldLen == "" ){
				fieldLen = "10";
			}
			buffer.append("Numeric("+fieldLen+","+digitals+") ");
		}else if("Double".equals(fieldType)){
			buffer.append("varchar(20) ");
		}else if("Image".equals(fieldType)){	
			buffer.append(fieldType +" ");
		}else{
			buffer.append(fieldType+"("+fieldLen+") ");
		}
		
		if("M".equals(mandatory)){
			buffer.append("NOT NULL ");
		}else{
			buffer.append("NULL ");
		}
		
		if(!lastOne){
			buffer.append(",");
		}
		
		buffer.append(" /*"+fieldComment+"*/");
		
		return buffer.toString();
		
	}
	
	private static String handleStr(String str){
		return str.trim().replaceAll("　", "");
	}
	
	/*
CREATE TABLE [dbo].[treset_password_apply] (
[id] varchar(64) NOT NULL ,
[entity_code] varchar(64) NULL ,
[entity_name] varchar(64) NULL ,
[attach] varchar(200) NULL ,
[mobile] varchar(64) NULL ,
[email] varchar(64) NULL ,
[apply_date] datetime NULL ,
[approve_date] datetime NULL ,
[approve_opinion] varchar(100) NULL,
[status] char(1) NULL,
[create_by] varchar(32) NULL ,
[update_by] varchar(32) NULL ,
[create_date] datetime NULL ,
[update_date] datetime NULL ,
[remarks] varchar(200) NULL ,
[del_flag] char(1) NULL,
[proc_ins_id] varchar(50) NULL
);
	 */

}