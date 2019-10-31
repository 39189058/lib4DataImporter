package com.qlmsoft.lib4.dataimporter;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.text.MessageFormat;
import java.util.Iterator;
import java.util.List;
import java.util.UUID;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.google.common.collect.Lists;
public class ParseDicFromWord {
	
	static String SQLTEMP = "INSERT INTO [dbo].[sys_dict] ([id], [value], [label], [type], [description], [sort], [parent_id], [create_by], [create_date], [update_by], [update_date], [del_flag],[ext1],[ext2],[ext3],[ext4],[ext5])" 
	+"VALUES ({0}, {1}, {2}, {3}, {4}, {5},'0', '1', GETDATE(), '1', GETDATE(),  '0',{6},{7},{8},{9},{10});";
	
	static MessageFormat messageFormat = new MessageFormat(SQLTEMP);

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
					if(text.startsWith("6.2 TBPRINCIPALUNITDIC")){
						doFlag = 1;
					}
					
					if(doFlag == 0){
						continue;
					}
					
					System.out.println(text);
					
					tableNames.add(parseTableName(text));
					tableHints.add(parseTablehint(text));
					
					System.out.println(parseTableName(text));
					
					
					if(text.startsWith("6.30 TBSEISMICINTENSITYSCALEDIC")){
						doFlag = 2;
					}
					
					if(doFlag == 2){
						break;
					}
				}
				
				BufferedWriter writer = new BufferedWriter(new FileWriter(new File("G:\\sys_dict.txt"), true));
				
				Iterator<XWPFTable> it = xwpf.getTablesIterator();// 得到word中的表格
				
				int tableIndex = 0;
				
				int startFlag = 0;

				while (it.hasNext()) {
					
					

					XWPFTable table = it.next();
					
					List<XWPFTableRow> rows = table.getRows();
					
					//System.out.println(" table.getRows()"+ table.getRows());
					
					int cols = rows.get(0).getTableCells().size();
					
					if(cols == 3){
						startFlag = 1;
					}
					
					if(startFlag == 0){
						continue;
					}
					
					
					
					String tablename = tableNames.get(tableIndex);
					String tablehint = tableHints.get(tableIndex);
					
					System.out.println("/*"+tablename+"*/");
					
					System.out.println("delete from sys_dict where type='"+tablename+"';");
					
					writer.append("/*"+tablename+"*/\n");
					writer.append("delete from sys_dict where type='"+tablename+"';\n");
					
					String sql = "";
					// 读取每一行数据, 从1开始，过滤第一行
					for (int i = 1; i < rows.size(); i++) {
						XWPFTableRow row = rows.get(i);
						// 读取每一列数据
						List<XWPFTableCell> cells = row.getTableCells();
						
						sql = genClause(cells, tablename, tablehint);	
						
						writer.append(sql+"\n");

					}
					
					tableIndex ++;
					
					if(tableIndex == tableNames.size()){
						break;
					}
					
				}
				
				writer.close();
				
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
		int index2 = line.indexOf("DIC")+3;
		return line.substring(index1, index2);
	}
	
	private static String parseTablehint(String line){
		int index1 = line.indexOf("DIC");
		int index2 = line.indexOf("表");
		return line.substring(index1+3, index2+1);
	}
	
	private static String handleStr(String str){
		return str.trim().replaceAll("　", "").replaceAll("\\*", "");
	}
	
	private static String genClause(List<XWPFTableCell> cells, String tablename, String tablehint){
		
		String id = UUID.randomUUID().toString().replaceAll("-", "");
		String value = ""; 
		String label = ""; 
		String type = tablename;  
		String description = tablehint;  
		String sort = ""; 
		
		String ext1 = "";
		String ext2 = "";
		String ext3 = "";
		String ext4 = "";
		String ext5 = "";
		
		int cellsSize = cells.size();
		
		if(cellsSize>=2){
			value = handleStr(cells.get(0).getText());
			label = handleStr(cells.get(1).getText());
			sort = value + "0";
		}
		
		if(cellsSize>=3){
			String seq = handleStr(cells.get(0).getText());
			value = handleStr(cells.get(1).getText());
			label = handleStr(cells.get(2).getText());	
			sort = seq + "0";
		}
		
		if(cellsSize>=4){
			ext1 = handleStr(cells.get(3).getText());	
		}
		
		if(cellsSize>=5){
			ext2 = handleStr(cells.get(4).getText());	
		}

		if(cellsSize>=6){
			ext3 = handleStr(cells.get(5).getText());	
		}
		
		String[] params = {
				"'"+id+"'",
				"'"+value+"'",
				"'"+label+"'",
				"'"+type+"'", 
				"'"+description+"'", 
				"'"+sort+"'",
				"'"+ext1+"'",
				"'"+ext2+"'",
				"'"+ext3+"'",
				"'"+ext4+"'",
				"'"+ext5+"'"
				};
		return messageFormat.format(params);
		
	}

}