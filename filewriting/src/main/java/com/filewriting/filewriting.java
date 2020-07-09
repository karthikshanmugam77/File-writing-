package com.filewriting;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.opencsv.CSVWriter;

public class filewriting {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		try {

			Class.forName("com.mysql.cj.jdbc.Driver");
			Connection conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/sakila", "root", "root@123");
			System.out.println("Connection established...");

			PreparedStatement ps = conn.prepareStatement("select * from actor");
			ResultSet rs = ps.executeQuery();

			File f = new File("C:\\Users\\karthikeyan.bala\\Downloads\\write.txt");
			FileWriter fr = new FileWriter(f);

			HSSFWorkbook book = new HSSFWorkbook();
			HSSFSheet sheet = book.createSheet("Excel sheet");
			FileOutputStream fileOut = new FileOutputStream(new File("C:\\Users\\karthikeyan.bala\\Downloads\\ExtractFile.xlsx"));
			HSSFRow rowhead = sheet.createRow(0);
			rowhead.createCell(0).setCellValue("Actor id");
			rowhead.createCell(1).setCellValue("First name");
			rowhead.createCell(2).setCellValue("Last name");
			rowhead.createCell(3).setCellValue("Active status");
			
			XWPFDocument document = new XWPFDocument();
			XWPFParagraph p = document.createParagraph();
			FileOutputStream out2 = new FileOutputStream("C:\\Users\\karthikeyan.bala\\Downloads\\simple.docx");
			XWPFRun r = p.createRun();
			
			CSVWriter writer = new CSVWriter(new FileWriter("C:\\Users\\karthikeyan.bala\\Downloads\\Details.csv"));
			
			int index=0;
			
			while (rs.next()) {

				fr.write(String.format(" ID -> " + rs.getInt(1) + " name -> " + rs.getString(2)));
				fr.write(System.lineSeparator());
				
				HSSFRow row = sheet.createRow(index);
				row.createCell(0).setCellValue(rs.getInt(1));
				row.createCell(1).setCellValue(rs.getString(2));
				row.createCell(2).setCellValue(rs.getString(3));
				row.createCell(3).setCellValue(rs.getTimestamp(4));
				index++;
				
				r.setText(" ID -> " + rs.getInt(1) + " name -> " + rs.getString(2));
				
				String  str[] = new String[2];
				str[0] =  rs.getString(2) ;
				str[1] = rs.getString(3);
				writer.writeNext(str);

			}
			document.write(out2);
			System.out.println("Data is saved in DOCX file.");	
			book.write(fileOut);
			System.out.println("Data is saved in EXCEL file.");	
			fr.close();
			System.out.println("Data is saved in TXT file.");	
			fileOut.close();
			System.out.println("Data is saved in CSV file.");	
			writer.close();
			

		} catch (Exception e) {
			System.out.println(e);
		}
	}

}
