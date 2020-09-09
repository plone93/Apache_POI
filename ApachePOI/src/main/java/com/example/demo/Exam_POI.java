package com.example.demo;

import java.io.FileOutputStream;
import java.util.Calendar;

import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Exam_POI {
	public static void main(String[] args) {
		System.out.println("POI실행");
		Exam_POI poi= new Exam_POI();
		poi.program();
		 
			/*
			 * row : 행
			 * column : 열
			 * */
	}
	
	public void program() {
		System.out.println("프로그램 실행");
		String format = "xls";
		
		/*워크북 생성*/
		Workbook workbook = createWorkbook(format);
		
		/*워크북안에 시트 생성*/
		Sheet sheet = workbook.createSheet("Test Sheet");
		
		/*시트에서 셀 취득*/
		Cell cell = getCell(sheet, 0, 0);
		cell.setCellValue("Test POI");/*셀에 데이터 작성*/
		
		cell = getCell(sheet, 0, 1);
		cell.setCellValue("100");
		
		cell = getCell(sheet, 0, 2);
		cell.setCellValue(Calendar.getInstance().getTime());
		
		/*셀에 데이터 포맷 저장*/
		CellStyle style = workbook.createCellStyle();
		
		/*날짜 포맷*/
		style.setDataFormat(HSSFDataFormat.getBuiltinFormat("m/d/yy h:mm"));
		
		/*정렬 포맷*/
		//style.setAlignment(HorizontalAligment.CENTER);
		//style.setVerticalAlignment(VerticalAligment.TOP);
		
		/*셀 색 지정*/
		style.setFillBackgroundColor(IndexedColors.BLACK.index);
		
		/*폰트 설정*/
		Font font = workbook.createFont();
		font.setColor(IndexedColors.RED.index);
		cell.setCellStyle(style);
		
		/*셀 너비 자동 지정*/
		sheet.autoSizeColumn(0);
		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(2);
		
		cell = getCell(sheet, 1, 0);
		cell.setCellValue(1);
		
		cell = getCell(sheet, 1, 1);
		cell.setCellValue(2);
		
		/*함수식*/
		cell.setCellFormula("SUM(A2:B2)");
		writeExcel(workbook, "c:\\test\\test."+format);
	}
	
	/*워크북 생성*/
	public Workbook createWorkbook(String format) {
		
		if("xls".equals(format)) {/*표준 xls*/
			return new HSSFWorkbook();
		} else if("xlsx".equals(format)) {/*확장 xlsx*/
			return new HSSFWorkbook();
		}
		throw new NoClassDefFoundError();
	}
	
	/*시트로 부터 row(행)을 취득, 생성*/
	public Row getRow(Sheet sheet, int rowNum) {
		Row row = sheet.getRow(rowNum);
		if(row == null) {
			row = sheet.createRow(rowNum);
		}
		return row;
	}
	
	/*Row로 부터 Cell을 취득, 생성하기*/
	public Cell getCell(Row row, int cellNum) {
		Cell cell = row.getCell(cellNum);
		if(cell == null) {
			cell = row.createCell(cellNum);
		}
		return cell;
	}
	
	public Cell getCell(Sheet sheet, int rowNum, int cellNum) {
		Row row = getRow(sheet, rowNum);
		return getCell(row, cellNum);
	}
	
	public void writeExcel(Workbook workbook, String filePath) {
		try (FileOutputStream stream = new FileOutputStream(filePath)){
			workbook.write(stream);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
