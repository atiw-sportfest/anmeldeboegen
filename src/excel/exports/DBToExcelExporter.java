package excel.exports;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Comparator;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFConditionalFormattingRule;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DBToExcelExporter {

	private static int z1 = 0;
	private static int z2 = 0;
	private static int zA = 0;

	public static void export(OutputStream os, String klasse, List<DBToExcelSchueler> schueler, List<DBToExcelDisziplin> disziplinen) throws IOException {

		disziplinen.sort(new Comparator<DBToExcelDisziplin>() {
			@Override
			public int compare(DBToExcelDisziplin o1, DBToExcelDisziplin o2) {
				return o1.compareTo(o2);
			}
		});

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Anmeldebogen");
		sheet.protectSheet("fs151");

		disziplinen(workbook, sheet, disziplinen);
		schueler(workbook, sheet, schueler);
		auswahl(workbook, sheet, schueler, disziplinen);
		bewertungSchueler(workbook, sheet, schueler);
		bewertungDisziplinen(workbook, sheet, schueler, disziplinen);
		extras(workbook, sheet, schueler, disziplinen);

		sheet.getRow(0).createCell(1).setCellValue(klasse);

		try {
			workbook.write(os);
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		z1 = 0; 
		z2 = 0; 
		zA = 0; 

		System.out.println("Done");
	}

	private static void disziplinen(XSSFWorkbook workbook, XSSFSheet sheet, List<DBToExcelDisziplin> disziplinen) {
		XSSFCellStyle rotationStyleOrange = workbook.createCellStyle();
		rotationStyleOrange.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 255, (byte) 230, (byte) 153 }));
		rotationStyleOrange.setRotation((short) 90);
		rotationStyleOrange.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		rotationStyleOrange.setBorderTop(BorderStyle.THIN);
		rotationStyleOrange.setBorderBottom(BorderStyle.THICK);
		rotationStyleOrange.setBorderLeft(BorderStyle.THIN);
		rotationStyleOrange.setBorderRight(BorderStyle.THIN);

		XSSFCellStyle rotationStyleGreen = workbook.createCellStyle();
		rotationStyleGreen.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 198, (byte) 224, (byte) 180 }));
		rotationStyleGreen.setRotation((short) 90);
		rotationStyleGreen.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		rotationStyleGreen.setBorderTop(BorderStyle.THIN);
		rotationStyleGreen.setBorderBottom(BorderStyle.THICK);
		rotationStyleGreen.setBorderLeft(BorderStyle.THIN);
		rotationStyleGreen.setBorderRight(BorderStyle.THIN);

		sheet.createRow(2);
		sheet.createRow(3);
		for (int j = 0; j < disziplinen.size(); j++) {
			Cell cell1 = sheet.getRow(2).createCell(j + 3);
			if (disziplinen.get(j).isMannschaft()) {
				cell1.setCellStyle(rotationStyleGreen);
				z2++;
			} else {
				cell1.setCellStyle(rotationStyleOrange);
				z1++;
			}
			cell1.setCellValue(disziplinen.get(j).getName());

			Cell cell0 = sheet.getRow(3).createCell(j + 3);
			cell0.setCellValue(disziplinen.get(j).getId());

			sheet.autoSizeColumn(j + 3);
		}

		sheet.getRow(3).setHeight((short) 0);
		sheet.setColumnWidth(0, (short) 0);

		XSSFCellStyle styleOrange = workbook.createCellStyle();
		styleOrange.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 255, (byte) 230, (byte) 153 }));
		styleOrange.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleOrange.setAlignment(HorizontalAlignment.CENTER);
		styleOrange.setVerticalAlignment(VerticalAlignment.CENTER);
		styleOrange.setBorderTop(BorderStyle.THIN);
		styleOrange.setBorderBottom(BorderStyle.THIN);
		styleOrange.setBorderLeft(BorderStyle.THIN);
		styleOrange.setBorderRight(BorderStyle.THIN);

		XSSFCellStyle styleGreen = workbook.createCellStyle();
		styleGreen.setFillForegroundColor(new XSSFColor(new byte[] { (byte) 198, (byte) 224, (byte) 180 }));
		styleGreen.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleGreen.setAlignment(HorizontalAlignment.CENTER);
		styleGreen.setVerticalAlignment(VerticalAlignment.CENTER);
		styleGreen.setBorderTop(BorderStyle.THIN);
		styleGreen.setBorderBottom(BorderStyle.THIN);
		styleGreen.setBorderLeft(BorderStyle.THIN);
		styleGreen.setBorderRight(BorderStyle.THIN);

		if (z1 > 0) {
			CellRangeAddress cra = new CellRangeAddress(1, 1, 3, 3 + z1 - 1);
			sheet.addMergedRegion(cra);
		}

		if (z2 > 0) {
			CellRangeAddress cra = new CellRangeAddress(1, 1, 3 + z1, 3 + z1 + z2 - 1);
			sheet.addMergedRegion(cra);
		}

		sheet.createRow(0);
		sheet.createRow(1);
		for (int i = 3; i < 3 + z1; i++) {
			Cell headercell = sheet.getRow(1).createCell(i);
			headercell.setCellStyle(styleOrange);
			headercell.setCellValue("Individualsportarten");
		}

		for (int i = 3 + z1; i < 3 + z1 + z2 - 1; i++) {
			Cell headercell = sheet.getRow(1).createCell(i);
			headercell.setCellStyle(styleGreen);
			headercell.setCellValue("Manschaftssportarten");
		}
	}

	private static void schueler(XSSFWorkbook workbook, XSSFSheet sheet, List<DBToExcelSchueler> schueler) {
		XSSFCellStyle styleSchueler = workbook.createCellStyle();
		styleSchueler.setAlignment(HorizontalAlignment.CENTER);
		styleSchueler.setVerticalAlignment(VerticalAlignment.CENTER);
		styleSchueler.setBorderTop(BorderStyle.THIN);
		styleSchueler.setBorderBottom(BorderStyle.THIN);
		styleSchueler.setBorderLeft(BorderStyle.THIN);
		styleSchueler.setBorderRight(BorderStyle.THIN);

		XSSFCellStyle styleSchueler2 = workbook.createCellStyle();
		styleSchueler2.setAlignment(HorizontalAlignment.CENTER);
		styleSchueler2.setVerticalAlignment(VerticalAlignment.CENTER);
		styleSchueler2.setBorderTop(BorderStyle.THIN);
		styleSchueler2.setBorderBottom(BorderStyle.THIN);
		styleSchueler2.setBorderLeft(BorderStyle.THIN);
		styleSchueler2.setBorderRight(BorderStyle.THICK);

		for (int i = 0; i < schueler.size(); i++) {
			Row row = sheet.createRow(i + 4);
			Cell cell0 = row.createCell(0);
			cell0.setCellValue(schueler.get(i).getId());
			Cell cell1 = row.createCell(1);
			cell1.setCellStyle(styleSchueler);
			cell1.setCellValue(schueler.get(i).getName());
			Cell cell2 = row.createCell(2);
			cell2.setCellStyle(styleSchueler2);
			cell2.setCellValue(schueler.get(i).getVorname());

			row.setHeight((short) (row.getHeight() * 0.9f));
		}

		for (int i = 3; i < 3 + z1 + z2; i++) {
			sheet.setColumnWidth(i, (int) (sheet.getColumnWidth(i) * 1.7f));
		}

		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(2);
	}

	private static void auswahl(XSSFWorkbook workbook, XSSFSheet sheet, List<DBToExcelSchueler> schueler, List<DBToExcelDisziplin> disziplinen) {
		XSSFCellStyle borderstyle = workbook.createCellStyle();
		borderstyle.setBorderTop(BorderStyle.THIN);
		borderstyle.setBorderBottom(BorderStyle.THIN);
		borderstyle.setBorderLeft(BorderStyle.THIN);
		borderstyle.setBorderRight(BorderStyle.THIN);
		borderstyle.setAlignment(HorizontalAlignment.CENTER);
		borderstyle.setVerticalAlignment(VerticalAlignment.CENTER);
		borderstyle.setLocked(false);

		XSSFCellStyle borderstyleLeft = workbook.createCellStyle();
		borderstyleLeft.setBorderTop(BorderStyle.THIN);
		borderstyleLeft.setBorderBottom(BorderStyle.THIN);
		borderstyleLeft.setBorderLeft(BorderStyle.THICK);
		borderstyleLeft.setBorderRight(BorderStyle.THIN);
		borderstyleLeft.setAlignment(HorizontalAlignment.CENTER);
		borderstyleLeft.setVerticalAlignment(VerticalAlignment.CENTER);
		borderstyleLeft.setLocked(false);

		XSSFFont font = workbook.createFont();
		font.setFontHeight(sheet.getRow(4).getHeight());
		borderstyle.setFont(font);

		for (int i = 0; i < schueler.size(); i++) {
			for (int j = 0; j < disziplinen.size(); j++) {
				if (j + 3 == 3 + z1 || j + 3 == 3 + z1) {
					sheet.getRow(i + 4).createCell(j + 3).setCellStyle(borderstyleLeft);
				} else {
					sheet.getRow(i + 4).createCell(j + 3).setCellStyle(borderstyle);
				}
			}
		}
	}

	private static void bewertungSchueler(XSSFWorkbook workbook, XSSFSheet sheet, List<DBToExcelSchueler> schueler) {
		XSSFCellStyle borderstyle = workbook.createCellStyle();
		borderstyle.setBorderTop(BorderStyle.THIN);
		borderstyle.setBorderBottom(BorderStyle.THIN);
		borderstyle.setBorderLeft(BorderStyle.THICK);
		borderstyle.setBorderRight(BorderStyle.THICK);
		borderstyle.setVerticalAlignment(VerticalAlignment.CENTER);

		zA = 3 + z1 + z2;
		String formel2 = "IF(COUNTA(Dy:qy)<2,\"Es fehlen noch Einträge\"," + "IF(COUNTA(Dy:qy)<5,\"Noch Belegungen möglich\"," + "IF(COUNTA(Dy:qy)>5,\"Max 5. Einträge zulässig, bitte korrigieren\","
				+ "\"Super Beteiligung\")))";

		formel2 = formel2.replace("q", ""+(char) (65 + 2 + z1 + z2));
		
		for (int i = 0; i < schueler.size(); i++) {
			sheet.getRow(4 + i).createCell(zA).setCellFormula(formel2.replace("y", "" + (5 + i)));
			sheet.getRow(4 + i).getCell(zA).setCellStyle(borderstyle);
		}
		sheet.setColumnWidth(zA, 8000);
	}

	private static void bewertungDisziplinen(XSSFWorkbook workbook, XSSFSheet sheet, List<DBToExcelSchueler> schueler, List<DBToExcelDisziplin> disziplinen) {
		XSSFCellStyle borderstyle = workbook.createCellStyle();
		borderstyle.setBorderTop(BorderStyle.THICK);
		borderstyle.setBorderBottom(BorderStyle.THIN);
		borderstyle.setBorderLeft(BorderStyle.THIN);
		borderstyle.setBorderRight(BorderStyle.THICK);
		borderstyle.setVerticalAlignment(VerticalAlignment.CENTER);

		XSSFCellStyle borderstyle2 = workbook.createCellStyle();
		borderstyle2.setBorderTop(BorderStyle.THIN);
		borderstyle2.setBorderBottom(BorderStyle.THIN);
		borderstyle2.setBorderLeft(BorderStyle.THIN);
		borderstyle2.setBorderRight(BorderStyle.THICK);
		borderstyle2.setVerticalAlignment(VerticalAlignment.CENTER);

		XSSFCellStyle borderstyle3 = workbook.createCellStyle();
		borderstyle3.setBorderTop(BorderStyle.THICK);
		borderstyle3.setBorderBottom(BorderStyle.THIN);
		borderstyle3.setBorderLeft(BorderStyle.THIN);
		borderstyle3.setBorderRight(BorderStyle.THIN);
		borderstyle3.setVerticalAlignment(VerticalAlignment.CENTER);
		borderstyle3.setAlignment(HorizontalAlignment.CENTER);

		XSSFCellStyle borderstyle5 = workbook.createCellStyle();
		borderstyle5.setBorderTop(BorderStyle.THIN);
		borderstyle5.setBorderBottom(BorderStyle.THIN);
		borderstyle5.setBorderLeft(BorderStyle.THIN);
		borderstyle5.setBorderRight(BorderStyle.THIN);
		borderstyle5.setVerticalAlignment(VerticalAlignment.CENTER);
		borderstyle5.setAlignment(HorizontalAlignment.CENTER);

		XSSFCellStyle borderstyle3Left = workbook.createCellStyle();
		borderstyle3Left.setBorderTop(BorderStyle.THICK);
		borderstyle3Left.setBorderBottom(BorderStyle.THIN);
		borderstyle3Left.setBorderLeft(BorderStyle.THICK);
		borderstyle3Left.setBorderRight(BorderStyle.THIN);
		borderstyle3Left.setVerticalAlignment(VerticalAlignment.CENTER);
		borderstyle3Left.setAlignment(HorizontalAlignment.CENTER);

		XSSFCellStyle borderstyle5Left = workbook.createCellStyle();
		borderstyle5Left.setBorderTop(BorderStyle.THIN);
		borderstyle5Left.setBorderBottom(BorderStyle.THIN);
		borderstyle5Left.setBorderLeft(BorderStyle.THICK);
		borderstyle5Left.setBorderRight(BorderStyle.THIN);
		borderstyle5Left.setVerticalAlignment(VerticalAlignment.CENTER);
		borderstyle5Left.setAlignment(HorizontalAlignment.CENTER);

		XSSFCellStyle borderstyle6 = workbook.createCellStyle();
		borderstyle6.setBorderTop(BorderStyle.THICK);

		CellRangeAddress cra1 = new CellRangeAddress(4 + schueler.size(), 4 + schueler.size(), 1, 2);
		CellRangeAddress cra2 = new CellRangeAddress(5 + schueler.size(), 5 + schueler.size(), 1, 2);
		CellRangeAddress cra3 = new CellRangeAddress(6 + schueler.size(), 6 + schueler.size(), 1, 2);

		sheet.addMergedRegion(cra1);
		sheet.addMergedRegion(cra2);
		sheet.addMergedRegion(cra3);

		XSSFRow row1 = sheet.createRow(4 + schueler.size());
		XSSFRow row2 = sheet.createRow(5 + schueler.size());
		XSSFRow row3 = sheet.createRow(6 + schueler.size());
		XSSFRow row4 = sheet.createRow(7 + schueler.size());

		XSSFCell cell1 = row1.createCell(1);
		XSSFCell cell2 = row2.createCell(1);
		XSSFCell cell3 = row3.createCell(1);

		cell1.setCellValue("Maximale Teilnehmer");
		cell1.setCellStyle(borderstyle);
		row1.createCell(2).setCellStyle(borderstyle);

		cell2.setCellValue("Zur Zeit angemeldet");
		cell2.setCellStyle(borderstyle2);
		row2.createCell(2).setCellStyle(borderstyle2);

		cell3.setCellValue("Status");
		cell3.setCellStyle(borderstyle);
		row3.createCell(2).setCellStyle(borderstyle);

		row1.setHeight((short) (row1.getHeight() * 0.9f));
		row2.setHeight((short) (row2.getHeight() * 0.9f));
		row3.setHeight((short) (row3.getHeight() * 0.9f));

		String formula2 = "COUNTA(x5:x" + (4 + schueler.size()) + ")";
		String formula3Cell = "x" + (6 + schueler.size());
		String formula3 = "IF(" + formula3Cell + "<y, " + formula3Cell + "-y, IF(" + formula3Cell + ">z, " + formula3Cell + "-z, \"o.k\"))";

		for (int i = 3; i < 3 + disziplinen.size(); i++) {
			cell1 = row1.createCell(i);
			cell1.setCellValue(disziplinen.get(i - 3).getMax());

			cell2 = row2.createCell(i);
			cell2.setCellFormula(formula2.replace("x", "" + (char) (65 + i)));

			cell3 = row3.createCell(i);
			cell3.setCellFormula(formula3.replace("x", "" + (char) (65 + i)).replace("y", "" + disziplinen.get(i - 3).getMin()).replace("z", "" + disziplinen.get(i - 3).getMax()));

			if (i == 3 + z1 || i == 3 + z1) {
				cell1.setCellStyle(borderstyle3Left);
				cell2.setCellStyle(borderstyle5Left);
				cell3.setCellStyle(borderstyle3Left);
			} else {
				cell1.setCellStyle(borderstyle3);
				cell2.setCellStyle(borderstyle5);
				cell3.setCellStyle(borderstyle3);
			}

			row4.createCell(i).setCellStyle(borderstyle6);
		}

		row4.createCell(1).setCellStyle(borderstyle6);
		row4.createCell(2).setCellStyle(borderstyle6);

		XSSFSheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

		XSSFConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.EQUAL, "\"o.k\"");
		rule1.createFontFormatting().setFontColorIndex(IndexedColors.BLACK.getIndex());

		XSSFConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.NOT_EQUAL, "\"o.k\"");
		rule2.createFontFormatting().setFontColorIndex(IndexedColors.RED.getIndex());

		String row = "" + (7 + schueler.size());
		CellRangeAddress[] regions = { CellRangeAddress.valueOf("D" + row + ":" + (char) (65 + disziplinen.size() + 2) + row) };

		sheetCF.addConditionalFormatting(regions, rule2, rule1);
	}

	private static void extras(XSSFWorkbook workbook, XSSFSheet sheet, List<DBToExcelSchueler> schueler, List<DBToExcelDisziplin> disziplinen) {
		XSSFCellStyle around = workbook.createCellStyle();
		around.setBorderTop(BorderStyle.THICK);
		around.setBorderBottom(BorderStyle.THICK);
		around.setBorderLeft(BorderStyle.THICK);
		around.setBorderRight(BorderStyle.THICK);
		around.getFont().setBold(true);
		XSSFCell cell = sheet.getRow(2).createCell(3 + disziplinen.size());
		cell.setCellValue("Bemerkungen");
		cell.setCellStyle(around);

		XSSFCellStyle top_left = workbook.createCellStyle();
		top_left.setBorderTop(BorderStyle.THICK);
		top_left.setBorderLeft(BorderStyle.THICK);
		sheet.getRow(4 + schueler.size()).createCell(3 + disziplinen.size()).setCellStyle(top_left);

		XSSFCellStyle left = workbook.createCellStyle();
		left.setBorderLeft(BorderStyle.THICK);
		sheet.getRow(5 + schueler.size()).createCell(3 + disziplinen.size()).setCellStyle(left);
		sheet.getRow(6 + schueler.size()).createCell(3 + disziplinen.size()).setCellStyle(left);
		sheet.getRow(0).createCell(3 + disziplinen.size()).setCellStyle(left);
		sheet.getRow(1).createCell(3 + disziplinen.size()).setCellStyle(left);

		XSSFCellStyle bottom = workbook.createCellStyle();
		bottom.setBorderBottom(BorderStyle.THICK);
		bottom.setAlignment(HorizontalAlignment.CENTER);
		sheet.getRow(3).createCell(1).setCellStyle(bottom);
		sheet.getRow(2).createCell(1).setCellStyle(bottom);
		sheet.getRow(2).getCell(1).setCellValue("Name");

		XSSFCellStyle bottom_right = workbook.createCellStyle();
		bottom_right.setBorderBottom(BorderStyle.THICK);
		bottom_right.setBorderRight(BorderStyle.THICK);
		bottom_right.setAlignment(HorizontalAlignment.CENTER);
		sheet.getRow(3).createCell(2).setCellStyle(bottom_right);
		sheet.getRow(2).createCell(2).setCellStyle(bottom_right);
		sheet.getRow(2).getCell(2).setCellValue("Vorname");

		XSSFCellStyle right = workbook.createCellStyle();
		right.setBorderRight(BorderStyle.THICK);
		sheet.getRow(0).createCell(2).setCellStyle(right);
		sheet.getRow(1).createCell(2).setCellStyle(right);

		sheet.getRow(1).createCell(0).setCellValue(schueler.size());
		sheet.getRow(2).createCell(0).setCellValue(disziplinen.size());

	}

}
