package controllers;

import java.awt.Color;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

import util.Arquivos;

public class XLSController {

  final String NOME_PLANILHA = "lusvincius";
  final int XLS_CELL_HEIGHT = 10;
  final int XLS_CELL_WIDTH = XLS_CELL_HEIGHT / 5;

  public void gerarPlanilha(Color[][] colors, int qtdLinhas, int qtdColunas, String destino) {
	Workbook workbook = new XSSFWorkbook();
	Sheet folha = workbook.createSheet(NOME_PLANILHA);

	for (int row = 0; row < qtdLinhas; row++) {
	  Row linha = folha.createRow(row);
	  linha.setHeightInPoints(XLS_CELL_HEIGHT);

	  for (int col = 0; col < qtdColunas; col++) {
		XSSFCellStyle cellStyle = estiloCelula(workbook, colors[row][col]);
		folha.setColumnWidth(col, XLS_CELL_WIDTH * 256);
		linha.createCell(col).setCellStyle(cellStyle);
	  }
	  
	  System.out.println((row * qtdColunas) * 100 / (qtdLinhas * qtdColunas) + "%");
	}

	Arquivos.gerarXLS(workbook, destino, NOME_PLANILHA.concat(".xlsx"));
	System.out.println("==Planilha criada com sucesso==");
  }

  public XSSFCellStyle estiloCelula(Workbook workbook, Color cor) {
	XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
	XSSFColor borderColor = new XSSFColor(new java.awt.Color(212, 212, 212));
    XSSFColor backgroundColor = new XSSFColor(new java.awt.Color(cor.getRed(), cor.getGreen(), cor.getBlue()));

	cellStyle.setFillForegroundColor(backgroundColor);
	cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	
	cellStyle.setBorderBottom(BorderStyle.THIN);
	cellStyle.setBorderTop(BorderStyle.THIN);
	cellStyle.setBorderRight(BorderStyle.THIN);
	cellStyle.setBorderLeft(BorderStyle.THIN);
	
	cellStyle.setBorderColor(BorderSide.BOTTOM, borderColor);
	cellStyle.setBorderColor(BorderSide.TOP, borderColor);
	cellStyle.setBorderColor(BorderSide.LEFT, borderColor);
	cellStyle.setBorderColor(BorderSide.RIGHT, borderColor);
	
	return cellStyle;
  }
  
}
