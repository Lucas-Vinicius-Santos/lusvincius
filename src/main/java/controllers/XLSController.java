package controllers;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import util.Arquivos;

import java.awt.Color;

public class XLSController {

  final String NOME_PLANILHA = "lusvincius";
  final int XLS_CELL_HEIGHT = 10;
  final int XLS_CELL_WIDTH = XLS_CELL_HEIGHT / 5;

  private final Workbook workbook;
  private final Sheet folha;
  private final XSSFCellStyle cellStyle;

  public XLSController() {
	  this.workbook = new XSSFWorkbook();
	  this.folha =  workbook.createSheet(NOME_PLANILHA);
	  this.cellStyle = (XSSFCellStyle) workbook.createCellStyle();
  }

  public void gerarPlanilha(Color[][] colors, int qtdLinhas, int qtdColunas, String destino, boolean utilizarBordas) {
    if (utilizarBordas) {
	  this.estilizarBordas();
    }

	for (int row = 0; row < qtdLinhas; row++) {
	  Row linha = folha.createRow(row);
	  linha.setHeightInPoints(XLS_CELL_HEIGHT);

	  for (int col = 0; col < qtdColunas; col++) {
		folha.setColumnWidth(col, XLS_CELL_WIDTH * 256);
		linha.createCell(col).setCellStyle(estiloCelula(colors[row][col]));
	  }

	  logProgressoCriacaoTabela(qtdLinhas, qtdColunas, row);
	}
    System.out.println("100% - ".concat(String.valueOf(qtdLinhas)));

	Arquivos.gerarXLS(workbook, destino, NOME_PLANILHA.concat(".xlsx"));
	System.out.println("==Planilha criada com sucesso==");
  }

	public void estilizarBordas() {
	XSSFColor borderColor = new XSSFColor(new java.awt.Color(212, 212, 212));

	this.cellStyle.setBorderBottom(BorderStyle.THIN);
	this.cellStyle.setBorderTop(BorderStyle.THIN);
	this.cellStyle.setBorderRight(BorderStyle.THIN);
	this.cellStyle.setBorderLeft(BorderStyle.THIN);

	this.cellStyle.setBorderColor(BorderSide.BOTTOM, borderColor);
	this.cellStyle.setBorderColor(BorderSide.TOP, borderColor);
	this.cellStyle.setBorderColor(BorderSide.LEFT, borderColor);
	this.cellStyle.setBorderColor(BorderSide.RIGHT, borderColor);
  }

  public XSSFCellStyle estiloCelula(Color cor) {
    XSSFColor backgroundColor = new XSSFColor(new java.awt.Color(cor.getRed(), cor.getGreen(), cor.getBlue()));

	this.cellStyle.setFillForegroundColor(backgroundColor);
	this.cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	return (XSSFCellStyle) this.cellStyle.clone();
  }

  private static void logProgressoCriacaoTabela(int qtdLinhas, int qtdColunas, int row) {
	System.out.print((row * qtdColunas) * 100 / (qtdLinhas * qtdColunas));
	System.out.println("% - ".concat(String.valueOf(row)));
  }
}
