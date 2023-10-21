package controllers;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import util.Arquivos;

import java.awt.Color;
import java.util.List;
import java.util.*;

public class XLSController {

    final String NOME_PLANILHA = "lusvincius";
    final int XLS_CELL_HEIGHT = 10;
    final int XLS_CELL_WIDTH = XLS_CELL_HEIGHT / 5;

    private final Workbook workbook;
    private final SXSSFSheet folha;
    private XSSFCellStyle cellStyle;
    private Map<Integer, XSSFCellStyle> estiloCelulaCache = new HashMap<>();

    public XLSController() {
        this.workbook = new SXSSFWorkbook(new XSSFWorkbook(), 1000, true);
        this.workbook.createSheet(NOME_PLANILHA);
        this.workbook.setSheetName(0, NOME_PLANILHA);
        this.folha = (SXSSFSheet) workbook.getSheetAt(0);
        this.folha.setRandomAccessWindowSize(100);
        this.cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        this.cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    public void gerarPlanilha(Color[][] colors, int qtdLinhas, int qtdColunas, String destino, boolean utilizarBordas) {
        if (utilizarBordas) {
            this.estilizarBordas();
        }

        List<XSSFCellStyle> listaEstiloCelulas = gerarListaEstiloCelulas(qtdLinhas, qtdColunas, colors);

        System.out.print(new Date()); System.out.println(" ==Montando planilha==");

        for (int row = 0; row < qtdLinhas; row++) {
            Row linha = folha.createRow(row);
            linha.setHeightInPoints(XLS_CELL_HEIGHT);
            folha.setColumnWidth(row, XLS_CELL_WIDTH * 256);

            for (int col = 0; col < qtdColunas; col++) {
                linha.createCell(col).setCellStyle(listaEstiloCelulas.remove(0));
            }
        }

        Arquivos.gerarXLS(workbook, destino, NOME_PLANILHA.concat(".xlsx"));
        System.out.print(new Date()); System.out.println(" ==Planilha criada com sucesso==");
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

    private static void logProgressoCriacaoTabela(int tamanhoTotal, int index, String label) {
        System.out.print(new Date().toString().concat(" === "));
        System.out.print(index * 100 / tamanhoTotal);
        System.out.println("% - ".concat(String.valueOf(index)).concat(" ").concat(label));
    }

    private List<XSSFCellStyle> gerarListaEstiloCelulas(int qtdLinhas, int qtdColunas, Color[][] colors) {
        List<XSSFCellStyle> listaEstiloCelulas = new ArrayList<>();
        List<Color> listaCoresInvertida = new ArrayList<>();
        for (int row = qtdLinhas-1; row >= 0; row--) {
            for (int col = qtdColunas-1; col >= 0; col--) {
                listaCoresInvertida.add(colors[row][col]);
            }
        }

        for (int index = 0; index < qtdLinhas * qtdColunas; index++) {
            Color cor = listaCoresInvertida.remove(listaCoresInvertida.size() -1);
            listaEstiloCelulas.add(getEstilo(cor));

            if (index % 10 == 0) {
                logProgressoCriacaoTabela(qtdLinhas * qtdColunas, index, "Células estilizadas");
            }
        }
        logProgressoCriacaoTabela(qtdLinhas * qtdColunas, qtdLinhas * qtdColunas, "Células estilizadas");

        return listaEstiloCelulas;
    }

    private XSSFCellStyle getEstilo(Color cor) {
        XSSFCellStyle novoEstilo;
        int arredondar = 8;
        cor = new Color(Math.round(cor.getRed() / arredondar)*arredondar,
                Math.round(cor.getGreen() / arredondar)*arredondar,
                Math.round(cor.getBlue() / arredondar)*arredondar);


        if (estiloCelulaCache.containsKey(cor.getRGB())) {
            novoEstilo = estiloCelulaCache.get(cor.getRGB());
        } else {
            novoEstilo = (XSSFCellStyle) this.cellStyle.clone();
            novoEstilo.setFillForegroundColor(new XSSFColor(new Color(cor.getRed(), cor.getGreen(), cor.getBlue())));
            estiloCelulaCache.put(cor.getRGB(), novoEstilo);
        }
        return novoEstilo;
    }
}
