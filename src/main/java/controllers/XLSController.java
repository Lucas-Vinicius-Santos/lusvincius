package controllers;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import util.Arquivos;
import util.Logger;

import java.awt.*;
import java.io.File;
import java.util.List;
import java.util.*;

public class XLSController {

    private final String NOME_PLANILHA = "lusvincius";

    private final Workbook workbook;
    private final SXSSFSheet folha;
    private final XSSFCellStyle cellStyle;
    private final Map<Integer, XSSFCellStyle> estiloCelulaCache = new HashMap<>();

    public XLSController() {
        this.workbook = new SXSSFWorkbook(new XSSFWorkbook(), 1000, true);
        this.workbook.createSheet(NOME_PLANILHA);
        this.workbook.setSheetName(0, NOME_PLANILHA);
        this.folha = (SXSSFSheet) workbook.getSheetAt(0);
        this.folha.setRandomAccessWindowSize(100);
        this.cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        this.cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    public File gerarPlanilha(Color[][] colors, int qtdLinhas, int qtdColunas, String destino, String nomeImagem, boolean utilizarBordas) {
        List<Color> listaInvertidaCores;
        List<XSSFCellStyle> listaEstiloCelulas;

        if (utilizarBordas) this.estilizarBordas();

        listaInvertidaCores = converterMatrizEmListaInvertidaCores(qtdLinhas, qtdColunas, colors);
        listaEstiloCelulas = gerarListaEstilosCelulas(qtdLinhas, qtdColunas, listaInvertidaCores);

        montarPlanilha(qtdLinhas, qtdColunas, listaEstiloCelulas);

        File planilha = Arquivos.gerarXLS(workbook, destino, nomeImagem.concat(".xlsx"));
        Logger.texto("Planilha criada com sucesso");
        return planilha;
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

    private static List<Color> converterMatrizEmListaInvertidaCores(int qtdLinhas, int qtdColunas, Color[][] colors) {
        List<Color> listaCoresInvertida = new ArrayList<>();
        for (int row = qtdLinhas -1; row >= 0; row--) {
            for (int col = qtdColunas -1; col >= 0; col--) {
                listaCoresInvertida.add(colors[row][col]);
            }
        }
        return listaCoresInvertida;
    }

    private List<XSSFCellStyle> gerarListaEstilosCelulas(int qtdLinhas, int qtdColunas, List<Color> listaCoresInvertida) {
        List<XSSFCellStyle> estilosCelulas = new ArrayList<>();

        for (int index = 0; index < qtdLinhas * qtdColunas; index++) {
            Color cor = listaCoresInvertida.remove(listaCoresInvertida.size() -1);
            estilosCelulas.add(getEstilo(cor));

            if (index % 100 == 0) {
                Logger.progresso(qtdLinhas * qtdColunas, index, "Células estilizadas");
            }
        }
        Logger.progresso(qtdLinhas * qtdColunas, qtdLinhas * qtdColunas, "Células estilizadas");

        return estilosCelulas;
    }

    private XSSFCellStyle getEstilo(Color cor) {
        XSSFCellStyle novoEstilo;

        if (estiloCelulaCache.containsKey(cor.getRGB())) {
            novoEstilo = estiloCelulaCache.get(cor.getRGB());
        } else {
            novoEstilo = (XSSFCellStyle) this.cellStyle.clone();
            novoEstilo.setFillForegroundColor(new XSSFColor(new Color(cor.getRed(), cor.getGreen(), cor.getBlue())));
            estiloCelulaCache.put(cor.getRGB(), novoEstilo);
        }
        return novoEstilo;
    }

    private void montarPlanilha(int qtdLinhas, int qtdColunas, List<XSSFCellStyle> listaEstiloCelulas) {
        int XLS_CELL_HEIGHT = 10, XLS_CELL_WIDTH = XLS_CELL_HEIGHT / 5;

        Logger.texto("Montando planilha... ");
        for (int row = 0; row < qtdLinhas; row++) {
            Row linha = folha.createRow(row);
            linha.setHeightInPoints(XLS_CELL_HEIGHT);

            for (int col = 0; col < qtdColunas; col++) {
                folha.setColumnWidth(col, XLS_CELL_WIDTH * 256);
                linha.createCell(col).setCellStyle(listaEstiloCelulas.remove(0));
            }
        }
    }
}
