package controllers;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import util.Arquivos;

import java.awt.Color;
import java.util.*;
import java.util.concurrent.*;

public class XLSController {

    final String NOME_PLANILHA = "lusvincius";
    final int XLS_CELL_HEIGHT = 10;
    final int XLS_CELL_WIDTH = XLS_CELL_HEIGHT / 5;

    private final Workbook workbook;
    private final Sheet folha;
    private XSSFCellStyle cellStyle;
    private Map<Color, XSSFCellStyle> estiloCelulaCache = new HashMap<>();

    public XLSController() {
        this.workbook = new XSSFWorkbook();
        this.folha = workbook.createSheet(NOME_PLANILHA);
        this.cellStyle = (XSSFCellStyle) workbook.createCellStyle();
    }

    public void gerarPlanilha(Color[][] colors, int qtdLinhas, int qtdColunas, String destino, boolean utilizarBordas) {
        if (utilizarBordas) {
            this.estilizarBordas();
        }

        XSSFCellStyle[][] matrizEstiloCelulas = gerarMatrizEstiloCelulas(qtdLinhas, qtdColunas, colors);
        System.out.println(new Date().toString().concat(" - Estilos criados"));

        for (int row = 0; row < qtdLinhas; row++) {
            Row linha = folha.createRow(row);
            linha.setHeightInPoints(XLS_CELL_HEIGHT);
            folha.setColumnWidth(row, XLS_CELL_WIDTH * 256);

            for (int col = 0; col < qtdColunas; col++) {
                linha.createCell(col).setCellStyle(matrizEstiloCelulas[row][col]);
            }

            logProgressoCriacaoTabela(qtdLinhas, qtdColunas, row, "Celulas");
        }
        logProgressoCriacaoTabela(qtdLinhas, qtdColunas, qtdLinhas, "Celulas");

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
        if (estiloCelulaCache.containsKey(cor)) {
            return estiloCelulaCache.get(cor);
        } else {
            XSSFCellStyle novoEstilo = (XSSFCellStyle) this.cellStyle.clone();
            novoEstilo.setFillForegroundColor(new XSSFColor(new java.awt.Color(cor.getRed(), cor.getGreen(), cor.getBlue())));
            novoEstilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            estiloCelulaCache.put(cor, novoEstilo);
            return novoEstilo;
        }
    }

    private static void logProgressoCriacaoTabela(int qtdLinhas, int qtdColunas, int row, String label) {
        System.out.print(new Date().toString().concat(" === "));
        System.out.print((row * qtdColunas) * 100 / (qtdLinhas * qtdColunas));
        System.out.println("% - ".concat(String.valueOf(row)).concat(" ").concat(label));
    }

    private XSSFCellStyle[][] gerarMatrizEstiloCelulas(int qtdLinhas, int qtdColunas, Color[][] colors) {
        XSSFCellStyle[][] matrizEstilo = new XSSFCellStyle[qtdLinhas][qtdColunas];

        Map<Pair<Integer, Integer>, Color> mapaCores = new HashMap<>();
        for (int row = 0; row < qtdLinhas; row++) {
            for (int col = 0; col < qtdColunas; col++) {
                mapaCores.put(new Pair<>(row, col), colors[row][col]);
            }
            logProgressoCriacaoTabela(qtdLinhas, qtdColunas, row, "Mapa");
        }


        for (int row = 0; row < qtdLinhas; row++) {
            for (int col = 0; col < qtdColunas; col++) {
                matrizEstilo[row][col] = estiloCelula(mapaCores.get(new Pair<>(row, col)));
            }
            logProgressoCriacaoTabela(qtdLinhas, qtdColunas, row, "Estilo");
        }

        return matrizEstilo;
    }

    private XSSFCellStyle[][] gerarMatrizEstiloCelulasThreads(int qtdLinhas, int qtdColunas, Color[][] colors) {
        XSSFCellStyle[][] matrizEstilo = new XSSFCellStyle[qtdLinhas][qtdColunas];
        ExecutorService executorService = null;

        try {
            System.out.println(new Date().toString().concat(" - MultiThreads"));
            executorService = Executors.newFixedThreadPool(2);
            executorService.execute(() -> {
                for (int row = 0; row < qtdLinhas / 2; row++) {
                    for (int col = 0; col < qtdColunas; col++) {
                        matrizEstilo[row][col] = estiloCelula(colors[row][col]);
                    }
                    logProgressoCriacaoTabela(qtdLinhas, qtdColunas, row, "Estilo (1)");
                }
            });

            executorService.execute(() -> {
                for (int row = qtdLinhas / 2; row < qtdLinhas; row++) {
                    for (int col = 0; col < qtdColunas; col++) {
                        matrizEstilo[row][col] = estiloCelula(colors[row][col]);
                    }
                    logProgressoCriacaoTabela(qtdLinhas, qtdColunas, row, "Estilo (2)");
                }
            });

            executorService.shutdown();
            executorService.awaitTermination(5, TimeUnit.MINUTES);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (executorService != null) {
                executorService.shutdown();
            }
        }

        return matrizEstilo;
    }
}
