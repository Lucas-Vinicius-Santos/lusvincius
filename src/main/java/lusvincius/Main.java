package lusvincius;

import controllers.CorController;
import controllers.XLSController;
import util.Arquivos;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.util.Scanner;

public class Main {

    private static final String PATH_ORIGEM = "input/";
    private static final String PATH_DESTINO = "output";

    private static final CorController corController = new CorController();
    private static final XLSController xlsController = new XLSController();

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
		String nomeImagem = getNomeImagem(scanner);
        Integer nivelQualidade = getNivelQualidade(scanner);
        Boolean planilhaComBordas = getPlanilhaComBordas(scanner);
        scanner.close();

        BufferedImage img = Arquivos.carregarImagem(PATH_ORIGEM, nomeImagem);
        Color[][] colors = corController.gerarMatrizCoresImg(img, img.getWidth(), img.getHeight(), nivelQualidade);
        File planilha = xlsController.gerarPlanilha(colors, img.getHeight(), img.getWidth(), PATH_DESTINO, nomeImagem, planilhaComBordas);
        Arquivos.abrirArquivo(planilha);
    }

    private static String getNomeImagem(Scanner scanner) {
        System.out.println("Digite o nome da imagem: ");
		System.out.flush();
        return scanner.nextLine();
    }

    private static Integer getNivelQualidade(Scanner scanner) {
        System.out.println("Informe o peso de fidelidade: ");
		System.out.flush();
        return scanner.nextInt();
    }

    private static Boolean getPlanilhaComBordas(Scanner scanner) {
        System.out.println("Deseja adicionar boras nas células? ");
		System.out.flush();
        return scanner.nextBoolean();
    }
}
