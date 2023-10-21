package lusvincius;

import controllers.CorController;
import controllers.XLSController;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.sql.SQLOutput;
import java.util.Scanner;

public class Main {

  private static CorController corController = new CorController();
  private static XLSController xlsController = new XLSController();

  public static BufferedImage carregarImagem(String path, String image) {
	try {
	  return ImageIO.read(new File(path.concat(image)));
	} catch (Exception e) {
	  throw new IllegalArgumentException(
		  "Não foi possivel carregar imagem! Verifique as informações passadas e tente novamente.");
	}
  }

  public static void main(String[] args) {
	final String PATH_ORIGEM = "input/";
	final String PATH_DESTINO = "output";
	String IMG = getNomeImagem();

  	BufferedImage img = carregarImagem(PATH_ORIGEM, IMG);
	Color[][] colors = corController.gerarMatrizCoresImg(img, img.getWidth(), img.getHeight());

	xlsController.gerarPlanilha(colors, img.getHeight(), img.getWidth(), PATH_DESTINO, false);
  }

	private static String getNomeImagem() {
		String IMG;
		Scanner scanner = new Scanner(System.in);
		System.out.println("Digite o nome da imagem: ");
		System.out.flush();
		IMG = scanner.nextLine();
		scanner.close();
		return IMG;
	}
}
