package lusvincius;

import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.File;

import javax.imageio.ImageIO;

import controllers.CorController;
import controllers.XLSController;

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

	final String PATH_ORIGEM = "C:\\Users\\lucasvinicius\\Pictures\\Nova pasta\\";
	final String PATH_DESTINO = "C:\\Users\\lucasvinicius\\Documents\\";
	final String IMG = "a.png";

	BufferedImage img = carregarImagem(PATH_ORIGEM, IMG);
	Color[][] colors = corController.gerarMatrizCoresImg(img, img.getWidth(), img.getHeight());

	xlsController.gerarPlanilha(colors, img.getHeight(), img.getWidth(), PATH_DESTINO);
  }
}
