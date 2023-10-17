package lusvincius;

import controllers.CorController;
import controllers.XLSController;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.lang.management.ManagementFactory;
import java.lang.management.RuntimeMXBean;

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
	final String PATH_ORIGEM = "C:\\Users\\55859\\Pictures\\lusvincius\\";
	final String PATH_DESTINO = "C:\\Users\\55859\\Documents\\";
	final String IMG = "testelusvincius212.jpg";
//	final String IMG = "testelusvincius106.jpg";

	BufferedImage img = carregarImagem(PATH_ORIGEM, IMG);
	Color[][] colors = corController.gerarMatrizCoresImg(img, img.getWidth(), img.getHeight());

	xlsController.gerarPlanilha(colors, img.getHeight(), img.getWidth(), PATH_DESTINO, false);
  }
}
