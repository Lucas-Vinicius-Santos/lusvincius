package controllers;

import java.awt.*;
import java.awt.image.BufferedImage;

public class CorController {

  public Color[][] gerarMatrizCoresImg(BufferedImage imagem, int width, int height) {
	Color[][] result = new Color[height][width];

	for (int row = 0; row < height; row++) {
	  for (int col = 0; col < width; col++) {
		result[row][col] = new Color(imagem.getRGB(col, row), true);
	  }
	}

	return result;
  }
}
