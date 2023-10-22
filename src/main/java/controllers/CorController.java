package controllers;

import java.awt.*;
import java.awt.image.BufferedImage;

public class CorController {

    public Color[][] gerarMatrizCoresImg(BufferedImage imagem, int width, int height, float quanlidade) {
        Color[][] matrizCores = new Color[height][width];

        for (int row = 0; row < height; row++) {
            for (int col = 0; col < width; col++) {
                Color color = new Color(imagem.getRGB(col, row), true);
                matrizCores[row][col] = getCorArredondandoRGB(color, quanlidade);
            }
        }

        return matrizCores;
    }

    private Color getCorArredondandoRGB(Color cor, float arredondar) {
        return new Color(
                Math.min((int) (Math.ceil((float) cor.getRed()   / arredondar) * arredondar), 255),
                Math.min((int) (Math.ceil((float) cor.getGreen() / arredondar) * arredondar), 255),
                Math.min((int) (Math.ceil((float) cor.getBlue()  / arredondar) * arredondar), 255)
        );
    }
}
