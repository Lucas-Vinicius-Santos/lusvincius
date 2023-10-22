package util;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

import javax.imageio.ImageIO;

public class Arquivos {

    private static List<String> extensoesDisponieis = List.of(".png", ".jpg", ".jpeg");

    public static File gerarXLS(Workbook pastaDeTrabalho, String dir, String nomeArquivo) {
        try {
            File arquivo = new File(dir + File.separator + nomeArquivo);
            FileOutputStream fileOutputStream = new FileOutputStream(arquivo);
            pastaDeTrabalho.write(fileOutputStream);
            fileOutputStream.close();
            pastaDeTrabalho.close();
            return arquivo;
        } catch (IOException e) {
            throw new IllegalArgumentException(
                    "\nNão foi possível gerar o arquivo!".concat(e.getMessage()));
        }
    }

    public static BufferedImage carregarImagem(String path, String image) {
        return getBufferedImage(path, image, extensoesDisponieis.size() - 1);
    }

    private static BufferedImage getBufferedImage(String path, String image, int index) {
        if (index < 0) {
            throw new IllegalArgumentException(
                    "\nNão foi possível carregar imagem! Verifique as informações passadas e tente novamente."
                    .concat("\nExtensões disponíveis: ".concat(extensoesDisponieis.toString())));
        }

        try {
            return ImageIO.read(new File(path.concat(image).concat(extensoesDisponieis.get(index))));
        } catch (Exception e) {
            return getBufferedImage(path, image, index - 1);
        }
    }

    public static void abrirArquivo(File arquivo) {
        try {
            if (arquivo.exists()) {
                Desktop.getDesktop().open(arquivo);
            }
        } catch (IOException e) {
            throw new IllegalArgumentException(
                    "\nAlgo deu errado ao tentar abrir o arquivo em seu computador!"
                    .concat("\n").concat(e.getMessage()));
        }
    }
}
