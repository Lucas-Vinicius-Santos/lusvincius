package util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;

public class Arquivos {

  
  public static File gerarXLS(Workbook pastaDeTrabalho, String dir, String nomeArquivo) {
	try {
	  File arquivo = new File(dir + File.separator + nomeArquivo);
	  FileOutputStream fileOutputStream = new FileOutputStream(arquivo);
	  pastaDeTrabalho.write(fileOutputStream);
	  fileOutputStream.close();
	  pastaDeTrabalho.close();
	  return arquivo;
	} catch (IOException e) {
	  throw new IllegalArgumentException("Não foi possivel gerar o arquivo!: " + e.getMessage());
	}
  }
}
