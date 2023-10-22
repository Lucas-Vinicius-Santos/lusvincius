package util;

import java.util.Date;

public class Logger {

    public static void progresso(int tamanhoTotal, int index, String label) {
        System.out.print(new Date().toString().concat(" === "));
        System.out.print(index * 100 / tamanhoTotal);
        System.out.println("% - ".concat(String.valueOf(index)).concat(" ").concat(label));
    }

    public static void texto(String label) {
        System.out.println(new Date().toString().concat(" === ").concat(label));
    }
}
