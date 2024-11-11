package kai9.libs;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class Kai9Util {

    /**
     * エラーメッセージを表示
     */
    public static void ErrorMsg(String msg) {
        throw new RuntimeException(msg);
    }

    /**
     * ディレクトリを再帰的に作成する
     */
    public static void MakeDirs(String path) {
        Path p = Paths.get(path);

        try {
            Files.createDirectories(p);
        } catch (IOException e) {
            System.out.println(e);
        }
    }

}
