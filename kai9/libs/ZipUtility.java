package kai9.libs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveOutputStream;
import org.apache.commons.compress.utils.IOUtils;

//フォルダ圧縮
public class ZipUtility {

    // 指定されたフォルダ内のファイルとサブフォルダを再帰的に圧縮するメソッド
    public static void zipFolder(String sourceFolderPath, String outputZipPath) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(outputZipPath);
                ZipArchiveOutputStream zos = new ZipArchiveOutputStream(fos)) {
            File sourceFile = new File(sourceFolderPath);
            // 圧縮処理を開始
            addFolderToZip(sourceFile, zos);
        }
    }

    // フォルダ内のファイルとサブフォルダを圧縮するメソッド
    private static void addFolderToZip(File folder, ZipArchiveOutputStream zos) throws IOException {
        for (File file : folder.listFiles()) {
            if (file.isDirectory()) {
                // サブフォルダ内のファイルを追加
                addFilesFromFolderToZip(file, file.getName(), zos);
            } else {
                // ファイルを追加
                try (FileInputStream fis = new FileInputStream(file)) {
                    ZipArchiveEntry entry = new ZipArchiveEntry(file.getName());
                    zos.putArchiveEntry(entry);
                    IOUtils.copy(fis, zos);
                    zos.closeArchiveEntry();
                }
            }
        }
    }

    // サブフォルダ内のファイルを圧縮するメソッド
    private static void addFilesFromFolderToZip(File folder, String folderName, ZipArchiveOutputStream zos) throws IOException {
        for (File file : folder.listFiles()) {
            if (file.isDirectory()) {
                // 再帰的にサブフォルダ内のファイルを追加
                addFilesFromFolderToZip(file, folderName + "/" + file.getName(), zos);
            } else {
                // ファイルを追加
                try (FileInputStream fis = new FileInputStream(file)) {
                    ZipArchiveEntry entry = new ZipArchiveEntry(folderName + "/" + file.getName());
                    zos.putArchiveEntry(entry);
                    IOUtils.copy(fis, zos);
                    zos.closeArchiveEntry();
                }
            }
        }
    }
}