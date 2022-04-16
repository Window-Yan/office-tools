package com.window.poi.bin;

import com.window.DocType;
import com.window.FileUtils;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * @author yanwt
 */
public class BinaryUtils {

    public static boolean save(DocType docType, POIFSFileSystem fs, OutputStream os, String password) {
        try {
            switch (docType) {
                case DOC:
                    HWPFDocument document = new HWPFDocument(fs);
                    Biff8EncryptionKey.setCurrentUserPassword(password);
                    document.write(os);
                    break;
                case XLS:
                    HSSFWorkbook workbook = new HSSFWorkbook(fs.getRoot(), true);
                    Biff8EncryptionKey.setCurrentUserPassword(password);
                    workbook.write(os);
                    break;
                case PPT:
                    HSLFSlideShow slideShow = new HSLFSlideShow(fs);
                    Biff8EncryptionKey.setCurrentUserPassword(password);
                    slideShow.write(os);
                    break;
                default:
                    return false;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return true;
    }

    public static String decrypt(String filepath, String password) {

        DocType docType = FileUtils.getDocType(filepath);

        String output = FileUtils.getOutputPath(filepath);
        if (output == null) {
            System.out.println("无效文件");
            return null;
        }

        Biff8EncryptionKey.setCurrentUserPassword(password);
        try (POIFSFileSystem fs = new POIFSFileSystem(new File(filepath), true);
             OutputStream os = new FileOutputStream(output)) {
            save(docType, fs, os, null);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return output;
    }

    public static String encrypt(String filepath, String password) {

        DocType docType = FileUtils.getDocType(filepath);

        String output = FileUtils.getOutputPath(filepath);
        if (output == null) {
            System.out.println("无效文件");
            return null;
        }

        try (POIFSFileSystem fs = new POIFSFileSystem(new File(filepath), true);
             OutputStream os = new FileOutputStream(output)) {
            save(docType, fs, os, password);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return output;
    }
}
