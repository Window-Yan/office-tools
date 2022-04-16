package com.window.poi.bin;

import com.window.FileUtils;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;

public class BinaryUtilsTest {

    long start = -1;

    long end = 0;

    String prefix = "files" + File.separator + "encrypt_files";

    String password = "123456";

    @BeforeEach
    public void setStart() {
        start = System.currentTimeMillis();
    }

    @AfterEach
    public void setEnd() {
        end = System.currentTimeMillis();
        System.out.println("转换耗时：" + (end - start));
    }

    // doc test

    @Test
    public void docDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "doc.doc");

        BinaryUtils.decrypt(filepath, password);
    }

    @Test
    public void docEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-doc.doc");

        BinaryUtils.encrypt(filepath, password);
    }

    // xls test

    @Test
    public void xlsDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "xls.xls");

        BinaryUtils.decrypt(filepath, password);
    }

    @Test
    public void xlsEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-xls.xls");

        BinaryUtils.encrypt(filepath, password);
    }

    // ppt test

    @Test
    public void pptDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "ppt.ppt");

        BinaryUtils.decrypt(filepath, password);
    }

    @Test
    public void pptEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-ppt.ppt");

        BinaryUtils.encrypt(filepath, password);
    }
}
