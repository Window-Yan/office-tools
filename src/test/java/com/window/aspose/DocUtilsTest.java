package com.window.aspose;

import com.window.FileUtils;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;

public class DocUtilsTest {

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

    // docx test

    @Test
    public void docxDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "docx.docx");

        DocUtils.decrypt(filepath, password);
    }

    @Test
    public void docxEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-docx.docx");

        DocUtils.encrypt(filepath, password);
    }

    // doc test

    @Test
    public void docDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "doc.doc");

        DocUtils.decrypt(filepath, password);
    }

    @Test
    public void docEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-doc.doc");

        DocUtils.encrypt(filepath, password);
    }

    // xlsx test

    @Test
    public void xlsxDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "xlsx.xlsx");

        DocUtils.decrypt(filepath, password);
    }

    @Test
    public void xlsxEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-xlsx.xlsx");

        DocUtils.encrypt(filepath, password);
    }

    // xls test

    @Test
    public void xlsDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "xls.xls");

        DocUtils.decrypt(filepath, password);
    }

    @Test
    public void xlsEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-xls.xls");

        DocUtils.encrypt(filepath, password);
    }

    // pptx test

    @Test
    public void pptxDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "pptx.pptx");

        DocUtils.decrypt(filepath, password);
    }

    @Test
    public void pptxEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-pptx.pptx");

        DocUtils.encrypt(filepath, password);
    }

    // ppt test

    @Test
    public void pptDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "ppt.ppt");

        DocUtils.decrypt(filepath, password);
    }

    @Test
    public void pptEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-ppt.ppt");

        DocUtils.encrypt(filepath, password);
    }
}
