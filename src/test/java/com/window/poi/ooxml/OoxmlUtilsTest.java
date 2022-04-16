package com.window.poi.ooxml;

import com.window.FileUtils;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;

public class OoxmlUtilsTest {

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

        OoxmlUtils.decrypt(filepath, password);
    }

    @Test
    public void docxEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-docx.docx");

        OoxmlUtils.encrypt(filepath, password);
    }

    // xlsx test

    @Test
    public void xlsxDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "xlsx.xlsx");

        OoxmlUtils.decrypt(filepath, password);
    }

    @Test
    public void xlsxEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-xlsx.xlsx");

        OoxmlUtils.encrypt(filepath, password);
    }

    // pptx test

    @Test
    public void pptxDecryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "pptx.pptx"); // aspose: s; poi: 1s

        OoxmlUtils.decrypt(filepath, password);
    }

    @Test
    public void pptxEncryptTest() {
        String filepath = FileUtils.getFilepath(prefix, "output-pptx.pptx"); // aspose 6s; poi 2s

        OoxmlUtils.encrypt(filepath, password);
    }
}
