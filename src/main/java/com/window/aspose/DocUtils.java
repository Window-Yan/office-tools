package com.window.aspose;

import com.aspose.cells.Workbook;
import com.aspose.slides.Presentation;
import com.aspose.words.*;
import com.window.DocType;
import com.window.FileUtils;

/**
 * @author Lenovocloud
 */
public class DocUtils {

    public static String encrypt(String filepath, String password) {
        DocType docType = FileUtils.getDocType(filepath);
        String output = FileUtils.getOutputPath(filepath);
        if (output == null) {
            System.out.println("无效文件");
            return null;
        }

        SaveOptions wordSaveOptions = null;
        int slidesSaveFormat = -1;
        switch (docType) {
            case DOC:
                wordSaveOptions = new DocSaveOptions(SaveFormat.DOC);
                ((DocSaveOptions)wordSaveOptions).setPassword(password);
                break;
            case DOCX:
                wordSaveOptions = new OoxmlSaveOptions(SaveFormat.DOCX);
                ((OoxmlSaveOptions)wordSaveOptions).setPassword(password);
                break;
            case XLS:
            case XLSX:
                try {
                    Workbook workbook = new Workbook(filepath);
                    workbook.getSettings().setPassword(password);
                    workbook.save(output);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                break;
            case PPT:
                slidesSaveFormat = com.aspose.slides.SaveFormat.Ppt;
                break;
            case PPTX:
                slidesSaveFormat = com.aspose.slides.SaveFormat.Pptx;
                break;
            default:
                System.out.println("不匹配的类型");
        }

        if (docType == DocType.DOC || docType == DocType.DOCX) {
            try {
                Document document = new Document(filepath);
                document.save(output, wordSaveOptions);
            } catch (Exception e) {
                e.printStackTrace();
            }

        }

        if (docType == DocType.PPT || docType == DocType.PPTX) {
            Presentation presentation = new Presentation(filepath);
            presentation.getProtectionManager().encrypt(password);
            presentation.save(output, slidesSaveFormat);
            presentation.dispose();
        }

        return output;
    }

    public static String decrypt(String filepath, String password) {
        DocType docType = FileUtils.getDocType(filepath);
        String output = FileUtils.getOutputPath(filepath);
        if (output == null) {
            System.out.println("无效文件");
            return null;
        }

        int slidesSaveFormat = -1;
        switch (docType) {
            case DOC:
            case DOCX:
                LoadOptions loadOptions = new LoadOptions(password);
                try {
                    Document document = new Document(filepath, loadOptions);
                    document.save(output);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                break;
            case XLS:
            case XLSX:
                com.aspose.cells.LoadOptions loadOptions1 = new com.aspose.cells.LoadOptions();
                loadOptions1.setPassword(password);
                try {
                    Workbook workbook = new Workbook(filepath, loadOptions1);
                    workbook.unprotect(password);
                    workbook.getSettings().setPassword(null);
                    workbook.save(output);
                } catch (Exception e) {
                    e.printStackTrace();
                }
                break;
            case PPT:
                slidesSaveFormat = com.aspose.slides.SaveFormat.Ppt;
                break;
            case PPTX:
                slidesSaveFormat = com.aspose.slides.SaveFormat.Pptx;
                break;
            default:
                System.out.println("不匹配的类型");
        }

        if (docType == DocType.PPT || docType == DocType.PPTX) {
            com.aspose.slides.LoadOptions loadOptions2 = new com.aspose.slides.LoadOptions();
            loadOptions2.setPassword(password);
            Presentation presentation = new Presentation(filepath, loadOptions2);
            presentation.getProtectionManager().removeEncryption();
            presentation.save(output, slidesSaveFormat);
            presentation.dispose();
        }

        return output;
    }

    public static void main(String[] args) {

    }
}
