package com.window.aspose;

import com.aspose.pdf.Document;
import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;
import com.window.Const;
import com.window.DocType;
import com.window.FileUtils;

import java.io.File;
import java.util.Arrays;

/**
 * @author window
 */
public class PdfUtils {

    /**
     * 拆分pdf
     * @param filepath 源文件路径
     * @param splitUnit 拆分后单个pdf的页数
     * @return 拆分后目录
     */
    public static String splitPdf(String filepath, int splitUnit) {
        int dotIndex = filepath.lastIndexOf(".");
        String suffix = filepath.substring(dotIndex);
        String pdf = "pdf";
        if (! pdf.equalsIgnoreCase(suffix)) {
            System.out.println("无效的文件格式");
            return null;
        }
        String outputPath = filepath.substring(0, dotIndex) + Const.SEPARATOR + "output-" + filepath.substring(dotIndex + 1);
        File outputDir = new File(outputPath);
        if (! outputDir.exists()) {
            outputDir.mkdirs();
        }
        long t1 = System.currentTimeMillis();
        Presentation pres = new Presentation(filepath);
        int total = pres.getSlides().size();
        System.out.println("幻灯片页数：" + total);
        for (int i = 1, j = 1; i <= total; i += (splitUnit), j++) {
            int end = Math.min((i + splitUnit - 1), total);
            int[] slides = new int[end - i + 1];
            for (int k = 0; k <= (end - i); k++) {
                slides[k] = i + k;
            }
            System.out.println(Arrays.toString(slides));
            pres.save(outputPath + Const.SEPARATOR + "output_" + j + ".pdf", slides, SaveFormat.Pdf);
            System.out.println("第" + j + "部分转换结束");
        }
        long t2 = System.currentTimeMillis();
        System.out.println("转换耗时：" + (t2 - t1));

        return outputPath;
    }

    /**
     * 线性化PDF(网络预览优化)
     * @param filepath 源文件路径
     * @return 线性化后文件路径
     */
    public static String linearizedPdf(String filepath) {
        Document document = new Document(filepath);
        boolean linearized = document.isLinearized();
        System.out.println("是否线性化:" + linearized);
        if (linearized) {
            return filepath;
        } else {
            document.optimize();
            System.out.println("是否线性化:" + document.isLinearized());
            String output = FileUtils.getOutputPath(filepath);
            if (output != null) {
                document.save(output);
            }
            return output;
        }
    }

    public static void main(String[] args) {

    }
}
