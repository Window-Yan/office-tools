package com.window.poi.ooxml;

import com.window.FileUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.security.GeneralSecurityException;
import java.util.Arrays;

public class OoxmlUtils {

    /**
     * 解密ooxml(docx,xlsx,pptx)格式文档
     *
     * @param filepath 源文件路径
     * @param password 加密密码
     * @return 解密后的文件路径
     */
    public static String decrypt(String filepath, String password) {
        String output = FileUtils.getOutputPath(filepath);
        if (output == null) {
            System.out.println("空文件");
            return null;
        }
        String[] supportFormats = {"docx", "pptx", "xlsx"};
        String suffix = filepath.substring(filepath.lastIndexOf(".") + 1);
        Arrays.sort(supportFormats);
        System.out.println("支持的文件类型：" + Arrays.toString(supportFormats));
        System.out.println("文件后缀：" + suffix);
        if (Arrays.binarySearch(supportFormats, suffix) < 0) {
            System.out.println("不支持的解密文件类型");
            return null;
        }
        System.out.println("输出文件:" + output);
        File outputFile = new File(output);
        if (outputFile.exists()) {
            outputFile.delete();
        }
        try (POIFSFileSystem fs = new POIFSFileSystem(new File(filepath), true);
             BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(outputFile))) {
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor decryptor = Decryptor.getInstance(info);

            if (!decryptor.verifyPassword(password)) {
                System.out.println("密码错误");
            }

            InputStream dataStream = decryptor.getDataStream(fs);

            byte[] buf = new byte[1024];
            int len;
            while ((len = dataStream.read(buf)) != -1) {
                bos.write(buf, 0, len);
                bos.flush();
            }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (GeneralSecurityException e) {
            e.printStackTrace();
        }
        return output;
    }

    public static String encrypt(String filepath, String password) {
        long start = System.currentTimeMillis();
        if (! FileUtils.checkIfExist(filepath)) {
            System.out.println("无效文件");
            return null;
        }
        String output = FileUtils.getOutputPath(filepath);
        if (output == null) {
            System.out.println("无效文件");
            return null;
        }
        System.out.println("文件大小：" + ((double)(new File(filepath)).length()) / (1024 * 1024) + "M" );
        try (POIFSFileSystem fs = new POIFSFileSystem()){
//            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
            // 支持设置更多加密参数

            Encryptor encryptor = info.getEncryptor();
            encryptor.confirmPassword(password);

            // Read in an existing OOXML file and write to encrypted output stream
            // don't forget to close the output stream otherwise the padding bytes aren't added
            try (OPCPackage opc = OPCPackage.open(new File(filepath), PackageAccess.READ_WRITE)) {
                ZipSecureFile.setMinInflateRatio(0.001);
                OutputStream dataStream = encryptor.getDataStream(fs);
                opc.save(dataStream);
                dataStream.close();
            } catch (GeneralSecurityException e) {
                e.printStackTrace();
            }

            // Write out the encrypted version
            try (FileOutputStream fos = new FileOutputStream(output)) {
                fs.writeFilesystem(fos);
            }
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
        long end = System.currentTimeMillis();
        System.out.println("耗时：" + (end - start) + "ms");

        return output;
    }

    public static void main(String[] args) {
        String password = "123456";
        String prefix = "files" + File.separator + "encrypt_files";
        String filepath = FileUtils.getFilepath(prefix, "docx.docx");

//        decrypt(filepath, password);
        encrypt(filepath, password);
    }
}
