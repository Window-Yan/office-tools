package com.window;

import java.io.File;

public class FileUtils {

    /**
     * 根据输入路径，设定输出路径
     * @param path 源文件路径
     * @return 输出文件路径
     */
    public static String getOutputPath(String path) {
        int index = path.lastIndexOf(Const.SEPARATOR);
        if (index == -1) {
            return null;
        }
        return new StringBuilder(path.substring(0, index)).append(Const.SEPARATOR)
                .append("output-").append(path.substring(index + 1)).toString();
    }

    public static boolean checkIfExist(String filepath) {
        return (new File(filepath)).exists();
    }

    public static DocType getDocType (String filepath) {
        String suffix = filepath.substring(filepath.lastIndexOf(".") + 1);

        return DocType.valueOf(suffix.toUpperCase());
    }

    public static String getFilepath (String ...paths) {
        String basePath =  System.getProperty("user.dir");
        String path = String.join(File.separator, paths);

        return basePath + File.separator + path;
    }

    public static void main(String[] args) {
        System.out.println(getFilepath("files"));
    }
}
