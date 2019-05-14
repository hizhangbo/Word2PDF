package io.github.hizhangbo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @Author: zhangbo
 * @DateTime: 2019/5/14 9:54
 */
public class Main {
    public static void main(String[] args) throws Exception {

        String inputFilePath = "./催收概要.docx";
//        String inputFilePath = "F:/催收平台/催收概要.docx";
        File inFile = new File(inputFilePath);
        InputStream inStream = new FileInputStream(inFile);

        String outputFilePath = "./催收概要.pdf";
//        String outputFilePath = "F:/催收概要.pdf";
        File outFile = new File(outputFilePath);

        try {
            //Make all directories up to specified
            outFile.getParentFile().mkdirs();
        } catch (NullPointerException e) {
            //Ignore error since it means not parent directories
        }

        outFile.createNewFile();
        OutputStream outStream = new FileOutputStream(outFile);

//        Converter converter = new DocToPDFConverter(inStream, outStream, true, true);
        Converter converter = new DocxToPDFConverter(inStream, outStream, true, true);
        converter.convert();
    }
}