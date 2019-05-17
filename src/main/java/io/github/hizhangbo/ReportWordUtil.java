package com.zhilingsd.base.common.utils;

import com.zhilingsd.base.common.vo.ReportExportVo;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.springframework.http.ResponseEntity;
import org.springframework.util.CollectionUtils;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * 替换word模板内容
 * @Author: zhangbo
 * @DateTime: 2019/5/17 9:34
 */
@Slf4j
public class ReportWordUtil {
    /**
     * @description 导出ZIP文件
     **/
    public static ResponseEntity<byte[]> getWorldZipFile(String intputPath, List<ReportExportVo> list) throws IOException {
        //最大10M的world文件
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream(10 * 1024);
        ZipOutputStream zipOut = new ZipOutputStream(byteArrayOutputStream);
        try {
            if (!CollectionUtils.isEmpty(list)) {
                for (int i = 0; i < list.size(); i++) {
                    ReportExportVo vo = list.get(i);
                    //输出地址 输入地址 加随机数
                    InputStream is = new FileInputStream(intputPath);
                    XWPFDocument doc = new XWPFDocument(is);

                    replaceContent(doc, vo);

                    //把doc输出到输出流中
                    ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
                    doc.write(byteOutputStream);

                    zipOut.putNextEntry(new ZipEntry(System.currentTimeMillis() + ".docx"));
                    byte[] bytes = byteOutputStream.toByteArray();
                    zipOut.write(bytes);
                    zipOut.closeEntry();
                    byteOutputStream.close();
                }
                zipOut.close();
                String zipName = DateUtil.convertDateToString(DateUtil.DATE_TIME_PATTERN, new Date()) + ".zip";
                return SpringWebFileUtil.download(byteArrayOutputStream.toByteArray(), zipName);
            }
        } catch (XmlException e) {
            e.printStackTrace();
        } finally {
            if (zipOut != null) {
                zipOut.close();
            }
            if (byteArrayOutputStream != null) {
                byteArrayOutputStream.close();
            }
        }
        return null;
    }

    /**
     * 导出单个文件 world
     *
     * @param intputPath 输入地址
     * @throws Exception 导出单个文件
     */
    public static ResponseEntity<byte[]> getWorldFile(String intputPath, ReportExportVo vo) throws Exception {
        ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
        try {
            //输出地址 输入地址 加随机数
            InputStream is = new FileInputStream(intputPath);
            XWPFDocument docx = new XWPFDocument(is);
            replaceContent(docx, vo);
            //把doc输出到输出流中
            docx.write(byteOutputStream);
            byteOutputStream.close();
            String docName = DateUtil.convertDateToString(DateUtil.DATE_TIME_PATTERN, new Date()) + ".docx";
            return SpringWebFileUtil.download(byteOutputStream.toByteArray(), docName);
        } finally {
            if (null != byteOutputStream) {
                byteOutputStream.close();
            }
        }
    }

    private static String getWiteDate(String onDoorDate) {
        String result = "";
        try {
            Date date = DateUtil.addDate(DateUtil.convertStringToDate(DateUtil.DATE_MINUTE_CHINESE_YMD, onDoorDate), 3);
            result = DateUtil.convertDateToString(DateUtil.DATE_MINUTE_CHINESE_YMD, date);
        } catch (ParseException ex) {
            log.error(ex.getMessage(), ex);
        }
        return result;
    }


    private static void replaceContent(XWPFDocument doc, ReportExportVo vo) throws XmlException {
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            XmlCursor cursor = paragraph.getCTP().newCursor();
            cursor.selectPath("declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//*/w:txbxContent/w:p/w:r");

            List<XmlObject> ctrsintxtbx = new ArrayList<>();

            while (cursor.hasNextSelection()) {
                cursor.toNextSelection();
                XmlObject obj = cursor.getObject();
                ctrsintxtbx.add(obj);
            }
            for (XmlObject obj : ctrsintxtbx) {
                CTR ctr = CTR.Factory.parse(obj.xmlText());
                //CTR ctr = CTR.Factory.parse(obj.newInputStream());
                XWPFRun bufferrun = new XWPFRun(ctr, (IRunBody) paragraph);
                String text = bufferrun.getText(0);
                if (text != null) {
                    for (String word : vo.getExportValue().keySet()) {
                        if (text.contains(word)) {
                            text = text.replace(word, vo.getExportValue().get(word));
                            bufferrun.setText(text, 0);
                        }
                    }
                }
                obj.set(bufferrun.getCTR());
            }
        }
    }
}
