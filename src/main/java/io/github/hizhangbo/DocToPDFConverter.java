package io.github.hizhangbo;

import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;

import org.docx4j.Docx4J;
import org.docx4j.convert.in.Doc;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 * @Author: zhangbo
 * @DateTime: 2019/5/14 9:47
 */
public class DocToPDFConverter extends Converter {
    public DocToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    @Override
    public void convert() throws Exception {

        loading();

        InputStream iStream = inStream;
        WordprocessingMLPackage wordMLPackage = getMLPackage(iStream);

//        //加载字体文件（解决linux环境下无中文字体问题）
//        PhysicalFonts.addPhysicalFonts("SimSun", new URL("simsun.ttc"));
//        //宋体&新宋体
//        PhysicalFont simsunFont = PhysicalFonts.get("SimSun");
//        fontMapper.put("SimSun", simsunFont);


        Mapper fontMapper = new IdentityPlusMapper();
        fontMapper.put("隶书", PhysicalFonts.get("LiSu"));
        fontMapper.put("宋体", PhysicalFonts.get("SimSun"));
        fontMapper.put("微软雅黑", PhysicalFonts.get("Microsoft Yahei"));
        fontMapper.put("黑体", PhysicalFonts.get("SimHei"));
        fontMapper.put("楷体", PhysicalFonts.get("KaiTi"));
        fontMapper.put("新宋体", PhysicalFonts.get("NSimSun"));
        fontMapper.put("华文行楷", PhysicalFonts.get("STXingkai"));
        fontMapper.put("华文仿宋", PhysicalFonts.get("STFangsong"));
        fontMapper.put("宋体扩展", PhysicalFonts.get("simsun-extB"));
        fontMapper.put("仿宋", PhysicalFonts.get("FangSong"));
        fontMapper.put("仿宋_GB2312", PhysicalFonts.get("FangSong_GB2312"));
        fontMapper.put("幼圆", PhysicalFonts.get("YouYuan"));
        fontMapper.put("华文宋体", PhysicalFonts.get("STSong"));
        fontMapper.put("华文中宋", PhysicalFonts.get("STZhongsong"));

        wordMLPackage.setFontMapper(fontMapper);


        processing();
        Docx4J.toPDF(wordMLPackage, outStream);
//        FOSettings foSettings = Docx4J.createFOSettings();
//        foSettings.setWmlPackage(wordMLPackage);
//        Docx4J.toFO(foSettings, outStream, Docx4J.FLAG_EXPORT_PREFER_XSL);

        finished();
    }

    protected WordprocessingMLPackage getMLPackage(InputStream iStream) throws Exception {

        //Disable stdout temporarily as Doc convert produces alot of output
        System.setOut(new PrintStream(new OutputStream() {
            @Override
            public void write(int b) {
                //DO NOTHING
            }
        }));

        WordprocessingMLPackage mlPackage = Doc.convert(iStream);
        return mlPackage;
    }

}