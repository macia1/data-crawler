package word.articleoutput;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import word.entities.ArticleEntity;
import word.util.Util;
import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.math.BigInteger;
import java.net.URLDecoder;
import java.util.Arrays;
import java.util.List;


/**
 * @Author: 朝花夕誓
 * @Date: 2020/11/23 11:58
 * @Version 1.0
 */
@Slf4j
public class WordOutPut {

    private XWPFDocument doc = new XWPFDocument();

    /**
     * 文章内容写出
     * @param articleEntity 内容实体
     * @param outPutPath 文件路径
     * @throws Exception
     */
    public void wordWrite(List<ArticleEntity> articleEntity, String outPutPath) throws Exception {
        if (!outPutPath.endsWith(".docx")) throw new Exception("文件格式错误");
        articleEntity.forEach(article->{
            titleDecoration(article.getTitle());
            urlDecoration(article.getUrl());
            decoration("全文：", true);
            contentDecoration(article.getContent());
            decoration("", true);
            decoration("", true);
            decoration("", true);
            // 写完内容另起一页
//            doc.createParagraph().getCTP().addNewPPr().addNewSectPr();
        });

        // 文件写出
        try (FileOutputStream out = new FileOutputStream(outPutPath)){
            doc.write(out);
        } catch (IOException e) {
            log.info("文件写出失败....");
            e.printStackTrace();
        }
    }

    /**
     * 正文内容装饰
     * @param content
     */
    private void contentDecoration(String content){
        content = content.trim().replaceAll((char)12288 + "|[\\n\\r\\u00A0\\s\\t ]*", "");
        if(content.contains("。")){
            // 内容 写入
            String[] split = content.split("。");
            List<String> strings = Arrays.asList(split);
            for (int i = 0; i < strings.size(); i++) {
                String text = strings.get(i);
                while (text.length() < 30 && i < strings.size()-1){
                    int j = i+1;
                    log.info("文本过短： {}", text);
                    i=j;
                    String nextStr = strings.get(j);
                    text = text + "。" + nextStr;
                }
                if (!text.endsWith("。")){
                    decoration(text + "。", false);
                }else {
                    decoration(text, false);
                }
            }
        }else {
            decoration(content + "。", false);
        }
    }

    /**
     * 普通文本装饰器
     * @param doc 文件流
     * @param text 写入内容
     * @param isBold 是否加粗
     */
    private void decoration(String text, Boolean isBold){
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = commonDecoration(paragraph);
        run.setBold(isBold);
        run.setText(text);
    }

    /**
     * 标题装饰器
     */
    private void titleDecoration(String title){
        XWPFParagraph paragraph = doc.createParagraph();
        addCustomHeadingStyle(doc, "标题 2", 2);
        paragraph.setStyle("标题 2");
        XWPFRun run = commonDecoration(paragraph);
        run.setText(title);
        run.setBold(true);
    }

    /**
     * url装饰器
     */
    private void urlDecoration(String url){
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = commonDecoration(paragraph);
        run.setBold(true);
        run.setText("链接：");
        XWPFHyperlinkRun hyperlinkRun = paragraph.createHyperlinkRun(url);
        try {
            url = url.replaceAll("[?#].*", "");
            url = URLDecoder.decode(url, "utf-8");
        } catch (UnsupportedEncodingException e) {
            log.error("解码失败：{}", url);
            e.printStackTrace();
        }
        hyperlinkRun.setText(url);
        hyperlinkRun.setFontFamily("Microsoft YaHei");
        hyperlinkRun.setBold(true);
        hyperlinkRun.setUnderline(UnderlinePatterns.SINGLE);
        hyperlinkRun.setColor(Util.getColorString(new Color(5, 99, 193)));
        paragraph.addRun(hyperlinkRun);
    }

    /**
     * 通用段落配置
     * @return run
     */
    private XWPFRun commonDecoration(XWPFParagraph paragraph){
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        paragraph.setFirstLineIndent(450);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Microsoft YaHei");
        run.setFontSize(12);
        return run;
    }




    /**
     * 标题样式装饰器
     * @param docxDocument doc
     * @param strStyleId 标题名称
     * @param headingLevel 标题等级
     */
    public void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel ) {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        ctStyle.setQFormat(onoffnull);

        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        XWPFStyles styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);

    }
}
