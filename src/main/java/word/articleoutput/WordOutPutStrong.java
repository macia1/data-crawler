package word.articleoutput;

import lombok.extern.slf4j.Slf4j;
import org.apache.logging.log4j.util.Strings;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import word.entities.ArticleEntity;
import word.util.Util;
import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;

/**
 * @Author: 朝花夕誓
 * @Date: 2020/11/23 11:58
 * @Version 1.0
 */
@Slf4j
public class WordOutPutStrong {

    private XWPFDocument doc = new XWPFDocument();

    /**
     * 文章内容写出
     * @param articleEntity 内容实体
     * @param outPutPath 文件路径
     * @throws Exception
     */
    public void wordWrite(List<ArticleEntity> articleEntity, String outPutPath) throws Exception {
        if (!outPutPath.endsWith(".docx")) throw new Exception("文件格式错误");
        if (Objects.isNull(articleEntity)) throw new NullPointerException("传入数据不能为空");
        // 进行目录生成
        tableOfContent(articleEntity);
        for (int i = 0; i < articleEntity.size(); i++) {
            ArticleEntity article = articleEntity.get(i);
            String title = article.getTitle();
            String url = article.getUrl();
            String content = article.getContent();
            if (beanJudgment(article)){
                titleDecoration(title);
                urlDecoration(url);
                decoration("全文：", true);
                contentDecoration(content);
                if (i != articleEntity.size()-1){
                    // 写完内容另起一页
                    doc.createParagraph().getCTP().addNewPPr().addNewSectPr();
                }
            }
        }

        // 文件写出
        try (FileOutputStream out = new FileOutputStream(outPutPath)){
            doc.write(out);
        } catch (IOException e) {
            log.info("文件写出失败....");
            e.printStackTrace();
        }finally {
            doc.close();
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
        // 添加书签
        CTBookmark ctBookmark = paragraph.getCTP().addNewBookmarkStart();
        ctBookmark.setName(String.valueOf(Math.abs(title.hashCode())));
        ctBookmark.setId(BigInteger.valueOf(0));
        paragraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(0));
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
        url = url.replaceAll((char)12288 + "|\\s|\\u00A0", "");
        run.setText("链接：" + url);
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

    /**
     * 创建书签
     */
    public void tableOfContent(List<ArticleEntity> dataList){
        XWPFParagraph list = doc.createParagraph();
        XWPFRun listRun = list.createRun();
        listRun.setText("目录");
        listRun.setColor(Util.getColorString(new Color(5, 99, 193)));
        listRun.setFontFamily("等线 Light");
        listRun.setFontSize(20);

        dataList.forEach(data->{
            if (beanJudgment(data)) {
                String title = data.getTitle();
                XWPFParagraph paragraph = doc.createParagraph();
                paragraph.setAlignment(ParagraphAlignment.LEFT);
                paragraph.setIndentFromLeft(550);
//                CTPPr ppr = paragraph.getCTP().getPPr();
//                if (ppr == null) ppr = paragraph.getCTP().addNewPPr();
//                CTSpacing spacing = ppr.isSetSpacing()? ppr.getSpacing() : ppr.addNewSpacing();
//                spacing.setAfter(BigInteger.valueOf(20));
//                spacing.setBefore(BigInteger.valueOf(0));
//                spacing.setLineRule(STLineSpacingRule.AUTO);
//                spacing.setLine(BigInteger.valueOf(240));
                paragraph.setSpacingAfter(200);
                XWPFHyperlinkRun hyperlinkRun = paragraph.createHyperlinkRun("#" + Math.abs(title.hashCode()));
                hyperlinkRun.setFontFamily("Microsoft YaHei");
                hyperlinkRun.setFontSize(10);
                hyperlinkRun.setBold(true);
                hyperlinkRun.setText("\t" + title);
            }
        });
        doc.createParagraph().getCTP().addNewPPr().addNewSectPr();
    }

    /**
     * 判断实体是否满足条件
     */
    public Boolean beanJudgment(ArticleEntity articleEntity){
        if (Strings.isNotBlank(articleEntity.getTitle()) && Strings.isNotBlank(articleEntity.getUrl()) && Strings.isNotBlank(articleEntity.getContent())){
            return true;
        }
        return false;
    }

}
