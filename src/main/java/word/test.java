package word;


import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.style.TableStyle;
import com.deepoove.poi.util.TableTools;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

import java.io.*;
import java.util.List;

/**
 * java 操作word测试
 * @author 朝花夕誓
 * @date 2020/09/07
 */
public class test {

    /**
     * 1.文档XWPFDocument
     */
    public void createNewDocument(){
        /**
         * 1. 文档XWPFDocument
         */
        // 1.1创建新文档
        XWPFDocument doc = new XWPFDocument();
        // 1.2读取已有文档：段落、表格、图片
        try {
            XWPFDocument document = new XWPFDocument(new FileInputStream("./deepoove.docx"));
        } catch (IOException e) {
            e.printStackTrace();
        }
        // 段落
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        // 表格
        List<XWPFTable> tables = doc.getTables();
        // 图片
        List<XWPFPictureData> allPictures = doc.getAllPictures();
        // 页眉
        List<XWPFHeader> headerList = doc.getHeaderList();
        // 页脚
        List<XWPFFooter> footerList = doc.getFooterList();

        // 1.3生成文档
        try (FileOutputStream out = new FileOutputStream("simple.docx")){
            doc.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        /**
         *  2.  段落XWPFParagraph
         */
        // 2.1创建新段落
        XWPFParagraph p1 = doc.createParagraph();
        // 2.2设置段落格式
        // 对齐方式
        p1.setAlignment(ParagraphAlignment.CENTER);
        // 边框
        p1.setBorderBottom(Borders.DOUBLE);
        p1.setBorderTop(Borders.DOUBLE);
        p1.setBorderRight(Borders.DOUBLE);
        p1.setBorderLeft(Borders.DOUBLE);
        p1.setBorderBetween(Borders.SINGLE);
        // 2.3 基本元素XWPFRun
        // 2.4 段落文本
        // 2.4.1 读取段落内容
        for (XWPFParagraph paragraph : paragraphs) {
            // 获取文字
            String text = paragraph.getText();

            // 获取段落内所有的 XWPFRun
            List<XWPFRun> runs = paragraph.getRuns();

            // 2.4.2 创建XWPFRun文本
            // 段落末尾创建XWPFRun
            XWPFRun run = paragraph.createRun();
            run.setText("为这个段落追加文本");

            // 2.4.3 插入XWPFRun文本
            // 段落起始插入XWPFRun
            XWPFRun insertNewRun = paragraph.insertNewRun(0);
            insertNewRun.setText("在段落起始位置插入这个文本");

            // 2.4.4 修改XWPFRun文本
            List<XWPFRun> updateRuns = paragraph.getRuns();
            // setText默认为追加文本，参数0表示设置第0个位置的文本，覆盖上一次设置
            updateRuns.get(0).setText("追加文本", 0);
            updateRuns.get(0).setText("修改文本", 0);

            // 2.4.5 样式：颜色、字体
            // 颜色
            run.setColor("00ff00");
            // 斜体
            run.setItalic(true);
            // 粗体
            run.setBold(true);
            // 字体
            run.setFontFamily("Courier");
            // 下划线
            run.setUnderline(UnderlinePatterns.DOT_DOT_DASH);

            // 2.4.6 文本换行
            run.addCarriageReturn();

            /**
             * 2.5 段落图片
             */
            // 2.5.1 提取图片XWPFPicture
//        List<XWPFPictureData> allPictures = doc.getAllPictures();
            XWPFPictureData pictureData = allPictures.get(0);
            byte[] data = pictureData.getData();
            // 接下来就可以将图片字节数组写入输出流
            // 2.5.2 创建XWPFRun图片
            try {
                InputStream stream = new FileInputStream("./sayi.png");
                XWPFRun runPic = paragraphs.get(0).createRun();
                runPic.addPicture(stream, XWPFDocument.PICTURE_TYPE_PNG, "Generated", Units.toEMU(256), Units.toEMU(256));
                Units.toEMU(256);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        /**
         * 3.  创建表格XWPFTable
         */
        // 3.1 创建新表格
        XWPFTable table = doc.createTable(3, 3);

        // 3.2设置单元格文本
        table.getRow(1).getCell(1).setText("EXAMPLE OF TABLE");
        // 上面这一段代码和下面这一段代码是等价的：
        XWPFParagraph p = table.getRow(0).getCell(0).addParagraph();
        XWPFRun r1 = p.createRun();
        r1.setText("EXAMPLE OF TABLE");

        // 3.3设置单元格图片
        // 图片操作其实就是获取段落，然后等同操作段落中的图片。
        XWPFParagraph paragraph = table.getRow(0).getCell(0).addParagraph();
        XWPFRun rr = p1.createRun();
        // 同段落图片
        // 3.4 设置单元格样式：背景色、对齐方式
        // 背景色
        table.getRow(1).getCell(1).setColor("");
        // 获取单元格段落后设置对其方式
//        XWPFParagraph addPrargrapph = table

        /**
         * poi-tl: Word 模板引擎
         */
        // 4.1 TableTools
        // 4.1.1 用一行单元格合并某几列
        // 合并第一行的第0列到第8列单元格
        TableTools.mergeCellsHorizonal(table, 1, 0, 8);

        // 4.1.2  同一列单元格合并某几行
        // 合并第0列的第一行到第九行的单元格
        TableTools.mergeCellsVertically(table, 0, 1, 9);

        // 4.1.3 表格宽度
        // 设置表格宽度为A4纸最大宽度
        TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, 10);

        // 4.1.4 表格样式
        //设置表格居中
        TableStyle style = new TableStyle();
        style.setAlign(STJc.CENTER);
        TableTools.styleTable(table, style);


    }

}
