package word;

import com.alibaba.excel.EasyExcel;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.style.TableStyle;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.xwpf.XWPFParagraphWrapper;
import com.zhiwei.word.Entity.ExcelEntity;
import org.apache.logging.log4j.util.Strings;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;


/**
 * word 表格生成
 * @author 朝花夕誓
 * @date 2020/09/07
 */
public class WordGenerated {


    public static String[] headLine = new String[]{"No", "Headline", "Media Name", "Publish Date"};

    /**
     * 读取源文件
     * @param filePath
     * @return
     */
    public List<ExcelEntity> readExcel(String filePath){
        Objects.requireNonNull(filePath, "filePath 不能为空");
        List<ExcelEntity> entityList = EasyExcel.read(filePath, ExcelEntity.class, null).sheet().doReadSync();
        List<ExcelEntity> dataList = new ArrayList<>();
        entityList.forEach(i->{
            if (Strings.isNotBlank(i.getTime())){
                dataList.add(i);
            }
        });
        return dataList;
    }

    /**
     * 写表格
     */
    public HashMap<String, List<ExcelEntity>> writeTables(XWPFDocument doc, List<ExcelEntity> dataList, String writePath){
        HashMap<String, List<ExcelEntity>> brandCategoryHashMap = new HashMap<>();
        dataList.forEach(item->{
            // create table
            String brandCategory = item.getBrandCategory();
            if (brandCategoryHashMap.get(brandCategory) == null){
                ArrayList<ExcelEntity> excelEntities = new ArrayList<>();
                excelEntities.add(item);
                brandCategoryHashMap.put(brandCategory, excelEntities);
            }else {
                brandCategoryHashMap.get(brandCategory).add(item);
            }
        });

        for(Map.Entry<String, List<ExcelEntity>> entry : brandCategoryHashMap.entrySet()){
            System.out.println("entry key is " + entry.getKey() + "   entry vaule size is : " + entry.getValue().size());
            XWPFParagraph paragraph = doc.createParagraph();
            // 每张表格的表头
            XWPFRun run = paragraph.createRun();
            run.setFontFamily("Microsoft YaHei");
            run.setBold(true);
            run.setFontSize(12);
            run.setText(entry.getKey());

            // 设置表格 行数 列数
            XWPFTable table = doc.createTable(entry.getValue().size() + 1, 4);
            // 设置表格宽度为A4纸最大宽度
            TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, 10);

            //设置表格 居中
            TableStyle style = new TableStyle();
            style.setAlign(STJc.CENTER);
            TableTools.styleTable(table, style);

            // 表头 样式设置
            for (int i = 0; i < 4; i++){
                XWPFTableCell cell = table.getRow(0).getCell(i);
                cell.setColor("0099CC");
                XWPFParagraph tableParagraph = null;
                if(cell.getParagraphs().size() != 0){
                    tableParagraph = cell.getParagraphs().get(0);
                }else{
                    tableParagraph = cell.addParagraph();
                }

                // 垂直
                tableParagraph.setVerticalAlignment(TextAlignment.CENTER);
                // 水平
                tableParagraph.setAlignment(ParagraphAlignment.CENTER);

                XWPFRun cellRun = tableParagraph.createRun();
                cellRun.setFontFamily("Arial");
                cellRun.setText(headLine[i]);
                cellRun.setFontSize(9);
                cellRun.setBold(true);
                cellRun.setColor("FFFFFF");
//                cellRun.addBreak(BreakType.TEXT_WRAPPING);
            }

            // 控制行数
            for (int i = 1; i < entry.getValue().size() + 1; i ++){
                // 控制列数
                for (int j = 0; j < 4; j ++) {
                    XWPFTableCell cell = table.getRow(i).getCell(j);
                    //垂直居中
//                    cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    // 水平居中
//                    cell.getCTTc().getPList().get(0).getPPr().getJc().setVal(STJc.CENTER);
                    XWPFParagraph paragraphCell = cell.addParagraph();
                    paragraphCell.setVerticalAlignment(TextAlignment.CENTER);
                    paragraphCell.setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun runCell = paragraphCell.createRun();
                    runCell.setFontFamily("Microsoft YaHei");
                    runCell.setFontSize(9);
//                    runCell.setColor("0000f0");
                    cell : switch (j){
                        case 0:
                            runCell.setText(String.valueOf(i));
                            break cell;
                        case 1:
                            String koreanTitle = entry.getValue().get(i - 1).getKoreanTitle();
                            String title = entry.getValue().get(i - 1).getTitle();
                            runCell.setText(koreanTitle);
                            runCell.addBreak();
                            runCell.setText(title);
                            break cell;
                        case 2:
                            String channelDomainName = entry.getValue().get(i - 1).getChannelDomainName();
                            String channel = entry.getValue().get(i - 1).getChannel();
                            runCell.setText(channelDomainName);
                            runCell.addBreak();
                            runCell.setText(channel);
                            break cell;
                        case 3:
                            String time = entry.getValue().get(i - 1).getTime();
                            String s = time.replaceAll("[^\\d]", ".");
                            runCell.setText(s);
                            break cell;
                        default:
                            System.err.println("写入有误。");
                            break cell;
                    }
                }
            }
            // 写完每张表进行换行
            doc.createParagraph();
        }

        // 文件写出
        try (FileOutputStream out = new FileOutputStream(writePath)){
            doc.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return brandCategoryHashMap;

    }

    public static XWPFRun generalTitle(XWPFRun generalTitleConf){
        generalTitleConf.setBold(true);
        generalTitleConf.setFontFamily("Arial");
        generalTitleConf.setFontSize(10);
        return generalTitleConf;
    }

    public static XWPFRun generalContent(XWPFRun generalContentConf){
        generalContentConf.setText(" ");
        generalContentConf.setFontFamily("Microsoft YaHei");
        generalContentConf.setFontSize(9);
        return generalContentConf;
    }

    /**
     * 写文档
     */
    public void writeDoc(HashMap<String, List<ExcelEntity>> stringListHashMap, String writePath){
        XWPFDocument doc = null;
        try {
            doc = new XWPFDocument(new FileInputStream(writePath));
            XWPFParagraph paragraph = doc.createParagraph();
            // 首次追加进行另起一页
            CTPPr ctpPr = paragraph.getCTP().addNewPPr();
            ctpPr.addNewSectPr();
        } catch (IOException e) {
            e.printStackTrace();
        }

        for(Map.Entry<String, List<ExcelEntity>> entry :stringListHashMap.entrySet()){
            for (ExcelEntity excelEntity : entry.getValue()) {
                XWPFParagraph paragraph = doc.createParagraph();

                // 追加内容进行另起一页
                CTPPr ctpPr = paragraph.getCTP().addNewPPr();
                ctpPr.addNewSectPr();

                // Headline 内容写入
                XWPFRun headLine = paragraph.createRun();
                generalTitle(headLine);
                headLine.setText("Headline:");

                XWPFRun headLineContent = paragraph.createRun();
                generalContent(headLineContent);
                headLineContent.setText(excelEntity.getKoreanTitle());
                headLineContent.addCarriageReturn();
                headLineContent.setText("                  ");
                headLineContent.setText(excelEntity.getTitle());
//                CTBookmark ctBookmark = new XWPFParagraphWrapper(paragraph).insertNewBookmark(headLine);
//                ctBookmark.setName("1");

                // Publication 内容写入
                XWPFRun publication = paragraph.createRun();
                generalTitle(publication);
                publication.addCarriageReturn();
                publication.setText("Publication:");
                XWPFRun publicationContent = paragraph.createRun();
                generalContent(publicationContent);
                publicationContent.setText(excelEntity.getChannelDomainName() + " / " + excelEntity.getChannel());

                // Paper Date 内容写入
                XWPFRun paperDate = paragraph.createRun();
                paperDate.addCarriageReturn();
                paperDate.setText("Paper Date:");
                generalTitle(paperDate);
                XWPFRun paperDateContent = paragraph.createRun();
                generalContent(paperDateContent);
                paperDateContent.setText(excelEntity.getTime().replaceAll("[^\\d]", "."));

                // 韩文内容 写入
                XWPFRun summary = paragraph.createRun();
                generalTitle(summary);
                summary.addCarriageReturn();
                summary.setText("Summary:");
                summary.addCarriageReturn();
                if (!"".equals(excelEntity.getKoreanSummary())){
                    XWPFRun summaryContent = paragraph.createRun();
                    summaryContent.setText(excelEntity.getKoreanSummary());
                    summaryContent.setFontSize(9);
                    summaryContent.setFontFamily("BatangChe");
                    summaryContent.addCarriageReturn();
                    summaryContent.addCarriageReturn();
                    summaryContent.addCarriageReturn();
                }

                XWPFRun summaryContent = paragraph.createRun();

                // 中文内容 写入
                String[] split = excelEntity.getContent().split("。\\s");
                List<String> strings = Arrays.asList(split);
                Iterator<String> iterator = strings.iterator();
                while(iterator.hasNext()){
                    summaryContent.setText("       ");
                    String next = iterator.next();
                    if (!next.endsWith("。")){
                        summaryContent.setText(next + "。");
                    }else {
                        summaryContent.setText(next);
                    }
                    summaryContent.addCarriageReturn();
                }
                generalContent(summaryContent);
            }
        }
        // 文件写出
        try (FileOutputStream out = new FileOutputStream(writePath)){
            doc.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void addMarkBook(String readPath){
        XWPFDocument doc = null;
        try {
            doc = new XWPFDocument(new FileInputStream(readPath));
        } catch (IOException e) {
            e.printStackTrace();
        }
//        // 进行所有的表格读取
//        List<XWPFTable> tables = doc.getTables();
//        System.out.println("tables size is :" + tables.size());
//        tables.forEach(tb->{
//            List<XWPFTableRow> rows = tb.getRows();
//            rows.forEach(row->{
//                XWPFTableCell cell = row.getCell(1);
//                String text = cell.getText();
//            });
//            IBody body = tb.getBody();
//            List<XWPFParagraph> paragraphs = body.getParagraphs();
//            paragraphs.forEach(it->{
//                List<XWPFRun> runs = it.getRuns();
//                List<XWPFRun> runs = it.getRuns();
//                for (XWPFRun run : runs) {
//                    System.out.println(run.getText(run.getTextPosition()));
//                }
//                String text = it.getText();
//                System.out.println(text + "--------------");
//            });
//        });

        // 扫描 段落
        Iterator<XWPFParagraph> paragraphsIterator = doc.getParagraphsIterator();
        Map<String, XWPFParagraph> dataMap = new HashMap<>();
        while(paragraphsIterator.hasNext()){
            XWPFParagraph next = paragraphsIterator.next();
            String text = next.getText();
            System.out.println("-------------------------------------------------------------------");
            System.out.println(text);
            System.out.println("-------------------------------------------------------------------");
            if(dataMap.get(next) == null){
                dataMap.put(text, next);
            }else {
                XWPFParagraph paragraph = dataMap.get(text);
                List<XWPFRun> runs = paragraph.getRuns();
                XWPFParagraphWrapper wrapper = new XWPFParagraphWrapper(next);
                for (XWPFRun run : runs) {
                    wrapper.insertNewBookmark(run);
                }
            }
        }

        System.out.println("===========================================================================");

        // 扫描 table
        Iterator<XWPFTable> tablesIterator = doc.getTablesIterator();
        XWPFTable table;
        while (tablesIterator.hasNext()){
            table = tablesIterator.next();
            XWPFTableRow row;
            List<XWPFTableCell> cells;
            for (int i = 0; i < table.getNumberOfRows(); i++) {
                row = table.getRow(i);
                cells = row.getTableCells();
//                for (int j = 0; j < cells.size(); j++) {
                    String text = cells.get(1).getText();
                    System.out.println(text + "********************************");
                    for (Map.Entry<String, XWPFParagraph> entry : dataMap.entrySet()){
                        if(entry.getKey().contains(text.split("\\s")[0])){
//                            XWPFRun run = cells.get(1).addParagraph().createRun();
//                            run.setText(text);
//                            cells.get(2).removeParagraph(1);
                            System.out.println(text + "========================================11111=======================================================");
                            List<XWPFParagraph> paragraphs = cells.get(1).getParagraphs();
                            paragraphs.forEach(it->{
                                XWPFParagraph value = entry.getValue();
                                String text1 = value.getText();
                                System.out.println(text1 + "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^");
                                if (!"".equals(text1)){
//                                    XWPFParagraphWrapper xwpfParagraphWrapper = new XWPFParagraphWrapper(entry.getValue());
//                                    xwpfParagraphWrapper.insertNewBookmark(it.getRuns().get(0));
//                                    it.getRuns().forEach(ru->{
//                                        System.out.println("jjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjjj");
//                                    });
                                }

                            });
                            System.out.println("*****************************");
                            System.out.println(text);
                            System.out.println("*****************************");
                        }
                    }

//                    List<XWPFParagraph> paragraphs = cells.get(j).getParagraphs();
//                    paragraphs.forEach(item->{
//                        item.getRuns().forEach(run->{
//                            System.out.println("*****************************");
//                        });
//                    });
//                }

            }
        }



        // 文件写出
        try (FileOutputStream out = new FileOutputStream(readPath)){
            doc.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        // 写出路径
        String writePath = "C:\\Users\\Administrator\\Desktop\\simple.docx";
        // 创建文档
        XWPFDocument doc = new XWPFDocument();
        WordGenerated wordGenerated = new WordGenerated();

        String filePath = "D:\\CompanyWX\\WXWork\\1688854025252286\\Cache\\File\\2020-09\\数据上传模板.xlsx";
        List<ExcelEntity> entityList = wordGenerated.readExcel(filePath);

        // table write in
        HashMap<String, List<ExcelEntity>> stringListHashMap = wordGenerated.writeTables(doc, entityList, writePath);

        // document write in
        wordGenerated.writeDoc(stringListHashMap, writePath);

//        addMarkBook(writePath);

    }

}
