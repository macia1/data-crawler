package word;

import com.alibaba.excel.EasyExcel;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.style.TableStyle;
import com.deepoove.poi.util.TableTools;
import com.zhiwei.word.Entity.ExcelEntity;
import org.apache.logging.log4j.util.Strings;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.*;

public class WordGenerated02 {

    public static HashMap<String, List<ExcelEntity>> getDataMap(List<ExcelEntity> dataList){
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
        return brandCategoryHashMap;
    }

    public static void tableNoneBorder(XWPFTableCell cell){
        CTTc ctTc = cell.getCTTc();
        CTTcPr ctTcPr = ctTc.addNewTcPr();
        ctTc.getTcPr().addNewTcBorders().addNewRight().setVal(STBorder.NIL);    //设置无边框
        ctTc.getTcPr().addNewTcBorders().addNewLeft().setVal(STBorder.NIL);
        ctTc.getTcPr().addNewTcBorders().addNewTop().setVal(STBorder.NIL);
        ctTc.getTcPr().addNewTcBorders().addNewBottom().setVal(STBorder.NIL);
//        ctTcPr.addNewVAlign().setVal(STVerticalJc.CENTER);  // 字体居中
    }

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

    public void writeTables(XWPFDocument doc, HashMap<String, List<ExcelEntity>> brandCategoryHashMap, String writePath) {
        for (Map.Entry<String, List<ExcelEntity>> entry : brandCategoryHashMap.entrySet()) {
            // 设置表格 行数 列数
            XWPFTable table = doc.createTable(entry.getValue().size() + 1, 2);
            doc.insertTable(0, table);
            // 设置表格宽度为A4纸最大宽度
            TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, 10);

            //设置表格 居中
            TableStyle style = new TableStyle();
            style.setAlign(STJc.CENTER);
            TableTools.styleTable(table, style);

            // 表格合并
            TableTools.mergeCellsHorizonal(table, 0, 0, 1);
            // 表头 样式设置
            XWPFTableCell headCell = table.getRow(0).getCell(0);
            tableNoneBorder(headCell);
            headCell.setColor("0099CC");
//            XWPFParagraph tableParagraph = headCell.addParagraph();
            XWPFParagraph tableParagraph = null;
            if(headCell.getParagraphs().size() != 0){
                tableParagraph = headCell.getParagraphs().get(0);
            }else {
                tableParagraph = headCell.addParagraph();
            }
            // 垂直
            tableParagraph.setVerticalAlignment(TextAlignment.CENTER);
            // 水平
//            tableParagraph.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun cellRun = tableParagraph.createRun();
            cellRun.setFontFamily("BatangChe");
            cellRun.setText(entry.getValue().get(0).getKoreanTitle());
            cellRun.setFontSize(14);
            cellRun.setBold(true);
            cellRun.setColor("FFFFFF");
//            cellRun.addBreak(BreakType.TEXT_WRAPPING);

            // 控制行数
            for (int i = 1; i < entry.getValue().size() + 1; i++) {
                // 控制列数
                for (int j = 0; j < 2; j++) {
//                    XWPFTableRow row = table.getRow(i);
                    XWPFTableCell cell = table.getRow(i).getCell(j);
                    XWPFParagraph paragraphCell = cell.addParagraph();
                    XWPFRun runCell = paragraphCell.createRun();
                    ExcelEntity excelEntity = entry.getValue().get(i - 1);

                    // 无边框
                    tableNoneBorder(cell);
                    cell:
                    switch (j) {
                        case 0:
                            runCell.setFontFamily("BatangChe");
                            runCell.setFontSize(12);
                            String koreanTitle = excelEntity.getKoreanTitle();
                            runCell.setText(koreanTitle);

//                            runCell.addCarriageReturn();
                            runCell.addBreak(BreakType.TEXT_WRAPPING);

                            String title = excelEntity.getTitle();
                            XWPFRun chineseRun = paragraphCell.createRun();
                            chineseRun.setText(title);
                            chineseRun.setFontSize(11);
                            chineseRun.setFontFamily("Microsoft YaHei");
                            break cell;
                        case 1:
                            // 设置单元格宽度
                            CTTcPr ctTcPr = cell.getCTTc().addNewTcPr();
                            CTTblWidth ctTblWidth = ctTcPr.addNewTcW();
                            ctTblWidth.setType(STTblWidth.DXA);
                            ctTblWidth.setW(BigInteger.valueOf(360 * 4));

                            // 表格内容垂直居中
                            CTVerticalJc ctVerticalJc = ctTcPr.addNewVAlign();
                            ctVerticalJc.setVal(STVerticalJc.CENTER);

                            // 表格内容水平居中
//                            CTTc ctTc = cell.getCTTc();
//                            ctTc.getPList().get(i).addNewPPr().addNewJc().setVal(STJc.CENTER);

                            String time = excelEntity.getTime();
                            String s = time.replaceAll("[^\\d]", "-");
                            runCell.setFontSize(12);
                            runCell.setFontFamily("Calibri");
                            runCell.setText("       " +  s);
                            break cell;
                        default:
                            System.err.println("写入有误。");
                            break cell;
                    }
                }
                doc.createParagraph();
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

    /**
     * 韩文字体装饰
     */
    public static void koreanDecoration(XWPFRun run){
        run.setFontSize(10);
        run.setFontFamily("BatangChe");
        run.setBold(true);
    }

    /**
     * 中文字体装饰
     * @param args
     */
    public static void chineseDecoration(XWPFRun run){
        run.setFontFamily("Microsoft YaHei");
        run.setFontSize(11);
    }

    public static void main(String[] args) {
        // 写出路径
        String writePath = "C:\\Users\\Administrator\\Desktop\\simple02.docx";
        // 创建文档
        XWPFDocument doc = new XWPFDocument();
        WordGenerated02 wordGenerated02 = new WordGenerated02();

        String filePath = "D:\\CompanyWX\\WXWork\\1688854025252286\\Cache\\File\\2020-09\\数据上传模板.xlsx";
        List<ExcelEntity> entityList = wordGenerated02.readExcel(filePath);
        // 数据统计结果
        HashMap<String, List<ExcelEntity>> dataMap = getDataMap(entityList);
        // 表格写入
        wordGenerated02.writeTables(doc, dataMap, writePath);

        // 文章写入
        wordGenerated02.articleWriteIn(dataMap, writePath);

    }

    private void articleWriteIn(HashMap<String, List<ExcelEntity>> dataMap, String writePath) {
        XWPFDocument doc = null;
        try {
            doc = new XWPFDocument(new FileInputStream(writePath));
            // 首次写内容进行换行
            doc.createParagraph().getCTP().addNewPPr().addNewSectPr();
        } catch (IOException e) {
            e.printStackTrace();
        }
        for (Map.Entry<String, List<ExcelEntity>> entry : dataMap.entrySet()){
            for (int i = 0; i < entry.getValue().size(); i ++){
                XWPFTable table = doc.createTable(5, 2);
                TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, 10);
                TableTools.mergeCellsHorizonal(table, 3, 0, 1);
                TableTools.mergeCellsHorizonal(table, 4, 0, 1);

                TableStyle style = new TableStyle();
                style.setAlign(STJc.CENTER);

                for (int r = 0; r < 5; r ++) {
                    for (int j = 0; j < 2; j++){
                        if(r != 3 && r != 4){
                            XWPFTableCell cell = table.getRow(r).getCell(j);
                            tableNoneBorder(cell);
                            XWPFParagraph paragraph;
                            if(cell.getParagraphs().size() != 0){
                                paragraph = cell.getParagraphs().get(0);
                            }else {
                                paragraph = cell.addParagraph();
                            }
                            XWPFRun run = paragraph.createRun();
                            String judgment = r + String.valueOf(j);
                            ExcelEntity excelEntity = entry.getValue().get(i);
                            cell : switch(judgment){
                                case "00":
                                    koreanDecoration(run);
                                    run.setText("뉴스 타이틀:");
                                    break cell;
                                case "01":
                                    run.setFontFamily("BatangChe");
                                    run.setText(excelEntity.getKoreanTitle());
                                    XWPFRun chineseRun = cell.addParagraph().createRun();
                                    chineseRun.setText(excelEntity.getTitle());
                                    chineseDecoration(chineseRun);
                                    break cell;
                                case "10":
                                    koreanDecoration(run);
                                    run.setText("출판사:");
                                    break;
                                case "11":
                                    run.setText(excelEntity.getChannelDomainName() + " / " + excelEntity.getChannel());
                                    chineseDecoration(run);
                                    break;
                                case "20":
                                    run.setText("기사날짜:");
                                    koreanDecoration(run);
                                    break;
                                case "21":
//                                    // 设置单元格宽度
//                                    CTTblWidth ctTblWidth = cell.getCTTc().addNewTcPr().addNewTcW();
//                                    ctTblWidth.setType(STTblWidth.DXA);
//                                    ctTblWidth.setW(BigInteger.valueOf(360 * 7));
                                    run.setText(excelEntity.getTime().replaceAll("[^\\d]", "-"));
                                    run.setFontSize(11);
                                    run.setFontFamily("Arial");
                                    break;
                                default:
                                    System.err.println("(!3,!4)超出选择范围" + j);
                            }
                        }

                    }

                    if(r == 3 || r == 4){
                        XWPFTableCell cell = table.getRow(r).getCell(0);
                        tableNoneBorder(cell);
                        XWPFParagraph paragraph = cell.getParagraphs().get(0);
                        XWPFRun run = paragraph.createRun();
                        run.setFontFamily("BatangChe");
                        cell : switch(r){
                            case 3:
                                run.setText("개요:");
                                koreanDecoration(run);
                                break cell;
                            case 4:
                                paragraph.setVerticalAlignment(TextAlignment.CENTER);
                                run.addBreak();
                                run.setFontSize(12);
                                run.setText(entry.getValue().get(i).getKoreanSummary());
                                break cell;
                            default:
                                System.err.println("====");
                        }
                    }
                }

                // 写中文文档
                XWPFParagraph paragraph = doc.createParagraph();
                XWPFRun chineseRunTitle = paragraph.createRun();

                chineseRunTitle.addCarriageReturn();
                chineseRunTitle.addCarriageReturn();

                ExcelEntity excelEntity = entry.getValue().get(i);
                chineseDecoration(chineseRunTitle);
                chineseRunTitle.setBold(true);
                chineseRunTitle.setText(excelEntity.getTitle());

                XWPFRun run = paragraph.createRun();
                chineseDecoration(run);
                run.addCarriageReturn();
                run.setText(excelEntity.getTime().replaceAll("[^\\d]", "-") + "   " + "来源于: " + excelEntity.getChannel());
                run.addCarriageReturn();
                // 中文内容 写入
                String[] split = excelEntity.getContent().split("。\\s");
                List<String> strings = Arrays.asList(split);
                Iterator<String> iterator = strings.iterator();
                while(iterator.hasNext()){
                    run.setText("       ");
                    String next = iterator.next();
                    if (!next.endsWith("。")){
                        run.setText(next + "。");
                    }else {
                        run.setText(next);
                    }
                    run.addCarriageReturn();
                }
                // 写完内容进行换行
                doc.createParagraph().getCTP().addNewPPr().addNewSectPr();
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

}
