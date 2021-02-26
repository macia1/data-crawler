package word;

import com.alibaba.excel.EasyExcel;
import com.zhiwei.word.Entity.ExcelEntity;
import org.apache.logging.log4j.util.Strings;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * SCS 转图片版本
 */
public class WordGenerated03 {

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

    public static void main(String[] args) {
        String writePath = "C:\\Users\\Administrator\\Desktop\\simple03.docx";
        String excelPath = "D:\\CompanyWX\\WXWork\\1688854025252286\\Cache\\File\\2020-09\\数据上传模板.xlsx";


        WordGenerated03 wordGenerated03 = new WordGenerated03();

        List<ExcelEntity> entityList = wordGenerated03.readExcel(excelPath);

        HashMap<String, List<ExcelEntity>> dataMap = dataCount(entityList);

        wordGenerated03.write(dataMap, writePath);

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

    private void write(HashMap<String, List<ExcelEntity>> dataMap, String writePath) {
        XWPFDocument doc = new XWPFDocument();
        for(Map.Entry<String, List<ExcelEntity>> entry : dataMap.entrySet()){
            int rows = entry.getValue().size() * 3 + 1;
            XWPFTable table = doc.createTable(rows, 1);

            // table的第一行
            XWPFTableCell tableTitle = table.getRow(0).getCell(0);
            tableNoneBorder(table.getRow(0).getCell(0));
            tableTitle.setText(entry.getKey());
            FileInputStream stream = null;
            try {
                stream = new FileInputStream("D:\\maven-projects\\zhiwei-parent\\src\\main\\resources\\SCS.png");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }

            for (int s = 0; s < entry.getValue().size(); s ++) {
                for (int i = 1; i < 4; i++) {
                    XWPFParagraph paragraph;
                    XWPFTableCell cell = null;
                    cell = table.getRow(i + s * 3).getCell(0);

                    // 无边框装饰
                    tableNoneBorder(cell);

                    if (cell.getParagraphs().size() != 0){
                        paragraph = cell.getParagraphs().get(0);
                    }else {
                        paragraph = cell.addParagraph();
                    }
                    XWPFRun run = paragraph.createRun();
                    ExcelEntity excelEntity = entry.getValue().get(s);


                    switch (i){
                        case 1:
                            run.setFontFamily("BatangChe");
                            run.setBold(true);
                            run.setFontSize(12);
                            run.setText(excelEntity.getKoreanTitle());
                            break;
                        case 2:
                            run.setFontFamily("FangSong");
                            run.setFontSize(11);
                            run.setText(excelEntity.getTitle());
                            run.addBreak();

                            XWPFRun runTimeAndChannel = paragraph.createRun();
                            runTimeAndChannel.setFontSize(10);
                            runTimeAndChannel.setFontFamily("FangSong");

                            runTimeAndChannel.setText(excelEntity.getTime().replaceAll("[^\\d]", "-"));
                            runTimeAndChannel.setText("      ");
                            runTimeAndChannel.setText(excelEntity.getChannelDomainName());

                            CTTc ctTc = cell.getCTTc();
                            CTTcPr ctTcPr = ctTc.addNewTcPr();
                            ctTcPr.addNewTcBorders().addNewBottom().setVal(STBorder.SINGLE);
                            break;
                        case 3:
                            run.setFontFamily("BatangChe");
                            run.setFontSize(9);
                            run.setText(excelEntity.getKoreanSummary());
                            run.addBreak();
                            break;
                        default:
                            break;
                    }
                }
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

    private static HashMap<String, List<ExcelEntity>> dataCount(List<ExcelEntity> dataList) {
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

}
