package word;

import com.deepoove.poi.xwpf.XWPFParagraphWrapper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;

public class WordBookmarkTest {

    @Test
    public void test01(){
        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph paragraph = doc.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("这是一个书签");
        run.setFontSize(24);
        run.setBold(true);

        // change line
        doc.createParagraph().getCTP().addNewPPr().addNewSectPr();
        doc.createParagraph().getCTP().addNewPPr().addNewSectPr();
        doc.createParagraph().getCTP().addNewPPr().addNewSectPr();

        XWPFParagraph paragraph1 = doc.createParagraph();
        XWPFRun run1 = paragraph1.createRun();
        run1.setText("这是一个坑");
        run1.setBold(true);

        //创建超链接
        XWPFParagraphWrapper xwpfParagraphWrapper = new XWPFParagraphWrapper(paragraph);
//        xwpfParagraphWrapper.insertNewHyperLinkRun(run1, "1");
        xwpfParagraphWrapper.insertNewHyperLinkRun(0, "http:deepoove.com");

        // 插入书签
        XWPFParagraphWrapper xwpfParagraphWrapper1 = new XWPFParagraphWrapper(paragraph1);
        xwpfParagraphWrapper1.insertNewBookmarkStart(123);
        xwpfParagraphWrapper1.insertNewBookmark(run1);


        try {
            FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\simpleTest.docx");
            doc.write(fileOutputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
