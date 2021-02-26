package crawler.youku.parse;

import org.junit.jupiter.api.Test;

import javax.script.ScriptException;
import java.io.IOException;

/**
 * @Author 朝花夕誓
 * @Date 2021/2/25 17:25
 * @Version 1.0
 * @Description
 */
public class VideoBarrageCrawlerTest {

    VideoBarrageCrawler videoBarrageCrawler = new VideoBarrageCrawler();

    //    private String videoUrl = "https://v.youku.com/v_show/id_XNDk1MzY2NzgyMA==.html?spm=a2h0c.8166622.PhoneSokuProgram_1.dposter&s=aaed627feea749d7a99d";

    // 超前点播url
    private String videoUrl = "https://v.youku.com/v_show/id_XNDk1ODkwNzcyNA==.html?s=aaed627feea749d7a99d";

    @Test
    void test() throws NoSuchMethodException, ScriptException, IOException {
        videoBarrageCrawler.getBarrage(videoUrl);
    }

}
