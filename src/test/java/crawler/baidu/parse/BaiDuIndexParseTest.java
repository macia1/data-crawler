package crawler.baidu.parse;

import org.junit.Test;

/**
 * @Author 朝花夕誓
 * @Date 2021/2/25 17:23
 * @Version 1.0
 * @Description
 */
public class BaiDuIndexParseTest {

    String type01 = "搜索指数";
    String type02 = "资讯指数";
    String keyword = "联想";
    String cookie = "you self cookie";

    @Test
    public void test(){
        System.out.println(BaiDuIndexParse.getBaiDuIndexData(type01, keyword, "2020-01-01", "2020-05-01", cookie));
    }

}
