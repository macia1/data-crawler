//package word.articleoutput;
//
//
//import org.apache.commons.lang3.StringUtils;
//import org.junit.jupiter.api.Test;
//import word.entities.ArticleEntity;
//
//import java.util.ArrayList;
//import java.util.List;
//import java.util.Objects;
//
//
///**
// * @Author: 朝花夕誓
// * @Date: 2020/11/23 14:28
// * @Version 1.0
// */
//public class WordOutPutTest {
//
//    @Test
//    public void test() throws Exception {
//        WordOutPut wordOutPut = new WordOutPut();
//        String writePath = "C:\\Users\\Administrator\\Desktop\\future.docx";
//        List<ArticleEntity> data = getData();
//        if (!Objects.isNull(data) && data.size() > 0) {
//            wordOutPut.wordWrite(data, writePath);
//        }
//    }
//
//    @Test
//    public void test03(){
//        System.out.println((int)' ');
//        System.out.println((int)'　');
//        System.out.println((char)12288);
//        System.out.println((int)' ');
//        getData();
//    }
//
//
//    public List<ArticleEntity> getData() {
//        MongoClient mongoClient = new MongoClient("127.0.0.1", 27017);
//        MongoDatabase database = mongoClient.getDatabase("search_data");
//        MongoCollection<Document> data = database.getCollection("keywordTest");
//        List<String> keywordList = new ArrayList<>();
//        keywordList.add("安卓逆向");
//        keywordList.add("安卓协议分析");
//        keywordList.add("Android 逆向");
//        keywordList.add("Android 反编译");
//        keywordList.add("安卓反编译");
//        keywordList.add("Android 静态调试");
//        keywordList.add("Android 动态调试");
//        keywordList.add("Android 调试");
////        keywordList.add("Java 布隆过滤器");
////        keywordList.add("Java 布谷鸟过滤器");
//        Bson filter = Filters.in("word", keywordList);
//        MongoCursor<Document> iterator = data.find(filter).iterator();
//        List<ArticleEntity> articleEntities = new ArrayList<>();
//        stop:
//        while (iterator.hasNext()) {
//            Document next = iterator.next();
//            ArticleEntity articleEntity = new ArticleEntity();
//            String url = next.getString("url");
//            String title = next.getString("title");
//            String content = next.getString("content");
//            if (StringUtils.isNotBlank(url) && StringUtils.isNotBlank(title) && StringUtils.isNotBlank(content)) {
//                articleEntity.setTitle(title);
//                articleEntity.setUrl(url);
//                articleEntity.setContent(content);
//                articleEntities.add(articleEntity);
//            }
//        }
//        return articleEntities;
//    }
//}
