package crawler.baidu.entity;

/**
 * @author macia
 * @date 2021/1/26 11:55
 */
public enum BaiDuIndexEnum {

    SEARCHINDEX("搜索指数", "http://index.baidu.com/api/SearchApi/index?area=0&word="),

    NEWSINDEX("资讯指数", "http://zhishu.baidu.com/api/FeedSearchApi/getFeedIndex?area=0&word=");

    private String name;

    private String url;

    /**
     * @param name
     * @param url
     */
    BaiDuIndexEnum(String name, String url) {
        this.name = name;
        this.url = url;
    }

    /**
     * @param name
     * @return
     */
    public static String match(String name){
        BaiDuIndexEnum[] values = BaiDuIndexEnum.values();
        for (int i = 0; i < values.length; i++) {
            BaiDuIndexEnum value = values[i];
            if (value.name.equals(name)){
                return value.url;
            }
        }
        return null;
    }

}
