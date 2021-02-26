package crawler.baidu.parse;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.JSONPath;
import com.squareup.okhttp.OkHttpClient;
import com.squareup.okhttp.Request;
import com.squareup.okhttp.Response;
import crawler.baidu.entity.BaiDuIndexEntity;
import crawler.baidu.entity.BaiDuIndexEnum;
import lombok.extern.log4j.Log4j2;
import org.apache.commons.lang3.time.FastDateFormat;

import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 百度指数获取
 * @author 朝花夕誓
 */
@Log4j2
public class BaiDuIndexParse {

    private static OkHttpClient okHttpClient = new OkHttpClient();
    /**
     * 输入日期格式检验re
     *
     */
    private static String dateRule = "\\d{4}-\\d{1,2}.?\\d{1,2}?";

    /**
     * 输入日期获取年份re
     *
     */
    private static String yearsRe = "^\\d{4}";

    /**
     * 输入时间检验
     * @param startTime
     * @param endTime
     * @return
     */
    private static Boolean timeLimitCheck(String startTime, String endTime){
        try {
            // 最大时间 间隔时间戳
            FastDateFormat fdf = FastDateFormat.getInstance("yyyy-MM-dd");
            long maxIntervalStamp = fdf.parse(maxInterval(startTime)).getTime();
            // 开始时间戳
            long startTimeStamp = fdf.parse(startTime).getTime();
            // 结束时间戳
            long endTimeStamp = fdf.parse(endTime).getTime();

            long timeInterval = endTimeStamp - startTimeStamp;
//            log.info("开始时间：{}， 结束时间：{}， 最大时间间隔：{}", startTime, endTime, maxIntervalStamp);
            if (timeInterval < 0){
                log.error("开始时间不能小于结束时间.");
                return false;
            }
            if (maxIntervalStamp > endTimeStamp){
                return true;
            }else {
                return false;
            }
        } catch (ParseException e) {
            log.info("时间解析异常 ", e);
        }
        return null;

    }

    /**
     *  获取时间戳测试
     *  存在问题
     * @param time 输入事件
     * @return 时间戳
     */
    private static long getTimeStamp(String time){
        System.out.println(time);
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd", Locale.CHINA);
        return LocalDateTime.parse(time, formatter).toEpochSecond(ZoneOffset.of("+8"));
    }

    /**
     * 获取开始时间一年后的时间
     * @param startTime
     * @return
     */
    private static String maxInterval(String startTime){
        Matcher yearMatcher = Pattern.compile(yearsRe).matcher(startTime);
        String before = null;
        int years = 0;
        while (yearMatcher.find()){
            before = yearMatcher.group();
            years = Integer.parseInt(yearMatcher.group());
        }
        String after = String.valueOf(years + 1);
        return startTime.replace(before, after);
    }

    /**
     * 获取以当前时间向前推 days 天数据指数
     * @param keyWord 关键词
     * @param days 天数
     * @return url
     */
    /*private static String getUrl(String keyWord, String days) {
        return url + keyWordEncoder(keyWord) + urlSuff + "&days=" + days;
    }*/

    /**
     * 关键词进行 utf-8 编码
     * @param keyWord
     * @return encodeKeyWord
     */
    private static String keyWordEncoder(String keyWord){
        try {
            return URLEncoder.encode(keyWord, "utf-8");
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取指定天数的 url 链接
     * @param keyWord 关键词
     * @param startDate 起始时间
     * @param endDate 结束时间
     * @return url
     */
    public static String getUrl(String type,String keyWord, String startDate, String endDate){
        if (startDate.matches(dateRule) && endDate.matches(dateRule)){
            String url = BaiDuIndexEnum.match(type);
            if (url != null){
                return url.concat("[[%7B%22name%22:%22") + keyWordEncoder(keyWord) + "%22,%22wordType%22:1%7D]]" + "&startDate=" + startDate + "&endDate=" + endDate;
            }
        }
        return null;
    }

    /**
     * 获取公钥
     * @param html
     * @return publicKey
     */
    private static String getPublicKey(String html){
        String publicKey;
        try {
            publicKey = JSONObject.parseObject(html).getJSONObject("data").getString("uniqid");
            return publicKey;
        } catch (Exception e) {
            log.error("获取公钥失败", e);
        }
        return null;
    }

    /**
     * 获取加密数据 allData
     * @param html
     * @return allData
     */
    private static Map<String, String> getEncryptedData(String html){
        Map<String, String> dataMap = new HashMap<>();
        String allData,startDate,endDate;
        try {
            if (html.contains("userIndexes")){
                JSONArray jsonArray = (JSONArray) JSONPath.read(html, "$.data.userIndexes");
                allData = jsonArray.getJSONObject(0).getJSONObject("all").getString("data");
                startDate = jsonArray.getJSONObject(0).getJSONObject("all").getString("startDate");
                endDate = jsonArray.getJSONObject(0).getJSONObject("all").getString("endDate");
                dataMap.put("data", allData);
                dataMap.put("startDate", startDate);
                dataMap.put("endDate", endDate);
            }else {
                JSONArray jsonArray = (JSONArray) JSONPath.read(html, "$.data.index");
                allData = jsonArray.getJSONObject(0).getString("data");
                startDate = jsonArray.getJSONObject(0).getString("startDate");
                endDate = jsonArray.getJSONObject(0).getString("endDate");
                dataMap.put("data", allData);
                dataMap.put("startDate", startDate);
                dataMap.put("endDate", endDate);
            }
            return dataMap;
        } catch (Exception e) {
            log.info("获取加密数据失败!", e);
        }
        return null;
    }

    /**
     * 数据解密
     * @param key 解密密钥
     * @param data 加密数据
     * @return
     *  测试案例 ：deCode("5aqLbvMtR,m.p+C-8946.71+5%03,2", "baM+t.ta+tp,M+ttat+ttpt+t.Cb+qLp+t.,.+tt.b+CtM,+Ct,.+tLat+taqL+tCpb+t.qL+t.Lb+qLL+q.a+ba.+,tC+Lqb+,b,+,tt+,qq+Lab+tpLt+ba,+L,L+L.C+LC.+L,t+LCt+LtL+,Ca+MCp+,qq+L,C+Lp,+,Ca+bbM+qCq+aqp+M.,+,q,+,tb+,t,+Lq.+Lbb+Laa+Lqa+LLa+Lt.+LtC+Lqp+LqL+bM.+,aa+,,,+MbM+bab+baL+bbq+,aC+,CC+L,a+La,+,.b+Lqb+,a.+L,M+LLC+ppC+ptq+pp,+CqL+pCa+CMa+p.p+p.b+Cq,+ppC+ppq+pLa+pC.+pCL+p.p+p.C+Cqt+ppM+pbp+p,C+pqt+b.C+bM.+ba.+tbpb+tLtM+qLM+a.L+Ma.+,a.+t.,t+bCq+,CL+bC,+b.M+L,L+Lpq+LCM+Ltb+pCt+ppp+pL.+pbM+L,C+Lat+LqC+LpL+L,b+LMM+bpC+M.,+bbL");
     */
    private static List<String> deCode(String key, String data){
        String a = key;
        String i = data;
        Map<Character, Character> n = new HashMap<>();
        List<Character> s = new ArrayList<>();
        String decodeData = "";
        List<String> decodeList;

        for (int j = 0; j < (int)Math.floor(a.length()/2); j++) {
            n.put(a.charAt(j), a.charAt((int)Math.floor(a.length()/2)+j));
            log.info("n key is : {}  , n value is : {}", a.charAt(j), a.charAt((int)Math.floor(a.length()/2)+j));
        }
        for (int k = 0; k < data.length(); k ++){
            s.add(n.get(i.charAt(k)));
        }
        for (Character character : s) {
            decodeData = decodeData.concat(String.valueOf(character));
        }
        decodeList = Arrays.asList(decodeData.split(","));
        log.info("数据大小为：{}", decodeList.size());
        return decodeList;
    }

    /**
     * 获取网页结构
     * @return
     */
    private static String getHtml(Request request){
        for (int i = 0; i < 3; i++) {
            try {
                Response response = okHttpClient.newCall(request).execute();
                return response.body().string();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    private static Request getRequest(String url, Map<String, Object> headerMap){
        Request request = new Request.Builder().url(url).method("GET", null).build();
        headerMap.forEach((key, value)->{
            System.out.println("header map key is : " + key +  "      header map value is : " + String.valueOf(value));
            request.newBuilder().addHeader(key, String.valueOf(value)).build();
        });
        return request;

    }

    /**
     * 获取解密密钥
     * @param publicKey
     * @param cookie
     * @return
     */
    private static String getKey(String publicKey, String cookie){
        Objects.requireNonNull(publicKey, "public key can't null");
        String url = "http://index.baidu.com/Interface/ptbk?uniqid=".concat(publicKey);
        String html = getHtml(getRequest(url, cookie));
        JSONObject jsonObject = JSONObject.parseObject(html);
        String data = jsonObject.getString("data");
        return data;

    }

    /**
     * 头信息初始化
     * @param cookie
     * @return
     */
    private static Request getRequest(String url, String cookie){
        return  new Request.Builder().url(url).method("GET", null)
                .addHeader("Accept", " application/json, text/plain, */*")
                .addHeader("Accept-Language", " zh-CN,zh;q=0.9,en;q=0.8")
                .addHeader("Cache-Control", " no-cache")
                .addHeader("Cookie", cookie)
                .addHeader("Host", " zhishu.baidu.com")
                .addHeader("Pragma", " no-cache")
                .addHeader("Proxy-Connection", " keep-alive")
                .addHeader("Referer", " http://zhishu.baidu.com/v2/main/index.html")
                .addHeader("User-Agent", " Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.190 Safari/537.36")
                .build();
    }

    /**
     * 日期推导
     * 仅用于 传入 days 时间的推导
     * @param dataList
     * @return
     */
    private static Map<String, String> dateDeriveForDays(List<String> dataList){
        long dayStamp = 86400000L;
        Map<String, String> dataMap = new HashMap<>();
        long nowTime = System.currentTimeMillis();
        System.out.println();
        for (int i = 0; i < dataList.size(); i++) {
            String dataTime = new SimpleDateFormat("yyyy年MM月dd日").format((nowTime - dayStamp) - dayStamp * (dataList.size() - i));
            log.info(dataTime , " : ", dataList.get(i));
            dataMap.put(dataTime, dataList.get(i));
        }
        return dataMap;
    }

    public static Map<String, String> dateDeriveForDays(List<String> dataList, String endDate){
        long dayStamp = 86400000L;
        Map<String, String> dataMap = new HashMap<>();
        long endTime = 0;
        try {
            endTime = new SimpleDateFormat("yyyy-MM-dd").parse(endDate).getTime();
        } catch (ParseException e) {
            e.printStackTrace();
        }
        System.out.println("数据集合大小为： " + dataList.size());
        for (int i = 0; i < dataList.size(); i++) {
            String dataTime = new SimpleDateFormat("yyyy年MM月dd日").format(endTime - dayStamp * (dataList.size() - i));
            System.out.println(dataTime + " : " + dataList.get(i));
            dataMap.put(dataTime, dataList.get(i));
        }
        return dataMap;
    }

    /**
     * 指定时间区间的推导
     * @param dataList
     * @param startDate
     * @return
     */
    private static List<BaiDuIndexEntity> dateDeriveForInterval(List<String> dataList, String startDate, String keyWord){
        List<BaiDuIndexEntity> entityList = new ArrayList<>();
        long dayStamp = 86400000L;
        try {
            long startTime = FastDateFormat.getInstance("yyyy-MM-dd").parse(startDate).getTime();
            for (int i = 0; i < dataList.size(); i++) {
                BaiDuIndexEntity baiDuIndexEntity = new BaiDuIndexEntity();
                String dataTime = new SimpleDateFormat("yyyy年MM月dd日").format(startTime + i * dayStamp);
//                log.info(dataTime + " : " + dataList.get(i));
                baiDuIndexEntity.setDate(dataTime);
                baiDuIndexEntity.setIndexNumber(dataList.get(i));
                baiDuIndexEntity.setKeyWord(keyWord);
                entityList.add(baiDuIndexEntity);
            }
            return entityList;
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 检查关键词是否被收录
     * @param keyWord
     * @param cookie
     * @return
     */
    private static Boolean checkKeyWordExists(String keyWord,String cookie){
        String checkUrl = "http://index.baidu.com/api/AddWordApi/checkWordsExists?word=";
        checkUrl = checkUrl + keyWordEncoder(keyWord);
        String html = getHtml(getRequest(checkUrl, cookie));
        if (!html.contains("addWordsNum")){
            return true;
        }
        return false;
    }

    /**
     * 输入天数获取搜索指数
     * @param keyword
     * @param days
     * @param cookie
     * @return dataMap
     */
  /*  private static Map<String, String> getBaiDuIndexData(String keyword, String days, String cookie){
        // 当前时间向前推移 url
        url = getUrl(keyword, days);
        log.info("当前访问的URL: {}", url);
        List<String> decodeDataList = getDecodeDataList(url, cookie);
        Map<String, String> dataMap = dateDeriveForDays(decodeDataList);
        return dataMap;
    }*/

    /**
     * 输入日期区间获取搜索指数
     * @param keyword
     * @param startDate
     * @param endDate
     * @param cookie
     * @return
     */
    public static List<BaiDuIndexEntity> getBaiDuIndexData(String type, String keyword, String startDate, String endDate, String cookie){
        // 对输入时间进行检测
        if (timeLimitCheck(startDate, endDate) == true){
            // 检测关键词是否被收录
            Boolean keyWordExists = checkKeyWordExists(keyword, cookie);
            if (keyWordExists == true){
                String url = getUrl(type, keyword, startDate, endDate);
                log.info("当前访问的 url is : {}", url);
                // 获取网页结构
                String html = getHtml(getRequest(url, cookie));
                // 获取加密数据
                Map<String, String> data = getEncryptedData(html);
//                log.info("secret data is : {}" , data.get("data"));
                // 获取公钥
                System.out.println("===============================================================================================");
                System.out.println("get encrypt data is : " + html);
                System.out.println("===============================================================================================");
                String publicKey = getPublicKey(html);
//                log.info("public key is : {}", publicKey);
                // 获取秘钥
                String key = getKey(publicKey, cookie);
                // 数据解密
                List<String> listData = deCode(key, data.get("data"));
                // 数据封装
                List<BaiDuIndexEntity> entityList = dateDeriveForInterval(listData, data.get("startDate"), keyword);
                return entityList;
            }else {
                log.error("关键词" + keyword +"未被收录，如要查看相关数据，您需要购买创建新词的权限。\n" +
                        "购买创建新词的权限后，您可以添加自己关注的关键词，添加后百度指数系统将在次日更新数据。" +
                        "\n购买链接： http://index.baidu.com/v2/main/index.html#/buy/word");
            }
        }else {
            log.error("输入的时间区间不能大于一年.");
        }

        return null;
    }

    /**
     * 获取解密数据
     * @param url
     * @param cookie
     * @return
     */
    private static List<String> getDecodeDataList(String url, String cookie){
        // 获取网页结构
        String html = getHtml(getRequest(url, cookie));
        // 获取加密数据
        Map<String, String> data = getEncryptedData(html);
        log.info("secret data is : {}", data.get("data"));
        // 获取公钥
        String publicKey = getPublicKey(html);
        log.info("public key is : {}", publicKey);
        // 获取秘钥
        String key = getKey(publicKey, cookie);
        // 数据解密
        List<String> listData = deCode(key, data.get("data"));
        return listData;
    }
}
