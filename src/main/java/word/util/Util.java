package word.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.sun.istack.internal.NotNull;
import lombok.extern.log4j.Log4j2;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;

/**
 * @author 朽木不可雕也
 * @date 2020/9/17 9:30
 * @week 星期四
 */
@Log4j2
public class Util {

    private Util() {
    }

    
    @NotNull
    public static String getColorString(@NotNull Color color) {
        return getHexString(color.getRed(), color.getGreen(), color.getBlue());
    }

    
    @NotNull
    public static String getHexString(@NotNull Integer... numbers) {
        StringBuilder builder = new StringBuilder();
        for (Integer number : numbers) {
            String numberStr = Integer.toHexString(number);
            //字符串长度不足2位，追加“0”
            if (numberStr.length() < 2) {
                builder.append("0");
            }
            builder.append(numberStr);
        }
        return builder.toString().toUpperCase();
    }

    /**
     * 将长度转换为像素
     *
     * @param length 长度
     * @param unit   单位
     * @return 像素
     */
    public static double getPixel(Double length,  LengthUnit unit) {
        if (unit == LengthUnit.centimeter) {
            //厘米换算为像素
            return length * 567D;
        }
        return 0D;
    }

    
    @NotNull
    public static String translationToString(String text, Language aims) throws Exception {
        Map<String, String> response = translation(text, aims);
        StringBuilder builder = new StringBuilder();
        response.forEach((str, message) -> builder.append(message));
        return builder.toString();
    }

    /**
     * 使用google的翻译接口进行翻译文本
     *
     * @param text 需要翻译的文本
     * @param aims 目标语言
     * @return 翻译得到的文本
     */
    
    @NotNull
    public static Map<String, String> translation(String text, @NotNull Language aims) throws Exception {
        String urlStr = "http://fanyi.youdao.com/translate?&doctype=json&type=" + aims.getUrlValue() + "&i=" + URLEncoder.encode(text, "UTF-8");
        HttpClient client = HttpClients.createDefault();
        HttpGet get = new HttpGet(urlStr);
        HttpResponse response = client.execute(get);

        JSONObject jsonObject = JSON.parseObject(EntityUtils.toString(response.getEntity()));
        if (jsonObject.getIntValue("errorCode") != 0) {
            throw new Exception("翻译出错：" + jsonObject.toJSONString());
        }
        JSONArray array = jsonObject.getJSONArray("translateResult");
        Map<String, String> map = new LinkedHashMap<>();
        for (int i = 0; i < array.size(); i++) {
            JSONArray array1 = array.getJSONArray(i);
            for (int j = 0; j < array1.size(); j++) {
                JSONObject jsonObject1 = array1.getJSONObject(j);
                map.putIfAbsent(jsonObject1.getString("src"), jsonObject1.getString("tgt"));
            }
        }
        return map;
    }

    /**
     * 修改图片的大小
     *
     * @param imagePath 图片的路径
     * @param width     目标宽
     * @param height    目标的高
     * @return 缩放后的图片
     * @throws IOException IO流一场
     */
    
    @NotNull
    @SuppressWarnings("unused")
    public static BufferedImage setImageSize(String imagePath, int width, int height) throws IOException {
        InputStream inputStream = null;
        try {
            inputStream = new FileInputStream(imagePath);
            //读取图片
            Image image = ImageIO.read(inputStream);
            BufferedImage bufferedImage = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
            bufferedImage.getGraphics().drawImage(image, 0, 0, width, height, null);
            return bufferedImage;
        } finally {
            try {
                if (Objects.nonNull(inputStream)) inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 翻译支持的语言
     * ZH_CN2EN 中文　»　英语
     * ZH_CN2JA 中文　»　日语
     * ZH_CN2KR 中文　»　韩语
     * ZH_CN2FR 中文　»　法语
     * ZH_CN2RU 中文　»　俄语
     * ZH_CN2SP 中文　»　西语
     * EN2ZH_CN 英语　»　中文
     * JA2ZH_CN 日语　»　中文
     * KR2ZH_CN 韩语　»　中文
     * FR2ZH_CN 法语　»　中文
     * RU2ZH_CN 俄语　»　中文
     * SP2ZH_CN 西语　»　中文
     */
    @SuppressWarnings("unused")
    public enum Language {
        /**
         * 默认，中文->英文
         */
        AUTO("AUTO"),
        /**
         * 中文->英文
         */
        Chinese_English("ZH_CN2EN"),
        /**
         * 中文->日语
         */
        Chinese_Japanese("ZH_CN2JA"),
        /**
         * 中文->韩语
         */
        Chinese_Korean("ZH_CN2KR"),
        /**
         * 中文->法语
         */
        Chinese_French("ZH_CN2FR"),
        /**
         * 中文->俄语
         */
        Chinese_Russian("ZH_CN2RU"),
        /**
         * 中文->西语
         */
        Chinese_Spanish("ZH_CN2SP"),
        /**
         * 英语->中文
         */
        English_Chinese("EN2ZH_CN"),
        /**
         * 日语->中文
         */
        Japanese_Chinese("JA2ZH_CN"),
        /**
         * 韩语->中文
         */
        Korean_Chinese("KR2ZH_CN"),
        /**
         * 法语->中文
         */
        French_Chinese("FR2ZH_CN"),
        /**
         * 俄语->中文
         */
        Russian_Chinese("RU2ZH_CN"),
        /**
         * 西语->中文
         */
        Spanish_Chinese("SP2ZH_CN");

        private String urlValue;

        Language(String urlValue) {
            this.urlValue = urlValue;
        }

        public String getUrlValue() {
            return urlValue;
        }
    }

    /**
     * 长度单位
     */
    public enum LengthUnit {
        /**
         * 厘米
         */
        centimeter
    }
}
