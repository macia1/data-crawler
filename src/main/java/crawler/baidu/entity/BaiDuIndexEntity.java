package crawler.baidu.entity;

import com.alibaba.excel.annotation.ExcelProperty;

/**
 * 百度实体类
 */
public class BaiDuIndexEntity {

    /**
     * 关键词
     */
    @ExcelProperty("关键词")
    private String keyWord;

    /**
     * 日期
     */
    @ExcelProperty("日期")
    private String date;

    /**
     * 搜索指数
     */
    @ExcelProperty("指数")
    private String indexNumber;

    @Override
    public String toString() {
        return "BaiDuIndexEntity{" +
                "keyWord='" + keyWord + '\'' +
                ", date='" + date + '\'' +
                ", indexNumber='" + indexNumber + '\'' +
                '}';
    }

    public String getKeyWord() {
        return keyWord;
    }

    public void setKeyWord(String keyWord) {
        this.keyWord = keyWord;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getIndexNumber() {
        return indexNumber;
    }

    public void setIndexNumber(String indexNumber) {
        this.indexNumber = indexNumber;
    }
}
