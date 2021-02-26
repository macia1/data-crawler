package word.Entity;

import com.alibaba.excel.annotation.ExcelProperty;

public class ExcelEntity {

//    @ExcelProperty("序号")
    @ExcelProperty(index = 0)
    private Integer number;

//    @ExcelProperty("标题")
    @ExcelProperty(index = 1)
    private String title;

//    @ExcelProperty("全文")
    @ExcelProperty(index = 2)
    private String content;

//    @ExcelProperty("韩文标题")
    @ExcelProperty(index = 3)
    private String koreanTitle;

//    @ExcelProperty("韩文摘要")
    @ExcelProperty(index = 4)
    private String koreanSummary;

//    @ExcelProperty("渠道")
    @ExcelProperty(index = 5)
    private String channel;

//    @ExcelProperty("渠道域名")
    @ExcelProperty(index = 6)
    private String channelDomainName;

//    @ExcelProperty("时间")
    @ExcelProperty(index = 7)
    private String time;

//    @ExcelProperty("品牌分类")
    @ExcelProperty(index = 8)
    private String brandCategory;

    public ExcelEntity() {
    }

    public ExcelEntity(Integer number, String title, String content, String koreanTitle, String koreanSummary, String channel, String channelDomainName, String time, String brandCategory) {
        this.number = number;
        this.title = title;
        this.content = content;
        this.koreanTitle = koreanTitle;
        this.koreanSummary = koreanSummary;
        this.channel = channel;
        this.channelDomainName = channelDomainName;
        this.time = time;
        this.brandCategory = brandCategory;
    }

    @Override
    public String toString() {
        return "ExcelEntity{" +
                "number=" + number +
                ", title='" + title + '\'' +
                ", content='" + content + '\'' +
                ", koreanTitle='" + koreanTitle + '\'' +
                ", koreanSummary='" + koreanSummary + '\'' +
                ", channel='" + channel + '\'' +
                ", channelDomainName='" + channelDomainName + '\'' +
                ", time='" + time + '\'' +
                ", brandCategory='" + brandCategory + '\'' +
                '}';
    }

    public Integer getNumber() {
        return number;
    }

    public void setNumber(Integer number) {
        this.number = number;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        this.content = content;
    }

    public String getKoreanTitle() {
        return koreanTitle;
    }

    public void setKoreanTitle(String koreanTitle) {
        this.koreanTitle = koreanTitle;
    }

    public String getKoreanSummary() {
        return koreanSummary;
    }

    public void setKoreanSummary(String koreanSummary) {
        this.koreanSummary = koreanSummary;
    }

    public String getChannel() {
        return channel;
    }

    public void setChannel(String channel) {
        this.channel = channel;
    }

    public String getChannelDomainName() {
        return channelDomainName;
    }

    public void setChannelDomainName(String channelDomainName) {
        this.channelDomainName = channelDomainName;
    }

    public String getTime() {
        return time;
    }

    public void setTime(String time) {
        this.time = time;
    }

    public String getBrandCategory() {
        return brandCategory;
    }

    public void setBrandCategory(String brandCategory) {
        this.brandCategory = brandCategory;
    }
}
