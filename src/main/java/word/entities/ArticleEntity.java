package word.entities;

import lombok.Data;
import lombok.ToString;

/**
 * @Author: 朝花夕誓
 * @Date: 2020/11/23 12:00
 * @Version 1.0
 */
@Data
@ToString
public class ArticleEntity {

    private String url;

    private String title;

    private String content;

}
