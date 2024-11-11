package kai9.libs;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import lombok.Getter;

/**
 * 定数
 */
@Configuration
@Getter
public class kai9properties {

    // WebUsernamePasswordAuthenticationFilterのsuccessfulAuthenticationで参照しているが、application.propertesから@Valueで読み取れず、メモリの管理上？Beanとかに乗せても読み取れないので定数として定義している
    // 秒数(60*60*3で3時間)
    public static int CookieMaxAge2 = 60 * 60 * 3;

    @Value("${CookieMaxAge}")
    private String CookieMaxAge;

    @Bean
    public String CookieMaxAge() {
        return this.CookieMaxAge;
    }
}