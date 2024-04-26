package spring.config;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.servlet.config.annotation.InterceptorRegistry;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

import spring.common.web.LogonCheckInterceptor;

//===================== �߰��� Class  ======================//
// Interceptor ����ϴ� WebMvcCongigurer ���� Bean
//=======================================================//
@Configuration
public class WebConfig implements WebMvcConfigurer {

	@Autowired
	@Qualifier("logonCheckInterceptor")
	private LogonCheckInterceptor logonCheckInterceptor;

	public WebConfig() {
		System.out.println("==> WebConfig default Constructor call.............");
	}

	@Override
	public void addInterceptors(InterceptorRegistry registry) {

		// URL Pattern �� Ȯ���ϰ�. interceptor �������� �����.
		registry.addInterceptor(logonCheckInterceptor).addPathPatterns("/user/*");

	}

}
