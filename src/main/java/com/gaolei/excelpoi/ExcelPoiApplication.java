package com.gaolei.excelpoi;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan(value = "com.gaolei.excelpoi.persistence.dao")
public class ExcelPoiApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExcelPoiApplication.class, args);
    }

}
