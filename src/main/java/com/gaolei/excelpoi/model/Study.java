package com.gaolei.excelpoi.model;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;

/**
 * @Author: GaoLei
 * @Date: 2019/10/21 19:25
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description:
 */
@Data
@Getter
@Setter

public class Study {
    private String name;
    private String pwd;
    private String id;
    private String code;
    private String tel;
    public Study(String name, String pwd, String id, String code, String tel) {
        this.name = name;
        this.pwd = pwd;
        this.id = id;
        this.code = code;
        this.tel = tel;
    }
}
