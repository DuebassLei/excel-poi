package com.gaolei.excelpoi.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * @Author: GaoLei
 * @Date: 2019/10/16 11:59
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description:
 */
@RequestMapping("sys")
@Controller
public class BaseController {
    @GetMapping("index")
    public String index(){
        return "sys/index";
    }
}
