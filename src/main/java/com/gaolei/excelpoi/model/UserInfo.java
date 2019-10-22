package com.gaolei.excelpoi.model;

import lombok.Data;
import lombok.Getter;
import lombok.Setter;

/**
 * @Author: GaoLei
 * @Date: 2019/10/22 10:54
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description: 需求确认表
 */
@Data
@Getter
@Setter
public class UserInfo {
    private  String date;
    private  String address;
    private  String name;
    private  String startDate;
    private  String endDate;
    private  String taskInfo1;
    private  String taskInfo2;
    private  String taskInfo3;
    private  String user1;
    private  String user2;
    public UserInfo(String date, String address, String name, String startDate, String endDate, String taskInfo1, String taskInfo2, String taskInfo3, String user1, String user2) {
        this.date = date;
        this.address = address;
        this.name = name;
        this.startDate = startDate;
        this.endDate = endDate;
        this.taskInfo1 = taskInfo1;
        this.taskInfo2 = taskInfo2;
        this.taskInfo3 = taskInfo3;
        this.user1 = user1;
        this.user2 = user2;
    }
}
