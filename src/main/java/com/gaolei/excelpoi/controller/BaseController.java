package com.gaolei.excelpoi.controller;

import com.gaolei.excelpoi.model.SysUser;
import com.gaolei.excelpoi.persistence.manager.SysUserService;
import com.gaolei.excelpoi.utils.ResultBean;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;

import java.util.*;

/**
 * @Author: GaoLei
 * @Date: 2019/10/16 11:59
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description:
 */
@Controller
@RequestMapping("sys/v1")
@Api(description = "系统用户")
public class BaseController {
    public static Logger log = LoggerFactory.getLogger(BaseController.class);
    @Autowired
    private SysUserService sysUserService;
    @GetMapping("index")
    public String index(){
        return "indexBase";
    }

    @GetMapping("user")
    @ResponseBody
    @ApiOperation(value="查询用户", httpMethod = "GET", notes = "查询用户")
    public ResultBean  getUser(@RequestParam("id") String id){
        SysUser user =  sysUserService.selectByPrimaryKey(id);
        return ResultBean.success(user);
    }

    @PostMapping("user/export")
    @ResponseBody
    @ApiOperation(value="导出用户", httpMethod = "POST",produces="application/json",notes = "导出用户")
    public ResultBean  exportUser(HttpServletResponse response) throws IOException{
        List<SysUser> userList =  sysUserService.getUserList(); // 获取用户数据
        Map<String, String> fieldMap = new LinkedHashMap<String, String>(); // 数据列信息
    	fieldMap.put("id", "编号");
     	fieldMap.put("name", "姓名");
     	fieldMap.put("pwd", "密码");
     	fieldMap.put("tel", "电话");
     	fieldMap.put("code", "编码");
     	fieldMap.put("comment", "备注");
        XSSFWorkbook workbook = new XSSFWorkbook(); // 新建工作簿对象
        XSSFSheet sheet = workbook.createSheet("UserList");// 创建sheet
        int rowNum = 0;
        Row row =  sheet.createRow(rowNum);// 创建第一行对象,设置表标题
        Cell cell;
        int cellNum = 0;
        for (String name:fieldMap.values()){
            cell = row.createCell(cellNum);
            cell.setCellValue(name);
            cellNum++;
        }
        int rows = 1;
         for (SysUser user: userList){//遍历数据插入excel中
            row = sheet.createRow(rows);
            int col = 0;
            row.createCell(col).setCellValue(user.getId()); // 编号id
            row.createCell(col+1).setCellValue(user.getName()); // 姓名Name
            row.createCell(col+2).setCellValue(user.getPwd()); // 密码pwd
            row.createCell(col+3).setCellValue(user.getTel()); // 电话tel
            row.createCell(col+4).setCellValue(user.getCode()); // 编码
            row.createCell(col+5).setCellValue(user.getComment()); // 备注comment
            rows++;
        }
        String fileName = "userInfo";
        OutputStream out =null;
        try {
            out = response.getOutputStream();
            response.reset();
            response.addHeader("Content-Disposition", "attachment; filename=" + fileName + ".xlsx");
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            out.flush();
            out.close();
        }
        return ResultBean.success();
    }
}
