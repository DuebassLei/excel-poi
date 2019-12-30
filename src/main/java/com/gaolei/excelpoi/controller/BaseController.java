package com.gaolei.excelpoi.controller;

import com.gaolei.excelpoi.model.Study;
import com.gaolei.excelpoi.model.SysUser;
import com.gaolei.excelpoi.model.UserInfo;
import com.gaolei.excelpoi.persistence.manager.SysUserService;
import com.gaolei.excelpoi.utils.ResultBean;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
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
import org.springframework.util.Base64Utils;
import org.springframework.web.bind.annotation.*;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;

/**
 * @Author: GaoLei
 * @Date: 2019/10/16 11:59
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description: Excel ,word 导出
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

    /**
     * 导出用户Excel
     * @throws IOException
     * @return
     */
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
    /**
     * 导出用户 Word
     * @throws IOException
     * @return
     */
    @PostMapping("user/doc")
    @ResponseBody
    @ApiOperation(value="导出用户doc", httpMethod = "POST",produces="application/json",notes = "导出用户doc")
    public ResultBean exportDoc() throws  IOException{

        Configuration configuration = new Configuration();
        configuration.setDefaultEncoding("utf-8");
        // 模板存放路径
        configuration.setClassForTemplateLoading(this.getClass(), "/templates/code");
        // 获取模板文件
        Template template = configuration.getTemplate("UserInfo.ftl");
        // 模拟数据
        List<Study> studyList =new ArrayList<>();
        studyList.add(new Study("gaolei","root","1","user","123456"));
        studyList.add(new Study("zhangsan","123","15","admin","12232"));

        Map<String,Object> resultMap = new HashMap<>();
        resultMap.put("studyList",studyList);
        File outFile = new File("UserInfoTest.doc");
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),"UTF-8"));
        try {
            template.process(resultMap,out);
            out.flush();
            out.close();
            return null;
        } catch (TemplateException e) {
            e.printStackTrace();
        }
        return ResultBean.success();
    }

    @PostMapping("user/requireInfo")
    @ResponseBody
    @ApiOperation(value="导出用户确认信息表doc", httpMethod = "POST",produces="application/json",notes = "导出用户确认信息表doc")
    public ResultBean  userRequireInfo() throws  IOException{
        Configuration configuration = new Configuration();
        configuration.setDefaultEncoding("utf-8");
        configuration.setClassForTemplateLoading(this.getClass(), "/templates/code");
        Template template = configuration.getTemplate("need.ftl");
        Map<String , Object> resultMap = new HashMap<>();
        List<UserInfo> userInfoList = new ArrayList<>();
        userInfoList.add(new UserInfo("2019","安全环保处质量安全科2608室","风险研判","9:30","10:30","风险研判","风险研判原型设计","参照甘肃分公司提交的分析研判表，各个二级单位维护自己的风险研判信息，需要一个简单的风险上报流程，各个二级单位可以看到所有的分析研判信息作为一个知识成果共享。","张三","李四"));
        resultMap.put("userInfoList",userInfoList);
        File outFile = new File("userRequireInfo.doc");
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),"UTF-8"));
        try {
            template.process(resultMap,out);
            out.flush();
            out.close();
            return null;
        } catch (TemplateException e) {
            e.printStackTrace();
        }
        return ResultBean.success();
    }
    /**
     * 导出带图片的Word
     * @throws IOException
     * @return
     */
    @PostMapping("user/exportPic")
    @ResponseBody
    @ApiOperation(value="导出带图片的Word", httpMethod = "POST",produces="application/json",notes = "导出带图片的Word")
    public ResultBean exportPic() throws IOException {
        Configuration configuration = new Configuration();
        configuration.setDefaultEncoding("utf-8");
        configuration.setClassForTemplateLoading(this.getClass(), "/templates/code");
        Template template = configuration.getTemplate("userPic.ftl");
        Map<String,Object> map = new HashMap<>();
        map.put("name","gaolei");
        map.put("date","2015-10-12");
        map.put("imgCode",imageToString());
        File outFile = new File("userWithPicture.doc");
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outFile),"UTF-8"));
        try {
            template.process(map,out);
            out.flush();
            out.close();
            return null;
        } catch (TemplateException e) {
            e.printStackTrace();
        }
        return  ResultBean.success();
    }

    /**
     * 图片流转base64字符串
     * */
    public static String imageToString() {
        String imgFile = "E:\\gitee\\excel-poi\\src\\main\\resources\\static\\img\\a.png";
        InputStream in = null;
        byte[] data = null;
        try {
            in = new FileInputStream(imgFile);
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        String imageCodeBase64 =  Base64Utils.encodeToString(data);

        return imageCodeBase64;
    }

    public static void main(String[] args) {
        System.out.println("image:"+imageToString());
    }

}
