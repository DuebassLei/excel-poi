package com.gaolei.excelpoi.persistence.manager;

import com.gaolei.excelpoi.model.SysUser;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * @Author: GaoLei
 * @Date: 2019/10/17 17:43
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description:
 */
public interface SysUserService {
    int deleteByPrimaryKey(String id);

    int insert(SysUser record);

    int insertSelective(SysUser record);

    SysUser selectByPrimaryKey(String id);

    int updateByPrimaryKeySelective(SysUser record);

    int updateByPrimaryKey(SysUser record);

    List<SysUser> getUserList();
}
