package com.gaolei.excelpoi.persistence.manager.impl;

import com.gaolei.excelpoi.model.SysUser;
import com.gaolei.excelpoi.persistence.dao.SysUserMapper;
import com.gaolei.excelpoi.persistence.manager.SysUserService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

/**
 * @Author: GaoLei
 * @Date: 2019/10/17 17:44
 * @Blog https://blog.csdn.net/m0_37903882
 * @Description:
 */
@Service
public class SysUserServiceImpl implements SysUserService {
    @Autowired
    private SysUserMapper sysUserMapper;

    @Override
    public int deleteByPrimaryKey(String id) {
        return sysUserMapper.deleteByPrimaryKey(id);
    }

    @Override
    public int insert(SysUser record) {
        return sysUserMapper.insert(record);

    }

    @Override
    public int insertSelective(SysUser record) {
        return  sysUserMapper.insertSelective(record);

    }

    @Override
    public SysUser selectByPrimaryKey(String id) {
        return sysUserMapper.selectByPrimaryKey(id);
    }

    @Override
    public int updateByPrimaryKeySelective(SysUser record) {
        return sysUserMapper.updateByPrimaryKeySelective(record);

    }

    @Override
    public int updateByPrimaryKey(SysUser record) {
        return sysUserMapper.updateByPrimaryKey(record);
    }

    @Override
    public List<SysUser> getUserList() {
        return sysUserMapper.getUserList();
    }
}
