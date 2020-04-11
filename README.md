# KMS激活Windows

---

## 一、服务器端

需要有属于自己的vps服务器，如果没有，使用虚拟机装一个linux系统也可以。

1.下载vlmcsd

     wget https://github.com/Wind4/vlmcsd/releases/download/svn1111/binaries.tar.gz

2.解压

    tar -zxvf svn1113.tar.gz

3.查看服务器cpu信息

    cat /proc/cpuinfo

4.根据自己的cpu类型切换到相应目录(我这里是intel的cpu)

    cd binaries/Linux/intel/static

5.赋予权限

    chmod 777 vlmcsd-x64-musl-static

6.运行脚本

    ./vlmcsd-x64-musl-static

## 二、Windows端

1.使用管理员权限打开cmd

2.查看windows版本

    wmic os get caption

3.根据windows版本导入相应的密钥

    slmgr /ipk xxxxx-xxxxx-xxxxx-xxxxx

3.配置kms服务器地址（即上面服务器的ip，可以使用ifconfig来查询）

    slmgr /skms xxx.xxx.xxx.xxx

4.手动激活

    slmgr /ato

5.查看激活状态

    slmgr /xpr

到此，Windows激活完成！

# KMS激活Offices

---

## 一、服务器端

同上，已经配置好了不需要重复配置

## 二、Windows端

激活Office，需要使用使用office安装目录下的ospp.vbs工具进行操作

1.切换到 ospp.vbs 的目录，ospp.vbs工具在你的Office的安装目录下，需要自己在计算机找

    cd C:\Program Files(x86)\Microsoft Office\Office16

2.导入相应版本的key（这里以Office专业增强版2019为例）

    cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP

3.配置kms服务器地址（即上面服务器的ip，可以使用ifconfig来查询）

```
slmgr /skms xxx.xxx.xxx.xxx
```

4.手动激活

```
cscript ospp.vbs /act
```

5.查看激活状态

```
cscript ospp.vbs /dstatus
```

到此，Office激活完成！
