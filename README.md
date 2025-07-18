# vuln_tracking

---

本项目旨在自动化跟踪管理历史漏洞，可以根据历史的漏洞数据（需要包含指定数据）来生成漏洞跟踪表。

## 生成效果

这里由于模板数据中没有相关数据，所以会有缺失。

![image-20250718104843578](img/image-20250718104843578.png)

## 漏洞数据整合

将历史报告的报表数据整合到一个excel整合表中，并补充测试时间，表格样式随意但表头需严格按照模板设定，也可以留空；

数据字段：

```
序号
系统
漏洞位置
漏洞名称
漏洞等級
测试时间
漏洞描述漏洞危害修复方案
```

![image-20250718102911685](img/image-20250718102911685.png)

## 生成漏洞跟踪

### 脚本使用

```shell
# 安装依赖
pip3 install -r requirements.txt
# 执行
python3 vuln_tracking.py your_input_file_path  you_output_file_path
```

<img src="img/image-20250718104252388.png" alt="image-20250718104252388" style="zoom: 50%;" />

![image-20250718110356753](img/image-20250718110356753.png)

