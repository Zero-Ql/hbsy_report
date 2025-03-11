## 项目简介
技术栈:Python、Requests、APScheduler、SMTP、BeautifulSoup

核心功能

1.自动化登录与数据抓取，模拟浏览器登录,动态解析表单参数,通过Session和Cookie维持会话,避免重复认证，逆向分析系统API,自动获取周报/月报的提交参数(如WID,roleld)。

2.定时任务与配置管理，基于APScheduler实现定时触发(每周六8点提交周报,月末8点提交月报),通过TOML配置文件隔离账号、邮箱等敏感信息,提升安全性。

3.异常处理与邮件通知全局捕获网络请求、数据解析异常,失败时自动发送告警邮件使用smtplib推送HTML模板化通知,包含报告内容、提交时间和系统链接。

4.日志监控
记录任务执行日志(文件+控制台),支持DEBUG级问题追踪。

技术亮点
轻量级设计:代码量控制在500行内,无穴余依赖
逆向工程:突破动态参数反爬策略,精准匹配系统更新逻辑。
用户友好:开箱即用,仅需配置邮箱和账号即可部署。

## 使用说明
1. 填写toml文件中的配置项
2. 在服务器端运行程序（推荐使用screen运行）
