import json
import time

import tomli
import http.cookiejar as cookielib
from datetime import datetime
import requests
from apscheduler.schedulers.background import BackgroundScheduler
from bs4 import BeautifulSoup
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from apscheduler.triggers.cron import CronTrigger
import logging

# 配置日志
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# 创建一个文件处理器，将日志写入文件
file_handler = logging.FileHandler('report.log')
file_handler.setLevel(logging.DEBUG)  # 设置文件处理器的日志级别

# 创建一个控制台处理器，将日志输出到控制台
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)  # 设置控制台处理器的日志级别

# 创建一个格式器并将其添加到处理器
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# 将处理器添加到logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)


class ReportUploader:
    def __init__(self, username, password, email_config):
        self.email_config = email_config
        self.username = username
        self.password = password
        self.session = requests.Session()
        self.session.cookies = cookielib.LWPCookieJar(filename="hbsyCookies.txt")
        self.base_headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                          "Chrome/128.0.0.0 Safari/537.36",
            "Accept-Language": "zh-CN,zh;q=0.9",
        }
        self.setJwCommonAppRole_do_cookies = None
        self.currentUser_do_cookies = None
        self.menus_do_cookies = None
        self.current_user = None
        self.role_id = None
        self.plan_id = None
        self.size = None
        self.wid = None

    @staticmethod
    def _handle_request_exception(operation):
        # 修改为静态方法，使其能够正确地作为装饰器使用
        def decorator(func):
            def wrapper(self, *args, **kwargs):
                try:
                    return func(self, *args, **kwargs)
                except requests.exceptions.RequestException as e:
                    raise RuntimeError(f"{operation}请求失败: {str(e)}")
                except (KeyError, IndexError) as e:
                    raise RuntimeError(f"{operation}数据处理错误: {str(e)}")
                except json.JSONDecodeError:
                    raise RuntimeError(f"{operation}响应JSON解析失败")

            return wrapper

        return decorator

    def _load_cookies(self):
        try:
            self.session.cookies.load(ignore_discard=True)
            print("成功加载cookies")
            return True
        except FileNotFoundError:
            print("未找到cookie文件")
            return False
        except Exception as e:
            print(f"加载cookies失败: {str(e)}")
            return False

    def _save_cookies(self):
        try:
            self.session.cookies.save(ignore_discard=True)
            logger.info("成功保存cookies")
        except Exception as e:
            logger.error(f"保存cookies失败: {str(e)}")

    @_handle_request_exception("登录")
    def login(self):
        if self._load_cookies():
            if self._validate_cookies():
                logger.info("使用已有cookies跳过登录")
                return True

        login_url = "http://ids.hbsy.cn/authserver/login?service=http%3A%2F%2Fjwxt.hbsy.cn%2Fjwapp%2Fsys%2Fhomeapp%2Fhome%2Findex.html"

        # 获取登录页面并提取表单参数
        response = self.session.get(login_url, headers=self.base_headers, timeout=10)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")
        form_data = {
            "username": self.username,
            "password": self.password,
            "submit": "",
            "lt": soup.find("input", {"name": "lt"})["value"],
            "dllt": soup.find("input", {"name": "dllt"})["value"],
            "execution": soup.find("input", {"name": "execution"})["value"],
            "_eventId": soup.find("input", {"name": "_eventId"})["value"],
            "rmShown": soup.find("input", {"name": "rmShown"})["value"],
        }

        # 提交登录表单
        response = self.session.post(
            login_url,
            data=form_data,
            headers={**self.base_headers, "Referer": login_url},
            allow_redirects=False
        )

        response.raise_for_status()

        # 处理登录后重定向
        if "Location" in response.headers:
            redirect_url = response.headers["Location"]
            response = self.session.get(
                redirect_url,
                headers=self.base_headers,
                allow_redirects=True
            )
            response.raise_for_status()

        # 验证登录是否成功
        if self._validate_cookies():
            self._save_cookies()
            logger.info("登录成功")
            return True
        raise RuntimeError("登录失败：无效的凭据或响应结构变化")

    def _validate_cookies(self):
        test2_url = "http://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/currentUser.do"
        response = self.session.get(test2_url, headers=self.base_headers)
        if response.status_code == 200:
            try:
                data = response.json()
                if data.get("code") == "0" and data.get("datas"):
                    self.current_user = data["datas"]
                    return True
            except json.JSONDecodeError:
                pass
        return False

    @_handle_request_exception("获取用户角色")
    def _get_role_id(self):
        if not self.current_user:
            self._validate_cookies()

        user_groups = self.current_user.get("userGroups", [])
        if not user_groups:
            raise ValueError("用户角色信息缺失")

        self.role_id = user_groups[0].get("roleId")
        if not self.role_id:
            raise ValueError("未找到有效角色ID")
        return self.role_id

    @_handle_request_exception("获取周报WID")
    def _get_wid(self):
        url = "https://jwxt.hbsy.cn/jwapp/sys/xsdgsxbm/modules/xssxgl/cxxssxxx.do"
        payload = {"SFSY": "1", "*order": "-XNXQDM"}

        response = self.session.post(url, data=payload, headers=self.base_headers,
                                     cookies=self.setJwCommonAppRole_do_cookies)
        data = response.json()

        try:
            self.wid = data["datas"]["cxxssxxx"]["rows"][0]["WID"]
            return self.wid
        except (KeyError, IndexError):
            raise ValueError("响应中未找到有效的WID")

    @_handle_request_exception("获取当前报告标题")
    def _get_bt(self, lx):
        numerals = {"0": "零", "1": "一", "2": "二", "3": "三", "4": "四", "5": "五", "6": "六", "7": "七", "8": "八",
                    "9": "九"}
        url = "https://jwxt.hbsy.cn/jwapp/sys/xsdgsxbm/modules/xssxgl/cxxssxbg.do"
        payload = {"JHXSWID": self.wid, "LX": lx}

        response = self.session.post(url, data=payload, headers=self.base_headers,
                                     cookies=self.setJwCommonAppRole_do_cookies)
        data = response.json()

        try:
            self.size = data["datas"]["cxxssxbg"]["totalSize"]
            title = int(self.size) + 1
            if lx == "zb":
                return f"第{title}周实习周报"
            else:
                return f"{numerals[str(title)]}月月报"

        except (KeyError, IndexError):
            raise ValueError("响应中未找到有效的totalSize")

    @_handle_request_exception("获取PLANID")
    def _get_planid(self):
        ggxxpz_do = self.session.get("https://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/ggxxpz.do?userType=student",
                                     data={'userType': 'student'}, cookies=self.menus_do_cookies)
        announcement_do = self.session.get(
            "https://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/announcement.do?userType=student",
            data={'userType': 'student'}, cookies=ggxxpz_do.cookies)

        educational_program_do = self.session.get(
            "https://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/student/educational-program.do",
            cookies=announcement_do.cookies)
        data = json.loads(educational_program_do.text)

        try:
            self.plan_id = data["datas"][0]["planId"]
            return self.plan_id
        except (KeyError, IndexError):
            raise ValueError("响应中未找到有效的PLANID")

    @_handle_request_exception("获取最新Cookie")
    def _update_cookies(self):
        menus_do = self.session.get("https://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/menus.do?userType=student",
                                    data={'userType': 'student'})
        self.menus_do_cookies = menus_do.cookies
        defaultDisplayConfig_do = self.session.post(
            "https://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/config/defaultDisplayConfig.do", data={
                'typeCode': 'student',
                'module': 'JXRC'
            }, cookies=self.menus_do_cookies)
        current_date1 = datetime.now().strftime("%Y-%m-%d")
        classWeek_do = self.session.get(
            f"https://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/teachingSchedule/classWeek.do?rq={current_date1}",
            data={
                'rq': current_date1,
            }, cookies=defaultDisplayConfig_do.cookies)
        current_date2 = datetime.now().strftime("%Y-%m")
        scores_do = self.session.get(
            f"https://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/student/scores.do?termCode=2024-{current_date2}",
            data={
                'termCode': f"2024-{current_date2}",
            }, cookies=classWeek_do.cookies)
        planId = self._get_planid()
        academic_status_do = self.session.get(
            f"https://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/student/academic-status.do?planId={planId}",
            data={
                'planId': planId,
            }, cookies=scores_do.cookies)
        detail_do = self.session.get(
            f"https://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/teachingSchedule/detail.do?rq={current_date1}&lxdm=student",
            data={
                'rq': current_date1,
                'lxdm': 'student'
            }, cookies=academic_status_do.cookies)
        self._get_role_id()
        index_do = self.session.get(
            f"https://jwxt.hbsy.cn/jwapp/sys/xsdgsxbm/*default/index.do?THEME=indigo&EMAP_LANG=zh&forceApp=xsdgsxbm"
            f"&_yhz=00000{self.role_id}&min=1",
            data={
                'THEME': 'indigo',
                'EMAP_LANG': 'zh',
                'forceApp': 'xsdgsxbm',
                '_yhz': f"00000{self.role_id}",
                'min': '1',
            }, cookies=detail_do.cookies)
        messages_pc_do = self.session.get(
            "https://jwxt.hbsy.cn/jwapp/sys/homeapp/api/home/messages_pc.do?userType=student",
            data={'userType': 'student'}, cookies=index_do.cookies)
        xsdgsxbm_do = self.session.get("https://jwxt.hbsy.cn/jwapp/sys/emappagelog/config/xsdgsxbm.do",
                                       cookies=messages_pc_do.cookies)
        setJwCommonAppRole_do = self.session.post(
            "https://jwxt.hbsy.cn/jwapp/sys/jwpubapp/pub/setJwCommonAppRole.do",
            data={'ROLEID': self.role_id},
            cookies=xsdgsxbm_do.cookies)
        self.setJwCommonAppRole_do_cookies = setJwCommonAppRole_do.cookies

    @_handle_request_exception("提交周报 and 月报")
    def submit_report(self, database):
        # 确保已获取必要参数
        if not all([self.setJwCommonAppRole_do_cookies, self.wid]):
            self._update_cookies()
            self._get_wid()

        url = "https://jwxt.hbsy.cn/jwapp/sys/xsdgsxbm/modules/xssxgl/bcxssxbg.do"
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lx = database["type"]
        title = self._get_bt(lx)

        payload = {
            "param": json.dumps([{
                "LX": lx,
                "BT": f"{title}",
                "XQ": database["report"][self.size]["content"],
                "TJSJ": current_time,
                "PYZT": "wpy",
                "JHXSWID": self.wid
            }])
        }

        response = self.session.post(url, data=payload, headers=self.base_headers,
                                     cookies=self.setJwCommonAppRole_do_cookies)
        response.raise_for_status()
        result = response.json()
        if result.get("code") == "0":
            logger.info(f"{title}提交成功")
            self.email_config.push(
                sitename=f"{title}",
                content=database["report"][self.size]["content"],
                url="http://ids.hbsy.cn/authserver/login?service=http%3A%2F%2Fjwxt.hbsy.cn%2Fjwapp%2Fsys%2Fhomeapp%2Fhome%2Findex.html"
            )
            return True
        raise RuntimeError(f"{title}提交失败: {result.get('msg', '未知错误')}")


class EmailPush:
    def __init__(self, sender, password, receiver):
        self.password = password
        self.receiver = receiver
        self.sender = sender

    def fill_template(self, sitename, content, url):
        # 邮件主题和内容
        subject = 'aminoac'
        body = f"""<html>
        <head>
        <style>
            .wrap span {{
            display: inline-block;
            }}
            .w260 {{
            width: 260px;
            }}
            .w20 {{
            width: 20px;
            }}
            .wauto {{
            width: auto;
            }}
        </style>
        </head>
        <body>
        <table style="width: 99.8%;height:99.8% ">
            <tbody>
            <tr>
                <td>
                <div
                    style="border-radius: 10px 10px 10px 10px;font-size:13px;    color: #555555;width: 666px;font-family:'Century Gothic','Trebuchet MS','Hiragino Sans GB',微软雅黑,'Microsoft Yahei',Tahoma,Helvetica,Arial,'SimSun',sans-serif;margin:50px auto;border:1px solid #eee;max-width:100%;background: #ffffff repeating-linear-gradient(-45deg,#fff,#fff 1.125rem,transparent 1.125rem,transparent 2.25rem);box-shadow: 0 1px 5px rgba(0, 0, 0, 0.15);">
                    <div
                    style="width:100%;background:#49BDAD;color:#ffffff;border-radius: 10px 10px 0 0;background-image: -moz-linear-gradient(0deg, rgb(67, 198, 184), rgb(255, 209, 244));background-image: -webkit-linear-gradient(0deg, rgb(67, 198, 184), rgb(255, 209, 244));height: 66px;">
                    <p
                        style="font-size:15px;word-break:break-all;padding: 23px 32px;margin:0;background-color: hsla(0,0%,100%,.4);border-radius: 10px 10px 0 0;">
                        您的<a style="text-decoration:none;color: #ffffff;" href="<%=siteUrl%>">
                        {sitename}
                        </a>推送成功啦！ </p>
                    </div>
                    <div style="margin:40px auto;width:90%">
                    <p>
                        以下为推送内容：
                    </p>
                    <div
                        style="background: #fafafa repeating-linear-gradient(-45deg,#fff,#fff 1.125rem,transparent 1.125rem,transparent 2.25rem);box-shadow: 0 2px 5px rgba(0, 0, 0, 0.15);margin:20px 0px;padding:15px;border-radius:5px;font-size:14px;color:#555555;">
                        {content}</div>
                    <p><a style="text-decoration:none; color:#12addb" href="{url}" target="_blank">[官网查看]</a></p>
                    <style type="text/css">
                        a:link {{
                        text-decoration: none
                        }}
                        a:visited {{
                        text-decoration: none
                        }}
                        a:hover {{
                        text-decoration: none
                        }}
                        a:active {{
                        text-decoration: none
                        }}
                    </style>
                    </div>
                </div>
                </td>
            </tr>
            </tbody>
        </table>
        </body>
        </html>"""

        msg = MIMEText(body, 'html', 'utf-8')
        msg['From'] = Header(self.sender)
        msg['To'] = Header(self.receiver)
        msg['Subject'] = Header(subject)

        return msg

    def push(self, sitename, content, url):
        try:
            smtpObj = smtplib.SMTP_SSL('smtp.qq.com', 465)
            smtpObj.login(self.sender, self.password)
            smtpObj.sendmail(self.sender, [self.receiver], self.fill_template(sitename, content, url).as_string())
            logger.info(f"邮件发送成功给{self.receiver}")
        except smtplib.SMTPException as e:
            logger.error(f"Error: 无法发送邮件给{self.sender}, 错误信息：{e}")
        finally:
            smtpObj.quit()


class Main:
    def __init__(self, config):
        self.username = config["user_config"]["username"]
        self.password = config["user_config"]["password"]
        self.email_sender = config["user_config"]["email_config"]["sender"]
        self.email_password = config["user_config"]["email_config"]["password"]
        self.email_receiver = config["user_config"]["email_config"]["receiver"]

    def main(self, database):
        try:
            uploader = ReportUploader(
                username=self.username,
                password=self.password,
                email_config=EmailPush(
                    # 发送方的邮箱及授权码（如无特殊情况无需改动）
                    sender=self.email_sender,
                    password=self.email_password,
                    # 收件人邮箱
                    receiver=self.email_receiver
                )
            )
            if uploader.login():
                uploader.submit_report(database)

        except Exception as e:
            logger.error(f"程序运行出错: {str(e)}")
            EmailPush(sender=self.email_sender, password=self.email_password, receiver=self.email_receiver).push(
                sitename=f"报告推送出错",
                content=f"程序运行出错，请前往控制台查看日志。Error: {str(e)}",
                url="http://ids.hbsy.cn/authserver/login?service=http%3A%2F%2Fjwxt.hbsy.cn%2Fjwapp%2Fsys%2Fhomeapp%2Fhome%2Findex.html"
            )


if __name__ == "__main__":
    scheduler = BackgroundScheduler()

    with open("config.toml", "rb") as f:
        config = tomli.load(f)

    # 每周天上午8点整提交周报
    scheduler.add_job(
        Main(config).main,
        CronTrigger.from_crontab("0 8 * * 6"),
        args=[config["database"][0]],
        name="周报"
    )

    # 每月末上午8点整提交月报
    scheduler.add_job(
        Main(config).main,
        CronTrigger(day="last", hour=8, minute=0),
        args=[config["database"][1]],
        name="月报"
    )

    try:
        logger.info("定时任务已启动，按 Ctrl+C 退出")
        scheduler.start()
        for job in scheduler.get_jobs():
            logger.info(f"任务名称: {job.name}, 下次执行时间: {job.next_run_time}")
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("收到终止信号，正在停止调度器...")
    finally:
        scheduler.shutdown(wait=False)
