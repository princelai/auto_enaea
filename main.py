import argparse
import json
import re
from datetime import timedelta
from pathlib import Path
from time import sleep

from loguru import logger
from rich.live import Live
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TaskID, TimeElapsedColumn
from rich.table import Table
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

pattern = re.compile(r"(\d+)")


def input_user_info():
    user_name = input("请输入用户名:")
    password = input("请输入密码:")
    return user_name, password


class StudyEnaea:
    def __init__(self, user_name, password):
        url = "https://www.enaea.edu.cn/"

        self.USER_NAME = user_name
        self.PASSWORD = password

        self.class_cate = {3: "必修", 4: "选修"}

        self.user_table = Table(title="教育干部网络学院")
        self.user_table.add_column("用户名", style="cyan")
        self.user_table.add_column("姓名", style="magenta")
        self.user_table.add_column("单位", style="green")

        self.overall_progress = Progress(
            "{task.description}",
            SpinnerColumn(),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
            TextColumn("{task.completed}/{task.total}"),
        )
        self.job_progress = Progress(
            "{task.description}",
            SpinnerColumn(),
            BarColumn(),
            TextColumn("[progress.percentage]{task.percentage:>3.1f}%"),
            TimeElapsedColumn(),
            TextColumn("{task.fields[class_left]}"),
        )

        self.progress_table = Table.grid()
        self.progress_table.add_row(self.user_table)
        self.progress_table.add_row(
            Panel.fit(
                self.overall_progress, title="项目列表", border_style="green", padding=(1, 1)
            )
        )
        self.progress_table.add_row(
            Panel.fit(self.job_progress, title="课程任务列表", border_style="red", padding=(1, 1))
        )

        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--window-size=1800x900")

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.maximize_window()
        self.driver.get(url)

    def start(self):
        # 登录
        self.log_in()
        # 初始化课程任务
        self.create_user_table()
        self.create_overall_task()
        # 循环学习
        self.learning()

    def quit(self):
        self.driver.close()
        self.driver.quit()

    def switch_to_tab(self, t):
        sleep(1)
        self.driver.switch_to.window(self.driver.window_handles[t])

    def close_tab(self, t):
        self.switch_to_tab(t)
        self.driver.close()
        self.switch_to_tab(t - 1)

    def wait_render(self, method, locate, err_desc, wait_time=5):
        try:
            element = WebDriverWait(self.driver, wait_time).until(
                EC.presence_of_element_located((method, locate))
            )
        except Exception as e:
            logger.error(f"{err_desc}:{e}")
            self.quit()
            exit(1)
        else:
            return element

    def log_in(self):
        loggin_button = self.wait_render(By.XPATH, '//*[@id="loginAbleSky"]/a[1]', '未能加载主页')
        loggin_button.click()

        _ = self.wait_render(By.ID, "login-wrap", '未能加载登录界面')
        logger.info("正在登录")

        # self.driver.save_screenshot(f"./screenshot/{datetime.now():%Y%m%d%H%M%S}.png")

        input1 = self.driver.find_element_by_xpath('//*[@id="pc-form"]/div[2]/div[1]/div[1]/input')
        input1.clear()
        input1.send_keys(self.USER_NAME)

        input2 = self.driver.find_element_by_xpath('//*[@id="pc-form"]/div[2]/div[1]/div[2]/input')
        input2.clear()
        input2.send_keys(self.PASSWORD)

        self.driver.find_element_by_id('autoLogin').click()

        loggin_button2 = self.driver.find_element_by_xpath('//*[@id="pc-form"]/div[2]/div[1]/div[4]/button')
        loggin_button2.click()

        _ = self.wait_render(By.CLASS_NAME, "u-main-wrap", '登录失败')
        logger.info("登录成功")

    def get_project_info(self, k, get_name=False):
        """
        获取课程项目详情，名称，时间等
        :param k:
        :param get_name:
        :return:
        """
        self.switch_to_tab(1)
        self.driver.refresh()
        _ = self.wait_render(By.ID, "sideNav", '课程方案加载失败')

        if get_name:
            project_title = self.driver.find_element_by_xpath('//*[@id="guide-sub-title"]/a').text.split('•')[1]
            if len(project_title) > 15:
                project_name = f"{project_title[:8]}...{project_title[-6:]}"
            else:
                project_name = project_title
        else:
            project_name = ''

        self.driver.find_element_by_xpath(f'//*[@id="sideNav"]/ul/li[3]/a').click()
        _ = self.wait_render(By.ID, "J_tabsContent", '课程详情页加载失败')
        self.driver.find_element_by_xpath('//*[@id="J_tabsContent"]/div/ul/li[2]/span').click()
        _ = self.wait_render(By.ID, "J_myOptionRecords", '未学课程详情页加载失败')
        class_require = self.driver.find_element_by_xpath('//*[@id="J_tabsContent"]/div/div[1]/div').text

        return project_name, pattern.findall(class_require)

    def create_user_table(self):
        """
        获取并创建用户信息表格
        :return:
        """
        first_page_url = self.driver.current_url
        self.driver.find_element_by_xpath('/html/body/div[2]/div[4]/div[2]/div[2]/p[1]/a[1]').click()
        _ = self.wait_render(By.ID, 'base_view', '用户信息页面加载失败')
        username = self.driver.find_element_by_xpath('//*[@id="base_view"]/div[2]').text
        real_name = self.driver.find_element_by_xpath('//*[@id="base_realName"]').text

        self.driver.find_element_by_xpath('//*[@id="font2"]').click()
        _ = self.wait_render(By.ID, 'work_view', '用户信息页面加载失败')
        work = self.driver.find_element_by_xpath('//*[@id="work_workUnit"]').text
        self.driver.get(first_page_url)
        _ = self.wait_render(By.CLASS_NAME, "u-main-wrap", '返回项目主页失败')
        self.user_table.add_row(username, real_name, work)

    def create_overall_task(self):
        for pl in self.driver.find_elements_by_xpath('//*[@id="J_userGradeList"]/div/div[3]/p[1]/a'):
            pl.click()

            for j in (3, 4):
                project_name, project_progress = self.get_project_info(j, get_name=True)
                self.overall_progress.add_task(f"{project_name} {self.class_cate.get(j)}",
                                               total=int(project_progress[0]), completed=int(project_progress[1]),
                                               project_link=pl, cate=j)
            self.close_tab(1)

    def update_task(self, pt_id, j):
        """
        获取课程项目详情及学习进度，更新全局进度和课程进度
        :param pt_id:
        :param j:
        :return:
        """
        _, project_progress = self.get_project_info(j)
        self.overall_progress.update(pt_id, completed=int(project_progress[1]))
        sleep(1.5)

        class_name = (elem.text[:10] for elem in self.driver.find_elements_by_class_name('course-title'))
        class_progress = [int(elem.text[:-1]) for elem in self.driver.find_elements_by_class_name('progressvalue')]
        class_link = (elem for elem in self.driver.find_elements_by_xpath('//*[@id="J_myOptionRecords"]/tbody/tr/td[6]/a'))
        class_time = (self.parse_time(elem.text) for elem in self.driver.find_elements_by_xpath('//*[@id="J_myOptionRecords"]/tbody/tr/td[2]/span'))
        class_left = (t * (1 - (p / 100)) for t, p in zip(class_time, class_progress))
        class_detail = ([*z] for z in zip(class_name, class_progress, class_link, class_left))

        if self.job_progress.tasks:
            # 追加课程任务
            job_task = {t.description: t.id for t in self.job_progress.tasks}
            for cd in class_detail:
                if cd[0] in job_task:
                    self.job_progress.update(job_task.get(cd[0]), completed=cd[1], class_link=cd[2], class_left=cd[3])
                else:
                    self.job_progress.add_task(cd[0], total=100, completed=cd[1], class_link=cd[2], class_left=cd[3], start=False, visible=False)
        else:
            # 第一次初始化课程任务
            for i, cd in enumerate(class_detail):
                if i == 0:
                    self.job_progress.add_task(cd[0], total=100, completed=cd[1], class_link=cd[2], class_left=cd[3], start=True, visible=True)
                elif i == 1:
                    self.job_progress.add_task(cd[0], total=100, completed=cd[1], class_link=cd[2], class_left=cd[3], start=False, visible=True)
                else:
                    self.job_progress.add_task(cd[0], total=100, completed=cd[1], class_link=cd[2], class_left=cd[3], start=False, visible=False)

    @staticmethod
    def parse_time(t):
        hh, mm, ss = t.split(':')
        return timedelta(hours=int(hh), minutes=int(mm), seconds=int(ss))

    def curr_class_progress(self):
        """
        从视频页面抓取当前课程进度
        :return:
        """
        self.switch_to_tab(2)
        _ = self.wait_render(By.CLASS_NAME, 'cvtb-MCK-CsCt-studyProgress', '课堂页面加载失败')
        elems = self.driver.find_elements_by_class_name('cvtb-MCK-CsCt-studyProgress')
        progress = [float(elem.text[:-1]) for elem in elems]
        return progress

    def switch_section(self):
        """
        切换当前视频分段
        :return:
        """
        progress_list = self.curr_class_progress()
        curr_seq = list(map(lambda x: x < 100, progress_list)).index(True)

        if curr_seq != 0:
            logger.debug(f"切换到第{curr_seq + 1}分段")
            self.driver.find_elements_by_class_name('cvtb-MCK-course-content')[curr_seq].click()

    def handle_option(self):
        """
        处理视频页面不定时弹出的选择题
        :return:
        """
        self.switch_to_tab(2)
        while True:
            try:
                element = self.driver.find_element_by_xpath('/html/body/div[6]/table/tbody/tr[2]/td[2]/div[3]/button')
            except Exception:
                break
            else:
                try:
                    self.driver.find_element_by_xpath('/html/body/div[6]/table/tbody/tr[2]/td[2]/div[2]/div/dl/dd/div[3]/div[1]/p/label/input').click()
                except Exception:
                    pass
                finally:
                    element.click()
                    sleep(1)

    def handle_rest(self):
        """
        处理视频页面每20分钟弹出一次的休息提示
        :return:
        """
        self.switch_to_tab(2)
        try:
            element = self.driver.find_element_by_xpath('//*[@id="rest_tip"]/table/tbody/tr[2]/td[2]/div[3]/button')
        except Exception:
            pass
        else:
            element.click()

    def learning(self):
        with Live(self.progress_table, refresh_per_second=5):
            for project_task in self.overall_progress.tasks:
                if project_task.completed < project_task.total:
                    self.switch_to_tab(0)
                    project_task.fields['project_link'].click()
                    self.update_task(project_task.id, project_task.fields['cate'])

                    while not project_task.finished:
                        curr_class_task = [t for t in self.job_progress.tasks if t.completed < t.total][0]
                        if len(self.driver.window_handles) == 2:
                            self.switch_to_tab(1)
                            curr_class_task.fields['class_link'].click()
                            self.switch_section()

                        self.switch_to_tab(2)
                        while (curr_progress := (sum(progress_list := self.curr_class_progress()) / len(progress_list))) < 100:
                            self.handle_option()
                            self.handle_rest()
                            self.update_task(project_task.id, project_task.fields['cate'])
                            sleep(58)
                        else:
                            self.close_tab(2)
                            self.update_task(project_task.id, project_task.fields['cate'])

                            self.job_progress.update(curr_class_task.id, completed=curr_class_task.total, class_left=timedelta(0))
                            if curr_class_task.id >= 1:
                                self.job_progress.update(TaskID(curr_class_task.id - 1), visible=False)
                            self.job_progress.start_task(TaskID(curr_class_task.id + 1))
                            self.job_progress.update(TaskID(curr_class_task.id + 1), visible=True)
                            self.job_progress.update(TaskID(curr_class_task.id + 2), visible=True)
                    else:
                        self.close_tab(1)
                else:
                    # 项目已完成待处理状态
                    self.overall_progress.update(project_task.id, completed=project_task.completed)
            else:
                logger.info("全部学习完成")
                self.driver.quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(prog='enaea')
    parser.add_argument('-u', '--user', type=str, help='用户名')
    parser.add_argument('-p', '--passwd', type=str, help='密码')
    parser.add_argument('-a', '--auto-login', action='store_true', help='自动登陆')
    ext_args = parser.parse_args()

    try:
        root_dir = Path(__file__).absolute().parent
    except NameError:
        root_dir = Path('.')

    user_info_file = root_dir / 'user_info.json'

    if ext_args.user and ext_args.passwd:
        username = ext_args.user
        passwd = ext_args.passwd
    elif ext_args.auto_login:
        try:
            with open(user_info_file.as_posix(), 'r') as f:
                user_info_dict = json.load(f)
            username = user_info_dict['username']
            passwd = user_info_dict['passwd']
        except Exception:
            username, passwd = input_user_info()
    else:
        username, passwd = input_user_info()

    with open(user_info_file.as_posix(), 'w') as f:
        json.dump({'username': username, 'passwd': passwd}, f)

    study = StudyEnaea(username, passwd)
    study.start()
