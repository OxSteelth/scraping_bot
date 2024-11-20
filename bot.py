import pyperclip
import selenium
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions, wait
from PyQt5.QtCore import *
from datetime import datetime
from time import sleep
import time
from random import randrange, randint

start_time = time.time()


class Bot:
    running = False
    log_signal = None

    def __init__(self, interval, login_url, email, password, outlet, target_url, date_from, date_to):
        # Login
        self.login_url = login_url
        self.email = email
        self.password = password
        self.outlet = outlet

        # Url and date range of the website you want to automatically access
        self.target_url = target_url
        self.date_from = date_from
        self.date_to = date_to

        # other
        self.interval = interval

    def is_settings_valid(self):
        if not self.login_url or not self.email or not self.password or not self.outlet or not self.target_url or not self.date_from or not self.date_to:
            return False
        return True

    def log(self, msg):
        now_time = time.time()
        elapsed_min = int((now_time - start_time) / 60)
        dt_str = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        output = "{}: [{} min(s)]  {}".format(dt_str, elapsed_min, msg)

        print(output)
        # write to file
        log_file = open("log.txt", "a", encoding='utf-8')
        log_file.write("{}\n".format(output))
        log_file.close()
        # generate signal
        self.log_signal.emit(output)

    def login(self):
        """Login to platform"""
        self.driver.find_element('name', 'account').send_keys(self.email)
        self.driver.find_element('name', 'password').send_keys(self.password)
        self.driver.find_element('xpath', '//*[@id="cForm_row_system_signIn_outlet"]/div/div[1]').click()
        self.driver.find_element('xpath', '//*[@id="cForm_row_system_signIn_outlet"]/div/div[2]/div[1]/input').send_keys(self.outlet)
        self.driver.find_element('xpath', '//*[@id="cForm_row_system_signIn_outlet"]/div/div[2]/div[2]/ul/li').click()
        self.driver.find_element('xpath', '//*[@id="cForm_action_system_signIn"]/button[1]').click()

    def select_collection(self):
        """Select collection to extract data"""
        while True:
            try:
                self.driver.get(self.target_url)
                sleep(3)
                break
            except:
                continue

    def open_report(self):
        """Open report to extract data"""
        self.driver.find_element('id', 'cForm_reporter_reportForm_dateFrom').send_keys(self.date_from)
        self.driver.find_element('id', 'cForm_reporter_reportForm_dateTo').send_keys(self.date_to)
        self.driver.find_element('id', 'siteTitle').click()
        self.driver.find_element('xpath', '//*[@id="cForm_row_reporter_reportForm_outlet"]/ul/li[1]').click()
        self.driver.find_element('xpath', '//*[@id="cForm_action_reporter_reportForm"]/button[1]').click()

    def close_student_hubs_dialog(self):
        """Close Student Hubs Dialog"""
        try:
            self.driver.find_element(
                'xpath',
                '//*[@id="app-mount"]/button[@type="button"][@aria-label="Close"]'
            ).click()
            return True
        except:
            self.log('Student Hubs dialog did not appear')
            return False

    def close_dialog(self):
        """Close Dialog"""
        try:
            self.driver.find_element(
                'xpath',
                '//div[@role="button"][@aria-label="Close"][@tabindex="0"]'
            ).click()
            self.log('Closed dialog.')
            return True
        except Exception as e:
            self.log('Close dialog did not appear')
            return False

    def select_channel(self, server_id, channel_id):
        """Select channel to send message"""
        while True:
            try:
                url = "https://discord.com/channels/" + server_id + "/" + channel_id
                self.driver.get(url)
                sleep(3)
                break
            except:
                continue

    def check_too_many(self):
        text = "You are sending too many new direct messages."
        if text in self.driver.page_source:
            self.log("Too many messages: sleeping for 2.5 mins")
            sleep(150)

    def extract(self, m_file):
        """Extract data one by one"""
        # Driver of the browser you use
        self.driver = webdriver.Chrome("chromedriver.exe")
        self.browse = self.driver.get(self.login_url)
        # start log
        self.log("Bot started ...")
        # Login to platform
        self.login()
        sleep(5)

        # Select collection
        self.select_collection()

        # Open report
        self.open_report()

        # wait until tables are loaded
        WebDriverWait(self.driver, 60).until(expected_conditions.presence_of_element_located((By.XPATH, '//*[@id="siteContent"]/form[@class="cTable"]')))
        sleep(1)

        # find target form
        form = self.driver.find_element('xpath', '//*[@id="siteContent"]/form[3]')

        # retrieve the fields
        fields = []
        columns = form.find_elements('xpath', './/table/thead/tr/th')
        for column in columns:
            try:
                fields.append(column.find_element('xpath', './/a').text.strip())
            except:
                pass
        fields.append('Customer')
        fields.append('Invoice')

        # extract data
        data = []
        rows = form.find_elements('xpath', './/table/tbody/tr')
        for row in rows:
            cells = row.find_elements('xpath', './/td')

            if len(fields) != len(cells):
                print(len(fields), len(cells))
                return

            datum = {}
            i = 0

            for cell in cells:
                try:
                    datum[fields[i]] = cell.find_element('xpath', './/div').text
                except:
                    try:
                        datum[fields[i]] = cell.find_element('xpath', './/a').get_attribute('href')
                    except:
                        datum[fields[i]] = ''

                data.append(datum)
                i += 1

        # remove last row where total amounts are put
        data.pop()
        print(data)

        # close driver
        self.driver.close()

        return
        act = ActionChains(self.driver)

        # aside_tag = self.driver.find_element('xpath', '//aside')
        # aside_height = aside_tag.size['height']
        # log('aside height', aside_height)

        # list_container = self.driver.find_element('xpath', '//aside/div/div')
        # list_height = list_container.size["height"]
        # log('height:', list_height)

        # some settings
        FINISH_COUNT = 300

        visited_members = [] # list of already visited memebrs
        for line in m_file:
            visited_members.append(line.rstrip())
        self.log(visited_members)

        previous_time = 0  # previous sent time
        total_sent = 0
        finish_counter = FINISH_COUNT
        scroll_downed = False  # if scroll downed?

        while self.running:
            # select channel first
            if not scroll_downed:  # if scroll downed, do not refresh channel
                self.select_channel(self.server_id, self.channel_id)
            scroll_downed = False
            self.log("[Finding members]")
            try:
                timestamp = WebDriverWait(self.driver, 60).until(
                    expected_conditions.presence_of_element_located(
                        (By.XPATH, '//*[@id="siteContent"]/div[@class="graphContainer"]')))
                sleep(1)
            except:
                sleep(60)
                self.log("Finding members failed.")
                continue
            # members count
            members = self.driver.find_elements(
                'xpath', '//aside/div/div/div[@role="listitem"]')
            members_count = len(members)

            new_member_clicked = False  # sent message to a new member ?
            # find member to send message
            new_member = None
            new_member_id = None

            for member in members:
                try:
                    member_id = member.find_element(
                        "xpath",
                        ".//div/div[2]/div[1]").text.replace("\n", " ")
                except:
                    break  # if not mounted yet, skip
                if "bot" in member_id.lower():
                    self.log("Bot!")
                    continue
                # if total_sent < 8: #skip first members
                #     visited_members.append(member_id)
                #     m_file.write(member_id + '\n')
                if member_id in visited_members:
                    continue
                new_member = member
                new_member_id = member_id
                break

            # click member && send message
            try:
                if new_member.is_displayed():
                    self.log('[clicking {}]'.format(new_member_id))
                    try:
                        new_member.click()
                    except selenium.common.exceptions.ElementClickInterceptedException as e:
                        continue

                    self.log('clicked {}'.format(new_member_id))
                    visited_members.append(new_member_id)
                    m_file.write(new_member_id + '\n')
                    self.log('Saved new member: {}'.format(new_member_id))

                    self.log('[Finding message box input]')
                    timestamp = WebDriverWait(self.driver, 30).until(
                        expected_conditions.presence_of_element_located(
                            (By.XPATH, '//input[@maxlength="999"]')))
                    self.log('Found message box input')

                    new_member_clicked = True  # visited a new member
                    # check time
                    cur_time = time.time()

                    # input text
                    chat_input = self.driver.find_element(
                        'xpath', '//input[@maxlength="999"]')
                    for msg in self.msg.split("\n"):
                        if not msg:
                            act.key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.ENTER).perform()
                            continue
                        pyperclip.copy(msg)
                        act.key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()
                        act.key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.ENTER).perform()
                    act.send_keys(Keys.RETURN).perform()

                    # press enter
                    sleep(randint(10,100))

                    total_sent += 1
                    self.log('Sent message to {}   (Total: {})'.format(
                        new_member_id, total_sent))
                    previous_time = time.time()
                    # check if too many messages
                    self.check_too_many()
            except Exception as e:
                self.log('Could not send message:{}'.format(str(e)))
                # if there is dialog, for example server boost , close it
                if self.close_dialog():
                    sleep(5)
                    continue  # no need to finish or scroll down

            if not new_member_clicked:

                finish_counter -= 1
                if finish_counter == 4:
                    a = 5
                # check if finish
                if finish_counter == 0:
                    self.log("Finished sending message.")
                    self.running = False
                # scroll down
                try:
                    self.log("Members count: {}".format(members_count))
                    middle = int(members_count * 0.7) + randrange(3)
                    if 0 < members_count - 4 < middle:
                        middle = members_count - 4
                    middle_id = members[middle].find_element(
                        'xpath',
                        ".//div/div[2]/div[1]").text.replace("\n", " ")
                    self.log(
                        "[Finish counter: {}, Scroll downing to {}]".format(
                            finish_counter, middle_id))
                    self.driver.execute_script(
                        'arguments[0].scrollIntoView(true);', members[middle])
                    self.log('Scroll down success.')
                    scroll_downed = True
                    sleep(2)
                except Exception as e:
                    self.log('Scroll down failed: {}'.format(str(e)))
            else:
                finish_counter = FINISH_COUNT
        # close driver
        self.driver.close()
