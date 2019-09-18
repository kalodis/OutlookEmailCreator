"""
Script for generating outlook.com account with randomly generated data
Built with: Selenium, 2Captcha and Faker package
"""
import os
import random
import secrets
import shutil
import string
import zipfile
from time import sleep
from pprint import pprint
from uuid import uuid4
from random import choice

import requests
from faker import Faker
from captcha_solver import CaptchaSolver
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

# 2Captcha API key here
from proxy_auth import manifest_json, background_js, plugin_file

API_2_CAPTCHA = '6ed35827f9a7e40e1c1acb767f8aa5e3'


class Proxies:
    proxy_list = []

    @staticmethod
    def load_proxies(file_path: str):
        """
        Reads a text file with proxies
        :param file_path: Path to proxy file with proxies in <user>:<pas>@<ip>:<port> format each on one line
        """
        lst = []
        if file_path:
            if os.path.exists(file_path):
                with open(file_path, 'r') as file:
                    lst = [x for x in file.read().split('\n') if x.strip()]
            else:
                print('File: {}. Does not exist.'.format(file_path))
        Proxies.proxy_list = lst

    @staticmethod
    def get_random_proxy():
        """ Returns a random proxy """
        return choice(Proxies.proxy_list)


class OutlookAccountCreator:
    """ Class for creating outlook.com account
    with randomly generated details"""
    URL = 'https://signup.live.com/signup'

    def __init__(self, use_proxy: bool = False):
        self.driver = self.__open_browser(use_proxy)

    def create_account(self):
        """
        Goes through website process of creating the account
        :return: dictionary with login information for the account
        """
        print('Creating new Outlook email account')
        self.driver.get(self.URL)
        sleep(2)
        self.driver.find_element_by_id('liveSwitch').click()
        sleep(2)

        person = self.__generate_random_details()
        birth_date = person['dob']

        # Enter Email
        ActionChains(self.driver) \
            .send_keys_to_element(self.driver.find_element_by_id('MemberName'), person['username']) \
            .send_keys(Keys.ENTER).pause(3).perform()
        # Enter Password
        ActionChains(self.driver) \
            .send_keys_to_element(self.driver.find_element_by_id('PasswordInput'), person['password']) \
            .send_keys(Keys.ENTER).pause(3).perform()
        # Enter First and Last Name
        ActionChains(self.driver) \
            .send_keys_to_element(self.driver.find_element_by_id('FirstName'), person['first_name']) \
            .send_keys_to_element(self.driver.find_element_by_id('LastName'), person['last_name']) \
            .send_keys(Keys.ENTER).pause(3).perform()

        # Enter Country and DOB
        self.driver.find_element_by_xpath(f'//option[@value="{person["country"]}"]').click()
        year_select = self.driver.find_element_by_xpath('//*[@id="BirthYear"]')
        year_select.find_element_by_xpath(f'//*[@value="{birth_date.year}"]').click()
        sleep(1)

        day_select = self.driver.find_element_by_xpath('//*[@id="BirthDay"]')
        day_select.find_element_by_xpath(f'//*[@value="{birth_date.day}"]').click()
        sleep(1)

        month_select = self.driver.find_element_by_xpath('//*[@id="BirthMonth"]')
        month_select.find_element_by_xpath(f'//*[@value="{birth_date.month}"]').click()
        sleep(1)

        self.driver.find_element_by_id('iSignupAction').click()
        sleep(5)

        # Solve Captcha
        try:
            captcha_form = self.driver.find_element_by_id('HipPaneForm')
        except:
            # Retry if something went wrong
            print('Failed while creating account...\nRetrying...')
            return self.create_account()

        # if 'Phone number' in captcha_form.get_attribute('innerHTML'):
        #     print('Form asks for phone number...\nTerminating...')
        #     self.driver.quit()
        #     exit()

        captcha_image_url = captcha_form \
            .find_element_by_tag_name('img').get_attribute('src')
        solution = self.__solve_captcha(captcha_image_url).replace(' ', '')
        self.driver.find_element_by_xpath('//input[@aria-label="Enter the characters you see"]') \
            .send_keys(solution)
        self.driver.find_element_by_id('iSignupAction').click()
        sleep(4)

        if 'account.microsoft' in self.driver.current_url:
            email = person['username'] + '@outlook.com'
            print(f'Account created successfully ({email})...')
            person['dob'] = person['dob'].strftime('%d, %b %Y')
            person['email'] = email
            pprint(person, indent=4)
            return person

        print('Failed to create account...')
        return None

    @staticmethod
    def __generate_random_details():
        """
        Generates random details for new account
        :return: dictionary with fake details
        """
        fake_details = Faker()
        name = fake_details.name()
        username = OutlookAccountCreator.__create_username(name)
        password = OutlookAccountCreator.__generate_password()
        first, last = name.split(' ', 1)

        while True:
            dob = fake_details.date_time()
            if dob.year < 2000:
                break
            if dob.month != 2:
                break
        while True:
            country = fake_details.country_code(representation="alpha-2")
            if country != "GB":
                break
        return {
            "first_name": first,
            'last_name': last,
            'country': country,
            'username': username,
            'password': password,
            'dob': dob
        }

    @staticmethod
    def __create_username(name: str):
        """
        Creates username based on name
        :param name: string with person name
        :return: string with username based on the name
        """
        return name.replace(' ', '').lower() + str(random.randint(1000, 10000))

    @staticmethod
    def __generate_password():
        """
        generates password 10 char long, with at least one number and symbol
        :return: string with new password
        """
        alphabet = string.ascii_letters + string.digits
        password = ''.join(secrets.choice(alphabet) for i in range(8))
        return password + random.choice('$#@!%^') + random.choice('0123456789')

    @staticmethod
    def __solve_captcha(captcha_url: str):
        """
        downloads captcha image and send to 2captcha to solve
        :param captcha_url: Captcha image url
        :return: string with captcha solution
        """
        img_name = f'{uuid4()}.jpg'
        if OutlookAccountCreator.__download_image(captcha_url, img_name):
            print('Solving Captcha...')
            solver = CaptchaSolver('2captcha', api_key=API_2_CAPTCHA)
            raw_data = open(img_name, 'rb').read()
            solution = solver.solve_captcha(raw_data)
            os.remove(img_name)
            print(f"Captcha solved (solution: {solution})...")
            return solution
        print('Failed to download captcha image...')

    @staticmethod
    def __download_image(image_url: str, image_name: str):
        """
        Downloads captcha image
        :param image_url: string with url to image
        :param image_name: string with image name
        :return: boolean, True if successful False is failed
        """
        print('Downloading Captcha image...')
        r = requests.get(image_url, stream=True)
        if r.status_code == 200:
            with open(image_name, 'wb') as f:
                r.raw.decode_content = True
                shutil.copyfileobj(r.raw, f)
                return True
        print('Captcha image download failed', image_url)
        return False

    @staticmethod
    def __open_browser(use_proxy: bool = False):
        # TODO: add user agent
        chrome_options = webdriver.ChromeOptions()
        if use_proxy:
            random_proxy = Proxies.get_random_proxy()
            # Parse Proxy
            auth, ip_port = random_proxy.split('@')
            user, pwd = auth.split(':')
            ip, port = ip_port.split(':')

            with zipfile.ZipFile(plugin_file, 'w') as zp:
                zp.writestr("manifest.json", manifest_json)
                zp.writestr("background.js", background_js % (ip, port, user, pwd))
            chrome_options.add_extension(plugin_file)

        return webdriver.Chrome(chrome_options=chrome_options)


if __name__ == '__main__':
    # Load proxies from file
    Proxies.load_proxies('proxies.txt')
    # Initialize account creator class
    account_creator = OutlookAccountCreator(use_proxy=True)
    # Run account creator
    account_creator.create_account()
