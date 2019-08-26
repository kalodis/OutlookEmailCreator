"""
Script for generating outlook.com account with randomly generated data
Built with: Selenium, 2Captcha and Faker package
"""
import os
import random
import secrets
import shutil
import string
from time import sleep
from pprint import pprint

import requests
from faker import Faker
from captcha_solver import CaptchaSolver
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys

# 2Captcha API key here
API_2_CAPTCHA = 'xxxxxxxxxxxxxxx'

class OutlookAccountCreator:
    """ Class for creating outlook.com account
    with randomly generated details"""
    URL = 'https://signup.live.com/signup'
    IMG = 'captcha.jpg'

    def __init__(self):
        self.driver = webdriver.Chrome()

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

        if 'Phone number' in captcha_form.get_attribute('innerHTML'):
            print('Form asks for phone number...\nTerminating...')
            self.driver.quit()
            exit()

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
        return {
            "first_name": first,
            'last_name': last,
            'country': fake_details.country_code(representation="alpha-2"),
            'username': username,
            'password': password,
            'dob': fake_details.date_time()
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
        if OutlookAccountCreator.__download_image(captcha_url):
            print('Solving Captcha...')
            solver = CaptchaSolver('2captcha', api_key=API_2_CAPTCHA)
            raw_data = open(OutlookAccountCreator.IMG, 'rb').read()
            solution = solver.solve_captcha(raw_data)
            os.remove(OutlookAccountCreator.IMG)
            print(f"Captcha solved (solution: {solution})...")
            return solution
        print('Failed to download captcha image...')

    @staticmethod
    def __download_image(image_url: str):
        """
        Downloads captcha image
        :param image_url: string with url to image
        """
        print('Downloading Captcha image...')
        r = requests.get(image_url, stream=True)
        if r.status_code == 200:
            with open(OutlookAccountCreator.IMG, 'wb') as f:
                r.raw.decode_content = True
                shutil.copyfileobj(r.raw, f)
                return True
        print('Captcha image download failed', image_url)
        return False


account_creator = OutlookAccountCreator()
account_creator.create_account()
