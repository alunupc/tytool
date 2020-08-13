import os
import zipfile
from os.path import join

import requests
from bs4 import BeautifulSoup as soup
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from unrar import rarfile

from log import Logger


class DownLoader:
    def __init__(self, timeout=None, url="", path=""):
        self.timeout = timeout
        self.chrome_options = webdriver.ChromeOptions()
        self.chrome_options.headless = True
        self.browser = webdriver.Chrome(chrome_options=self.chrome_options)
        self.browser.set_page_load_timeout(self.timeout)
        self.wait = WebDriverWait(self.browser, self.timeout)
        # self.path = path.rstrip(path.split("/")[-1])
        self.path = path
        self.url = url
        self.log = Logger('downloader.log', level='error')

    def __del__(self):
        self.browser.close()

    def download_file(self):
        try:
            self.browser.get(self.url)
            html = soup(self.browser.page_source, "lxml")
            for item in html.find_all('a'):
                if item.text.endswith('.pdf') or item.text.endswith('.xls') or item.text.endswith(
                        '.xlsx') or item.text.endswith('.doc') or item.text.endswith('.docx') or item.text.endswith(
                    '.rar') or item.text.endswith('.zip'):
                    element = self.browser.find_element_by_partial_link_text(item.text.split()[-1])
                    self.download(element.get_attribute('href'), item.text)
            for root, dirs, files in os.walk(self.path.replace("/", "\\")):
                for file in files:
                    if file.endswith(".rar") or file.endswith("zip"):
                        self.decompression(os.path.join(self.path.replace("/", "\\"), file),
                                           self.path.replace("/", "\\"))
        except Exception as e:
            print(e)
            self.log.logger.error(e)

        pass

    def download(self, url, name=""):
        try:
            res = requests.get(url)
            if name == "":
                file_name = url.split("/")[-1]
                if "?" in file_name:
                    file_name = file_name.split("?")[-1]
                if "=" in file_name:
                    file_name = file_name.split("=")[-1]
            else:
                file_name = name
                if "." not in file_name:
                    file_name += "." + url.split(".")[-1]
            with open(os.path.join(self.path.replace("/", "\\"), file_name), 'wb') as f:
                f.write(res.content)
        except Exception as e:
            self.log.logger.error(e)

    @staticmethod
    def decompression(file_path, target_path):
        """
        解压rar、zip压缩包到指定文件夹
        :param file_path: 压缩包路径
        :param target_path: 解压路径
        :return:
        """
        if file_path.endswith(".rar"):
            try:
                # mode的值只能为'r'
                rar_file = rarfile.RarFile(file_path, mode='r')
                # 得到压缩包里所有的文件
                rar_list = rar_file.namelist()
                for f in rar_list:
                    rar_file.extract(f, target_path)
            except Exception as e:
                Logger('decompress.log', level='error').logger.error(e)
        elif file_path.endswith(".zip"):
            # zip_file = zipfile.ZipFile(file_path, mode='r')
            with zipfile.ZipFile(file_path, mode='r') as zip_file:
                # 得到解压包所有文件
                zip_list = zip_file.namelist()
                for f in zip_list:
                    # 循环解压文件到指定目录
                    zip_file.extract(f, target_path)
                for file in os.listdir(target_path):
                    if file in zip_list:
                        try:
                            new_file = file.encode('cp437').decode('gbk')
                            print(new_file)
                            try:
                                os.rename(join(target_path, file), join(target_path, new_file))
                            except Exception as e:
                                os.remove(join(target_path, file))
                        except Exception as e:
                            Logger('decompress.log', level='error').logger.error(e)
        else:
            pass


if __name__ == '__main__':
    # path = r"C:\Users\localhost\Desktop\bb"
    # for root, dirs, files in os.walk(path.replace("/", "\\")):
    #     print(files)
    # DownLoader.decompression(r"C:\Users\localhost\Desktop\bb\bb.rar", r"C:\Users\localhost\Desktop\bb")
    path = "C:/Users/localhost/Desktop/bb/石家庄市2018年市本级和全市财政总决算报表.pdf"
    print(path.split("/")[-1])
    print('_' * 20, path.rstrip(path.split("/")[-1]))
    print('_' * 120)
    pass
