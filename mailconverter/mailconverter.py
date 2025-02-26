import os
import pythoncom
import win32com.client
from bs4 import BeautifulSoup
import glob
import argparse
from time import sleep
import os
import time
import subprocess
import win32com.client
import pygetwindow as gw
from pywinauto import Desktop


class MailConverter:
    # Maximum width in pixels for an image to fit in a DIN A4 page size (at a common screen DPI)
    __MAX_IMAGE_WIDTH = (8.27 - 2) * 96

    def __init__(self, path="data", debug=False, interface="console"):
        """
        Initialize the MailConverter with the given path.
        """
        self.__path = path
        self.__files = []
        self.__debug = debug
        self.__interface = interface

    def start(self):
        """
        Start the conversion process for .msg files in the specified directory.
        """
        if not self.__debug:
            self.__start_interface()
            parser = argparse.ArgumentParser(description='Process .msg files from a given directory.')
            parser.add_argument('directory', help='The path to the directory containing .msg files to be converted.')
            args = parser.parse_args()
            self.__path = args.directory
        
        self.__list_msg_files_in_directory()
        self.__process_files()
    def __start_interface(self):
        """
        Start the interface for the MailConverter.
        """
        if self.__interface == "console":
            self.__start_console_interface()
        
    def __get_plain_text_body(self, mail_item):
        """
        Extract the plain text body from a MailItem.
        """
        return mail_item.Body

    def __generate_html(self, plain_text):
        """
        Generate HTML from plain text.
        """
        html_text = plain_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        html_text = html_text.replace("'", "&apos;").replace('"', "&quot;").replace('\n', '<br>')
        html = f"<!DOCTYPE html><html><head><title>Mail Content</title></head><body><p>{html_text}</p></body></html>"
        return html

    def __extract_html(self, mail_item):
        """
        Extract HTML from a MailItem, or generate it if not available.
        """
        html = None
        if mail_item.BodyFormat == 2 or mail_item.BodyFormat == 1:
            html = mail_item.HTMLBody
        if not html:
            plain_text = self.__get_plain_text_body(mail_item)
            html = self.__generate_html(plain_text)
        return html

    def __insert_html(self, mail_item, html):
        """
        Insert HTML into a MailItem.
        """
        mail_item.HTMLBody = html
        return mail_item

    def __scale_image_dimensions(self, width, height, max_width):
        """
        Scale image dimensions proportionally to fit within the maximum width.
        """
        if width <= max_width:
            return width, height
        scaling_factor = max_width / width
        scaled_width = max_width
        scaled_height = int(height * scaling_factor)
        return scaled_width, scaled_height

    def __is_float(self, value):
        """
        Check if a value can be converted to a float.
        """
        try:
            float(value)
            return True
        except ValueError:
            return False

    def __isnumeric(self, value):
        """
        Check if a value is numeric.
        """
        if value.isnumeric():
            return True
        else:
            if self.__is_float(value):
                return True
            else:
                return False

    def __scale_images(self, html):
        """
        Scale images in the HTML content to fit within the maximum width.
        """
        soup = BeautifulSoup(html, 'html.parser')
        images = soup.find_all('img')
        for img in images:
            width, height = img.get('width'), img.get('height')
            if width and height and self.__isnumeric(width) and self.__isnumeric(height):
                width, height = int(float(width)), int(float(height))
                scaled_width, scaled_height = self.__scale_image_dimensions(width, height, self.__MAX_IMAGE_WIDTH)
                img['width'] = scaled_width
                img['height'] = scaled_height
            else:
                img['width'] = self.__MAX_IMAGE_WIDTH
        return str(soup)

    def __add_newlines_around_images(self, html):
        """
        Add newlines around images in the HTML content.
        """
        soup = BeautifulSoup(html, 'html.parser')
        images = soup.find_all('img')
        for img in images:
            newline_before = soup.new_tag('br')
            newline_after = soup.new_tag('br')
            img.insert_before(newline_before)
            img.insert_after(newline_after)
        return str(soup)

    def __add_a4_print_styles(self, html):
        """
        Add A4 print styles to the HTML content.
        """
        soup = BeautifulSoup(html, 'html.parser')
        head = soup.head
        if head is None:
            head = soup.new_tag('head')
            soup.insert(0, head)
        style_tag = soup.new_tag('style', type='text/css')
        style_tag.attrs['media'] = 'print'
        style_css = """
        @page {
            size: A4;
            margin: 20mm;
        }
        body {
            margin: 0;
            padding: 0;
        }
        .container {
            width: 100%;
            max-width: 210mm;
            margin: 0 auto;
        }
        """
        style_tag.string = style_css
        head.append(style_tag)
        body = soup.body
        if body and not body.find(class_='container'):
            container = soup.new_tag('div', **{'class': 'container'})
            for content in body.contents:
                container.append(content.extract())
            body.append(container)
        return str(soup)

    def __transform_html(self, html):
        """
        Transform the HTML content by adding newlines around images, scaling images, and adding A4 print styles.
        """
        html = self.__add_newlines_around_images(html)
        html = self.__scale_images(html)
        html = self.__add_a4_print_styles(html)
        return html

    def __transform_mail_item(self, mail_item):
        """
        Transform the MailItem by extracting, transforming, and inserting HTML content.
        """
        html = self.__extract_html(mail_item)
        html = self.__transform_html(html)
        mail_item = self.__insert_html(mail_item, html)
        return mail_item

    def __convert_msg_to_mht(self, msg_path, output_path):
        """
        Convert a .msg file to .mht format.
        """
        flag = True
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            mail_item = outlook.OpenSharedItem(msg_path)
            if not (mail_item.MessageClass.startswith("IPM.Schedule.Meeting") or mail_item.MessageClass == "IPM.Appointment"):
                mail_item = self.__transform_mail_item(mail_item)
            mail_item.SaveAs(output_path, 10)
            mail_item.Close(0)
        except Exception as e:
            print(f"An error occurred: {e}")
            flag = False
        finally:
            pythoncom.CoUninitialize()

        return flag
    
    def __list_msg_files_in_directory(self):
        """
        List all .msg files in the specified directory.
        """
        pattern = os.path.join(self.__path, '*.msg')
        msg_files = glob.glob(pattern)
        self.__files = msg_files

    def __transform_mht(self, file_path):
        """
        Transform the content of an .mht file by replacing German tags with English tags.
        """
        with open(file_path, 'r') as file:
            content = file.read()
        content = content.replace(">Von:<", ">From:<")
        content = content.replace(">Gesendet:<", ">Sent:<")
        content = content.replace(">An:<", ">To:<")
        content = content.replace(">Betreff:<", ">Subject:<")
        content = content.replace(">Anlagen:<", ">Attachements:<")
        content = content.replace(">Kategorien:<", ">Categories:<")
        content = content.replace(">Priorität:<", ">Priority:<")
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(content)

    def __open_files_on_different_monitors(self, msg_file, mht_file):
        """
        Open .msg and .mht files on different monitors.
        """
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItemFromTemplate(msg_file)
        mail.Display()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        doc = word.Documents.Open(mht_file)
        time.sleep(2)
        monitors = Desktop(backend="uia").monitors
        outlook_target_monitor = monitors[0]
        word_target_monitor = monitors[1]
        windows = gw.getAllTitles()
        try:
            for w in windows:
                if msg_file in w:
                    win = gw.getWindowsWithTitle(w)[0]
                    win.moveTo(outlook_target_monitor.left, outlook_target_monitor.top)
                    win.resizeTo(outlook_target_monitor.width, outlook_target_monitor.height)
                if mht_file in w:
                    win = gw.getWindowsWithTitle(w)[0]
                    win.moveTo(word_target_monitor.left, word_target_monitor.top)
                    win.resizeTo(word_target_monitor.width, word_target_monitor.height)
        except Exception as e:
            print(f"An error occurred while moving windows: {e}")

    def __process_files(self):
        """
        Process the list of .msg files and convert them to .mht format.
        """
        mht_dir_path = os.path.join(self.__path, 'mht')
        os.makedirs(mht_dir_path, exist_ok=True)
        for file_path in self.__files:
            if file_path.lower().endswith('.msg'):
                ## todo this could be achieved more pragmatically somewhere in the code we have the filenames earlier
                file_base_name = os.path.basename(file_path)[:-4]
                new_mht_file_path = os.path.join(mht_dir_path, file_base_name + '.mht')
                #################################################################################################
                if self.__convert_msg_to_mht(file_path, new_mht_file_path):
                    self.__transform_mht(new_mht_file_path)
                    self.__open_files_on_different_monitors(file_path, new_mht_file_path)
                
                    if self.__debug:
                        print(f"Converted {file_path} to {new_mht_file_path}")
                        sleep(0.5)
                    else:
                        input('Press Enter to continue...')
