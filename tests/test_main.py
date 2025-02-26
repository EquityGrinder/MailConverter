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

# Maximum width in pixels for an image to fit in a DIN A4 page size (at a common screen DPI)
MAX_IMAGE_WIDTH = (8.27 -  2) * 96 


def get_plain_text_body(mail_item):
    '''
    this function expects a MailItem from the Outlook and return the text from the body
    '''
    
    #     # Check if mail item's body format is plain text
    # if mail_item.BodyFormat == 1:
    #     text_body = mail_item.Body
    # elif mail_item.BodyFormat == win32com.client.constants.olFormatHTML:
    #     text_body = mail_item.HTMLBody
    #     # If you want to convert the HTML body to plain text, additional processing will be needed
    #     # You could use a library like BeautifulSoup or lxml to extract text from HTML
    #     # For the sake of simplicity, let's just extract text content as is for now
    #     # Please note that this will still contain HTML tags
    # elif mail_item.BodyFormat == win32com.client.constants.olFormatRichText:
    #     # Rich Text format can be a bit more complex to handle, and may require
    #     # conversion to another format to extract plain text. For now, let's
    #     # assume it's not handled.
    #     text_body = "Rich Text Format is not supported for plain text extraction."
    # else:
    #     text_body = "Unknown body format."

    return mail_item.Body
    
def generate_html(plain_text):
    '''
    This function generates a html from a plain text
    Special characters are also replaced to be HTML-compatible.
    '''
    # Convert special characters to HTML entities to prevent breaking the HTML
    html_text = plain_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    html_text = html_text.replace("'", "&apos;").replace('"', "&quot;").replace('\n', '<br>')

    # Wrap the text in HTML paragraph tags
    html = f"<!DOCTYPE html><html><head><title>Mail Content</title></head><body><p>{html_text}</p></body></html>"
    
    return html

def extract_html(mail_item):
    '''
    This function expect a MailItem from the COM outlook interface.
    It searches for the body  of the mail. If the body is not html it will generate a html
    
    
    '''
    html = None
    
    # Check if the mail item's body format is already HTML
    if mail_item.BodyFormat == 2 or mail_item.BodyFormat == 1:
        html = mail_item.HTMLBody

    if not html:
        plain_text = get_plain_text_body(mail_item)
        html = generate_html(plain_text)
    return html

def insert_html(mail_item, html):
    '''
    this file replaces the old html text_body with the new html text body   
    
        Args:
        mail_item: An instance of a MailItem whose HTML body needs to be updated.
        html: A string containing the new HTML content to be set as the MailItem's HTMLBody.
    Returns:
        mail_item: The updated MailItem instance with the new HTML content.
    '''
    # Set the HTMLBody property with the new HTML content
    mail_item.HTMLBody = html
    return mail_item



def scale_image_dimensions(width, height, max_width):
    """
    Scales the width and height of an image proportionally, given a maximum width constraint.
    """
    # If the image is already within the size limit, return the original dimensions
    if width <= max_width:
        return width, height
    
    # Calculate the scaling factor and apply it to both dimensions
    scaling_factor = max_width / width
    scaled_width = max_width
    scaled_height = int(height * scaling_factor)
    
    return scaled_width, scaled_height

def is_float(value):
    try:
        float(value)
        return True
    except ValueError:
        return False

def isnumeric(value):
    if value.isnumeric():
        return True
    else:
        if is_float(value):
            return True
        else:
            return False

def scale_images(html):
    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    
    # Find all <img> tags in the HTML
    images = soup.find_all('img')  
    
    # Loop through all found images and scale them if necessary
    for img in images:
        width, height = img.get('width'), img.get('height')
        
        # If width and height are specified, and are numerical values
        if width and height and isnumeric(width) and isnumeric(height):
            width, height = int(float(width)), int(float(height))
            
            scaled_width, scaled_height = scale_image_dimensions(width, height, MAX_IMAGE_WIDTH)
            
            # Update the image tag with new dimensions
            img['width'] = scaled_width
            img['height'] = scaled_height
        else:
            # If width and height are not provided, you can set a default max width
            # Note: this may distort images that are originally larger in height than width
            img['width'] = MAX_IMAGE_WIDTH


    # Return the transformed HTML as a string
    return str(soup)

def add_newlines_around_images(html):
    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    
    # Find all <img> tags in the HTML
    images = soup.find_all('img')

    # Loop through all found images and insert a <br> before and after each
    for img in images:
        # Create a new <br> tag to insert before the image
        newline_before = soup.new_tag('br')
        # Create another new <br> tag to insert after the image
        newline_after = soup.new_tag('br')
        
        # Insert the <br> tags before and after the <img> tag
        img.insert_before(newline_before)
        img.insert_after(newline_after)

    # Return the modified HTML as a string
    return str(soup)

def add_a4_print_styles(html):
    # Parse the HTML string using BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')

    # Check if head exists, else create one
    head = soup.head
    if head is None:
        head = soup.new_tag('head')
        soup.insert(0, head)

    # Create a <style> tag with print CSS for A4 size
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
        max-width: 210mm; /* A4 width minus margins */
        margin: 0 auto;
    }
    """
    style_tag.string = style_css
    head.append(style_tag)

    # Wrap content with a container div if not already present
    body = soup.body
    if body and not body.find(class_='container'):
        container = soup.new_tag('div', **{'class': 'container'})
        for content in body.contents:
            container.append(content.extract())
        body.append(container)

    # Return the modified HTML as a string
    return str(soup)

def transform_html(html):
    html = add_newlines_around_images(html)
    html = scale_images(html)
    html = add_a4_print_styles(html)
    return html

def transform_mail_item(mail_item):
    
    html = extract_html(mail_item)
    html = transform_html(html)
    mail_item = insert_html(mail_item, html)
    
    return mail_item


def convert_msg_to_mht(msg_path, output_path):
    pythoncom.CoInitialize()  # Initialize the COM library for the current thread
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        mail_item = outlook.OpenSharedItem(msg_path)
        # check here for meeting msg
        if not (mail_item.MessageClass.startswith("IPM.Schedule.Meeting") or mail_item.MessageClass == "IPM.Appointment"):
            mail_item = transform_mail_item(mail_item)
            # This will save the email in the .mht format which is a webpage archive format
        mail_item.SaveAs(output_path, 10)  # olFormatHTML is 5    
        mail_item.Close(0)  # olDiscard is 0
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        pythoncom.CoUninitialize()  # Uninitialize the COM library


def list_msg_files_in_directory(directory_path):
    # Construct the pattern for .msg files
    pattern = os.path.join(directory_path, '*.msg')
    # Use glob to find all .msg files in the directory
    msg_files = glob.glob(pattern)
    return msg_files

def transform_mht(file_path):
    '''This functuion sets all german tags to english'''

    
    # Read the content of the file
    with open(file_path, 'r') as file:
        content = file.read()

    # Replace the first occurrence
    content = content.replace(">Von:<", 
                              ">From:<")

    # Replace the first occurrence
    content = content.replace(">Gesendet:<", 
                              ">Sent:<")
    content = content.replace(">An:<", 
                              ">To:<")
    content = content.replace(">Betreff:<", 
                              ">Subject:<")
    content = content.replace(">Anlagen:<", 
                              ">Attachements:<")
    content = content.replace(">Kategorien:<", 
                              ">Categories:<")
    content = content.replace(">Priorität:<", 
                              ">Priority:<")

        
    # Write the modified content back to the file
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(content)


# Function to open files in Word and Outlook, and try to move them to different monitors
def open_files_on_different_monitors(msg_file, mht_file):
    # Open the MSG file with Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItemFromTemplate(msg_file)
    mail.Display()

    # Open the MHT file with Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    doc = word.Documents.Open(mht_file)

    # Allow some time for the applications to open
    time.sleep(2)

    # Get monitors information
    monitors = Desktop(backend="uia").monitors

    # Set target positions for Outlook and Word windows
    outlook_target_monitor = monitors[0]  # Default to the first monitor
    word_target_monitor = monitors[1]   # Second monitor if available

    # Get the list of all opened windows
    windows = gw.getAllTitles()

    # position Outlook and Word on different monitors
    try:
        for w in windows:
            if msg_file in w:  # Check if window title contains the .msg file name
                win = gw.getWindowsWithTitle(w)[0]
                win.moveTo(outlook_target_monitor.left, outlook_target_monitor.top)
                win.resizeTo(outlook_target_monitor.width, outlook_target_monitor.height)

            if mht_file in w:  # Check if window title contains the .mht file name
                win = gw.getWindowsWithTitle(w)[0]
                win.moveTo(word_target_monitor.left, word_target_monitor.top)
                win.resizeTo(word_target_monitor.width, word_target_monitor.height)

    except Exception as e:
        print(f"An error occurred while moving windows: {e}")


def process_files(file_list, directory):
    # Define the new directory path for the 'mht' subfolder
    mht_dir_path = os.path.join(directory, 'mht')
    # Create the 'mht' directory if it doesn't already exist
    os.makedirs(mht_dir_path, exist_ok=True)

    for file_path in file_list:
        # Only process files that end with .msg
        if file_path.lower().endswith('.msg'):
            # Get the base name without the .msg extension
            file_base_name = os.path.basename(file_path)[:-4]
            # Define the new MHT file path inside the 'mht' subfolder
            new_mht_file_path = os.path.join(mht_dir_path, file_base_name + '.mht')
            # Call the function to convert MSG to MHT
            convert_msg_to_mht(file_path, new_mht_file_path)

            # Transform the MHT file content (if needed)
            transform_mht(new_mht_file_path)
            
            open_files_on_different_monitors(file_path, new_mht_file_path)
            input('Press Enter to continue...')

def main(args, debug):

    if debug:
        msg_dir = args
    else:
        msg_dir = args.directory

    file_list = list_msg_files_in_directory(msg_dir)
    process_files(file_list, msg_dir)

if __name__ == "__main__":
    debug = True
    if debug:
        args = "c:/Users/r01grigo/repos/MailConverter/data"
    else:
        parser = argparse.ArgumentParser(description='Process .msg files from a given directory.')
        parser.add_argument('directory', help='The path to the directory containing .msg files to be converted.')
        args = parser.parse_args()
     
    main(args, debug)