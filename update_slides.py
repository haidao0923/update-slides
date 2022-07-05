from pptx import Presentation
from pptx.util import Inches
import win32com.client
import os
DATE_INDEX = 2
TIME_INDEX = 3
LOCATION_INDEX = 4
text_array = []

from PIL import Image

def get_all_text(powerpoint_name, presentation):
    global text_array
    text_array = []
    slide = presentation.slides[0]
    for element in slide.shapes:
        if element.has_text_frame:
            text_array.append(element);


def modify_text(index, new_text):
    text_frame = text_array[index].text_frame
    paragraph = text_frame.paragraphs[0]
    paragraph.runs[0].text = new_text

def save(powerpoint_name, presentation, image_folder, image_name):
    presentation.save(powerpoint_name)
    powerpoint_path = os.getcwd() + f'/{powerpoint_name}'
    Application = win32com.client.Dispatch('Powerpoint.Application')
    powerpoint = Application.Presentations.Open(powerpoint_path)
    image_path = os.getcwd() + f'/{image_folder}/{image_name}'
    Application.ActivePresentation.Slides[0].Export(image_path, 'PNG')
    powerpoint.Close()
    Application.Quit()
    os.system('taskkill /F /IM POWERPNT.EXE')

def execute(powerpoint_name, image_folder, image_name, date_text, time_text, location_text):
    presentation = Presentation(powerpoint_name)
    get_all_text(powerpoint_name, presentation)
    for i in range(len(text_array)):
        print(str(i) + text_array[i].text)
    modify_text(DATE_INDEX, date_text)
    modify_text(TIME_INDEX, time_text)
    modify_text(LOCATION_INDEX, location_text)
    save(powerpoint_name, presentation, image_folder, image_name);

image_date = '07_10_22'
text_date = '07/10/22'
online_image_date = '07_07_22'
online_text_date = '07/07/22'
'''
execute('GGTemplate.pptx', 'Weekly_Images', f'Engage_{image_date}.png',
        f'Sunday {text_date}', '2pm - 5pm', 'Crosland Tower 2nd Floor')
execute('GGTemplateSquare.pptx', 'Weekly_Images', f'Instagram_{image_date}.png',
        f'Sunday {text_date}', '2pm - 5pm', 'Crosland Tower 2nd Floor')
execute('GGOnlineTemplate.pptx', 'Weekly_Images', f'Engage_Online_{online_image_date}.png',
        f'Thursday {online_text_date}', '8pm - 10pm', 'Discord Voice Chat - Online')
execute('GGOnlineTemplateSquare.pptx', 'Weekly_Images', f'Instagram_Online_{online_image_date}.png',
        f'Thursday {online_text_date}', '8pm - 10pm', 'Discord Voice Chat - Online')
'''
first_image = Image.open(os.getcwd() + '/Weekly_Images/' + f'Instagram_Online_{online_image_date}.png')
display = Image.new('RGB', (first_image.width * 2 + 10, first_image.height))
display.paste(first_image, (0, 0))
second_image = Image.open(os.getcwd() + '/Weekly_Images/' + f'Instagram_{image_date}.png')
display.paste(second_image, (first_image.width, 0))
display.save(os.getcwd() + '/Weekly_Images/' + f'Combined_Instagram_{image_date}.png')
