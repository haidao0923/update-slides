from pptx import Presentation
from pptx.util import Inches
import win32com.client
import os
powerpoint_name = '[Name of Powerpoint to Open]'
image_folder = '[Folder to save images]'
presentation = Presentation(powerpoint_name)
slide = presentation.slides[0]
DATE_INDEX = 2
TIME_INDEX = 3
LOCATION_INDEX = 4
image_name = '[Name of Saved Image]'
date_text = '[Date of Meeting]'
location_text = '[Location of Meeting]'
text_array = []

def get_all_text():
    for element in slide.shapes:
        if element.has_text_frame:
            text_array.append(element);


def modify_text(index, new_text):
    text_frame = text_array[index].text_frame
    paragraph = text_frame.paragraphs[0]
    paragraph.runs[0].text = new_text

def save():
    presentation.save(powerpoint_name)
    powerpoint_path = os.getcwd() + f'/{powerpoint_name}'
    Application = win32com.client.Dispatch('Powerpoint.Application')
    powerpoint = Application.Presentations.Open(powerpoint_path)
    image_path = os.getcwd() + f'/{image_folder}/{image_name}'
    Application.ActivePresentation.Slides[0].Export(image_path, 'PNG')
    powerpoint.Close()
    Application.Quit()
    os.system('taskkill /F /IM POWERPNT.EXE')

get_all_text()
modify_text(DATE_INDEX, date_text)
modify_text(LOCATION_INDEX, location_text)
save();
