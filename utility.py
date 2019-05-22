# UTILIY FUNCTIONS FOR GETTING IMAGE DATA
from constants import *
import os
from pptx.util import Cm

# Reference: https://python-pptx.readthedocs.io/en/latest/api/shapes.html

def get_sub_folder_paths(path_to_main_folder):
    contents = os.listdir(path_to_main_folder)
    # Adds the path to the main folder in front for traversal
    contents = [path_to_main_folder + "/" + i for i in contents if not '.' in i]
    return contents

def get_sub_folder_content_paths(path_to_sub_folder):
    # Make a dictionary based on type
    res = {}
    contents = os.listdir(path_to_sub_folder)
    res[RGB] = [path_to_sub_folder + "/" + i for i in contents if RGB in i][0]
    res[OXY] = [path_to_sub_folder + "/" + i for i in contents if OXY in i][0]
    res[THI] = [path_to_sub_folder + "/" + i for i in contents if THI in i][0]
    res[NIR] = [path_to_sub_folder + "/" + i for i in contents if NIR in i][0]
    res[TWI] = [path_to_sub_folder + "/" + i for i in contents if TWI in i][0]
    return res

def add_picture(slide, picture, location, size = (-1, -1)):
    width = None
    height = None
    if size[0] > 0: width = size[0]
    if size[1] > 0: height = size[1]
    slide.shapes.add_picture(picture, location[0], location[1], width, height)

def add_title(slide, title_text):
    slide.shapes.title.text = title_text

def get_date(path):
    return os.path.basename(os.path.normpath(path))

# https://github.com/scanny/python-pptx/issues/195
def add_background(prs, slide, background_pic):
    left = top = Cm(0)
    pic = slide.shapes.add_picture(background_pic, left, top,
                                    width=prs.slide_width, height=prs.slide_height)
    slide.shapes._spTree.remove(pic._element)
    slide.shapes._spTree.insert(2, pic._element)

def add_table(slide, rows, cols, location, size = (-1, -1)):
    width = None
    height = None
    if size[0] > 0: width = size[0]
    if size[1] > 0: height = size[1]
    slide.shapes.add_table(rows, cols, location[0], location[1], width, height)