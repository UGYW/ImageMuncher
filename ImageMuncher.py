from pptx import *
from INPUTS import *
from utility import *
import os

temp_path_to_pptx = "temp.pptx"
temp_path_to_pics = ""

def main():
    if not os.path.exists(PATH_TO_DIRECTORIES):
        print(str(PATH_TO_DIRECTORIES) + " is not a valid path!")

    if not os.path.exists(PATH_TO_POWERPOINT):
        print(str(PATH_TO_POWERPOINT) + " is not a valid path!")

    # Make a new presentation
    prs = Presentation()

    sub_folder_paths = get_sub_folder_paths(PATH_TO_DIRECTORIES)

    # Make a new slide for each sub folder and its contents
    for sub_folder_path in sub_folder_paths:
        sub_folder_content_paths = get_sub_folder_content_paths(sub_folder_path)

        # Make a new slide
        slide = prs.slides.add_slide(prs.slide_layouts[1])  # title and content layout

        # Add the title - using the sub folder path
        add_title(slide, sub_folder_path)

        # Add the caption (the RGB image) image
        add_picture(slide, sub_folder_path[RGB], CAPTION_IMAGE_LOC, CAPTION_IMAGE_SIZE)

        # Add the four images
        add_picture(slide, sub_folder_content_paths[OXY], UPPER_LEFT, MAIN_IMAGE_SIZE)
        add_picture(slide, sub_folder_content_paths[THI], UPPER_RIGHT, MAIN_IMAGE_SIZE)
        add_picture(slide, sub_folder_content_paths[NIR], LOWER_LEFT, MAIN_IMAGE_SIZE)
        add_picture(slide, sub_folder_content_paths[TWI], LOWER_RIGHT, MAIN_IMAGE_SIZE)

    prs.save(PATH_TO_POWERPOINT)

if __name__ == '__main__':
    main()