
""" 
Author - MH
Most Recent Update - 05/01/2023

Description -
Imports a two column xlsx doc and converts the first column to file names, 
the second to an image of text limitted to a certain rectangle size, nicely 
formatted.

"""

# Direct the xlsx_doc variable to the xlsx document's directory. 
# To change font, download a new ttf file and direct the fontname to it's directory.


from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import textwrap
import pandas as pd

# Load variables
xlsx_doc = pd.read_excel(r'G:\\MissingMyClothesLoads\\strings.xlsx')
string_contents = xlsx_doc.values 
img_strings = string_contents[:,1]
sav_strings = string_contents[:,0]
sav_parent = "G:\\MissingMyClothesSaves\\" #set save directory here

# Loop over the image strings
for i in range(len(img_strings)):
    def getSize(txt, font): 
        test_img = Image.new('RGB', (1, 1))
        test_draw = ImageDraw.Draw(test_img)
        return test_draw.textsize(txt, font)

    if __name__ == '__main__':
        fontname = ".arial\\arial.ttf" #font ttf file's directory here
        fontsize = 40  
        text = str(img_strings[i])
        sav_text = str(sav_strings[i])
        
        # Use the widest letter in the English alphabet to set an upper limit on text size
        test_text = "w" 
       
        # Set text characteristics
        text_colour = "black"
        # outline_colour = "white" add to img render if desired
        background_colour = "white"
        font = ImageFont.truetype(fontname, fontsize)
        width, height = getSize(text, font)
        width_2, height_2 = getSize(test_text,font)
        
        # Determine the maximum number of letters per image line.
        max_num = round(500 / width_2) - 1 
        wrapper = textwrap.TextWrapper(width = max_num)
        word_list = wrapper.wrap(text = text) #turn wrapped text into a set of strings
        caption_new = ''
        
        for ii in word_list[:-1]:
            caption_new = caption_new + ii + '\n'
        
        caption_new += word_list[-1]
        
        # Create the image
        img = Image.new('RGB', (500, height+50*(len(word_list))+10), background_colour) #height of the image is based on the length of the word list
        placeholder_img = ImageDraw.Draw(img)
        text_width, text_height = placeholder_img.textsize(caption_new, font = font)
        img_width, img_height = img.size
        initial_height, padding = 10, 10
        
        # Centering and stacking of words under each other
        for line in word_list: 
            text_width, text_height = placeholder_img.textsize(line, font = font)
            placeholder_img.text(((img_width - text_width) / 2, initial_height), line,fill=text_colour, font = font)
            initial_height += text_height + padding
        
        save_directory = sav_parent + sav_text + ".png" 
        img.save(save_directory)
