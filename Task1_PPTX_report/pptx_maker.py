from pptx import Presentation
from pptx.util import Inches, Pt
import json
from pptx.chart.data import XySeriesData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE
import matplotlib.pyplot as plt
import tempfile
import os
import logging

#TODO AddImg exceptions, if needed?
#TODO test exception handling
#TODO commenting and docstrings


def addTitleSlide(presentation, title_text, subtitle_text):
    slide_layout = presentation.slide_layouts[0]  
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = title_text
    subtitle.text = subtitle_text

def addTextSlide(presentation, title_text, text):
    slide_layout = presentation.slide_layouts[1]
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    body = slide.placeholders[1]
    title.text = title_text
    body.text = text

def addListSlide(presentation, title_text, list_json): 
    slide_layout = presentation.slide_layouts[1]
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = title_text
    content = slide.placeholders[1]
    content.text = ""
    for list_item in list_json:
        level = list_item["level"]
        if not level>0:
            raise ValueError(f"Invalid 'level' attribute in JSON entry {list_item} in the list slide titled {title_text}")
        text = list_item["text"]
        paragraph = content.text_frame.add_paragraph()
        paragraph.text = text
        paragraph.level = level - 1 

        
def addImgSlide(presentation, title_text, img_path):
    slide_layout = presentation.slide_layouts[5] 
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = title_text
    slide.shapes.add_picture(img_path, Inches(1), Inches(2))

def readDataFile(filepath):  #TODO better exception handling. #Reads the data for plot slides. I assumed that the data consists of pairs of numbers in each line, separated by semicolons, as in the example
    data = []
    with open(filepath, 'r') as data_file:
        for line in data_file:
            line = line.strip()
            if line:
                values = line.split(';')
                if len(values) == 2:
                    try:
                        x_value = float(values[0])
                        y_value = float(values[1])
                        data.append((x_value, y_value))
                    except ValueError:
                        raise ValueError(f"Warning: Problem with line {line} in data file {filepath}.")
                else:
                    raise ValueError(f"Line {line} in data file {filepath} incorrectly formatted.")
    return data

def createPlotImage(datapoints, x_label, y_label):  #creates a matplotlib line plot and saves it as an image, returns its filepath
    plt.figure(figsize=(6,4))
    datapoints.sort()
    plt.plot([x for x, y in datapoints], [y for x, y in datapoints], marker='o') 
    plt.xlabel(x_label)
    plt.ylabel(y_label)
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
    plt.savefig(temp_file.name)
    plt.close()
    return temp_file.name

def addChartToPlotSlide(slide, datapoints, x_label, y_label):
    image_path = createPlotImage(datapoints, x_label, y_label)
    slide.shapes.add_picture(image_path, Inches(1), Inches(2))
    os.remove(image_path)

def addPlotSlide(presentation, title_text, data_path, x_label, y_label):
    slide_layout = presentation.slide_layouts[5] 
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = title_text
    datapoints = readDataFile(data_path)            
    addChartToPlotSlide(slide, datapoints, x_label, y_label)

def makePresentation(json_data):
    presentation = Presentation()
    try:
        json_root = json_data["presentation"]
    except KeyError:
        raise KeyError("Your JSON file has no top level object named 'presentation', despite it being required. Please check your JSON file.")
    
    for slide_data in json_root:
        try:
            slide_type = slide_data["type"]
            slide_title = slide_data["title"]
            slide_content = slide_data["content"]
        except KeyError as c:
            raise KeyError(f"In the JSON file, the slide object titled {slide_title} had no key {c}, despite it being required. Please check your JSON file.") 
        
        if slide_type == "title":
            addTitleSlide(presentation, slide_title, slide_content)
        elif slide_type == "text":
            addTextSlide(presentation, slide_title, slide_content)
            
        elif slide_type == "list":
            addListSlide(presentation, slide_title, slide_content)
                
        elif slide_type == "picture":
            addImgSlide(presentation, slide_title, slide_content)

        elif slide_type == "plot":
            try:
                slide_config = slide_data["configuration"]
                x_label = slide_config["x-label"]
                y_label = slide_config["y-label"]
            except:
                raise KeyError(f"In the JSON file, the plot slide object titled {slide_title} had no key {c}, despite it being required. Please check your JSON file.")
            addPlotSlide(presentation, slide_title, slide_content, x_label, y_label)

        else:
            raise ValueError("Incorrect slide type attribute ({slide_type}) given for the slide titled {slide_title}")

    return presentation



'''
MAIN LOOP OF THE PROGRAM
This is the part of the code where user interaction, error handling and logging happens.
The skeleton of this code is simple: asking the user for the JSON file, converting it using the makePresentation() function, then saving it into a .pptx file named by the user.
If an error is encountered in the course of this, the user is notified about its details, then the program starts over from requesting the JSON file.
Meanwhile, all important events are logged using the Python logging module and log entries are saved into the file pptx_maker.log.
'''

logging.basicConfig(filename='pptx_maker.log', filemode='a', format='%(name)s - %(levelname)s - %(message)s', level=logging.INFO)
print("This program makes presentations from JSON files.")
print("You will be asked to provide the JSON file to be summarized as a presentation.")
print("Make sure that the JSON file and other relevant files are in the program library.")

while True:
    filename_input = input(f"\n Enter the filename of the JSON file (without the .json extension) to be turned into a PPT: ")
    json_filename = filename_input + ".json"
    logging.info(f"Attempting to read {json_filename}.")
    try:
        with open(json_filename, 'r') as file:
            presentation_data = json.load(file)
    except FileNotFoundError:
        print(f"File {json_filename} not found. Please enter a valid filename. Restarting with a new file input request.")
        logging.error(f"File {json_filename} not found.")
        continue
    
    except json.JSONDecodeError as e:
        print(f"There was an issue interpreting your JSON file. Make sure the file is valid then start the process over. Restarting with a new file input request.")
        logging.error(f"JSON decoding error with {json_filename}.")
        continue

    try:
        presentation = makePresentation(presentation_data)        

    except FileNotFoundError as e:
        missing_file = e.filename
        print(f"The input file {missing_file} mentioned in your JSON source file was not found. Please check this then start the process over. Restarting with a new file input request.")
        logging.error(f"Source file {missing_file} not found.")
        continue

    except ValueError as e:
        print(f"In your JSON file, some keys were assigned invalid values. Please check the file then start the process over. Error details: {str(e)} Restarting with a new file input request.")
        logging.error(f"Value error encountered. Error message: {str(e)}")
        continue

    except KeyError as e:
        print(f"Your JSON file were missing required data. Please check the file, then start the process over. Error details: {str(e)} Restarting with a new file input request.")
        logging.error(f"Key error encountered. Error message: {str(e)}")
        continue

    except PermissionError as e:
        print(f"The program was denied permission to access file {e.filename}. Please make sure this program has the appropriate permissions, then start the process over. Restarting with a new file input request.")
        logging.error(f"Permission error encountered when trying to access {e.filename}")
        continue

    except TypeError as e:
        print(f"There were issues with some data you provided. Error details: {str(e)}. Please revise the mentioned data then start the process over. Restarting with a new file input request.")
        logging.error(f"Key error encountered. Error message: {str(e)}")
        continue

    print("Your JSON file has been successfully converted to a presentation.")
    
    try:
        ppt_filename_input = input(f"Please enter a filename for the .pptx file to be created from your JSON file: ")
        output_filename = ppt_filename_input + ".pptx"
        presentation.save(output_filename)
        logging.info(f"JSON input file {filename} succesfully converted, resulting presentation saved to {output_filename}.")
        input(f"The presentation has been saved into the file {output_filename}. Enter anything to exit the program.")
        break
    
    except PermissionError:
        print(f"There was a permission error when trying to save the presentation. Please make sure this program has the appropriate permissions, then start the process over. Restarting with a new file input request.")
        logging.error(f"Permission error when trying to save the presentation to {output_filename}.")
        continue


