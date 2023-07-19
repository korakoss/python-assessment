from pptx import Presentation
from pptx.util import Inches, Pt
import json
from pptx.chart.data import XySeriesData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE
import matplotlib.pyplot as plt
import tempfile
import os
import logging

#TODO rest of exception handling
#TODO commenting if needed
#TODO handle better when datafile not formatted correctly

def addTitleSlide(presentation, title_text, subtitle_text):
    '''Each of the addSomeSlide functions require a presentation arg and some other args depending on slide type.
    The fuctions add a slide to the presentation provided in the first argument. The contents of the slide are determined by the other args'''
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

def addListSlide(presentation, title_text, list_json): #TODO: raise errors if needed (eg wrong level numbers etc
    slide_layout = presentation.slide_layouts[1]
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = title_text
    content = slide.placeholders[1]
    content.text = ""
    for item in list_json:
        level = item["level"]
        text = item["text"]
        p = content.text_frame.add_paragraph()
        p.text = text
        p.level = level - 1 

        
def addImgSlide(presentation, title_text, img_path):
    slide_layout = presentation.slide_layouts[5] #choosing a "title only" layout to not have a textbox on the slide  
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = title_text
    slide.shapes.add_picture(img_path, Inches(1), Inches(2))

def readDataFile(filepath):  #TODO better exception handling. #Reads the data for plot slides. I assumed that the data consists of pairs of numbers in each line, separated by semicolons, as in the example
    data = []
    with open(filepath, 'r') as file:
        for line in file:
            line = line.strip()
            if line:
                values = line.split(';')
                if len(values) == 2:
                    try:
                        value1 = float(values[0])
                        value2 = float(values[1])
                        data.append((value1, value2))
                    except ValueError:
                        print(f"Warning: Problem with line {line} in data file {filepath}.")
                #else raise something maybe?
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
    slide_layout = presentation.slide_layouts[5] #"title only" layout to avoid adding an empty textbox by defaul
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = title_text
    datapoints = readDataFile(data_path)            
    addChartToPlotSlide(slide, datapoints, x_label, y_label)

def makePresentation(json_data):
    presentation = Presentation()
    for slide_data in json_data["presentation"]:
        slide_type = slide_data["type"]
        slide_title = slide_data["title"]
        slide_content = slide_data["content"]

        if slide_type == "title":
            addTitleSlide(presentation, slide_title, slide_content)
        elif slide_type == "text":
            addTextSlide(presentation, slide_title, slide_content)
            
        elif slide_type == "list":
            addListSlide(presentation, slide_title, slide_content)
                
        elif slide_type == "picture":
            addImgSlide(presentation, slide_title, slide_content)

        elif slide_type == "plot":
            slide_config = slide_data["configuration"]
            x_label = slide_config["x-label"]
            y_label = slide_config["y-label"]
            addPlotSlide(presentation, slide_title, slide_content, x_label, y_label)

    return presentation

#MAIN PART OF PROGRAM
logging.basicConfig(filename='pptx_maker.log', filemode='a', format='%(name)s - %(levelname)s - %(message)s', level=logging.INFO)
print("This program makes presentations from JSON files.")
print("You will be asked to provide the JSON file to be summarized as a presentation.")
print("Make sure that the JSON file and other relevant files are in the program library.")

while True:
    json_inp = input("Enter the filename of the JSON file (without the .json extension): ")
    filename = json_inp + ".json"
    logging.info(f"Attempting to read {filename}.")
    try:
        with open(filename, 'r') as file:
            presentation_data = json.load(file)
    except FileNotFoundError:
        print(f"File {filename} not found. Please enter a valid filename. \n")
        logging.error(f"File {filename} not found.")
    
    except json.JSONDecodeError as e:
        print(f"There was an issue interpreting your JSON file. Make sure the file is valid.")
        logging.error(f"JSON decoding error with {filename}.")

    try:
        presentation = makePresentation(presentation_data)        

    except FileNotFoundError as e:
        missing_file = e.filename
        print(f"The input file {missing_file} mentioned in your JSON source file was not found.")
        logging.error(f"Source file {missing_file} not found.")

    except ValueError:
        print("There was an issue, likely when interpreting the plot data. Make sure the data is in the correct format.")
        logging.error("Value error encountered.")
    

    print("Your JSON file has been successfully converted.")
    
    try:
        ppt_inp = input(f"Please enter a filename for the .pptx file to be created from your JSON file: ")
        output_filename = ppt_inp + ".pptx"
        presentation.save(output_filename)
        logging.info(f"JSON input file {filename} succesfully converted, resulting presentation saved to {output_filename}.")
        input(f"The presentation has been saved into the file {output_filename}. Enter anything to exit the program.")
        break
    
    except PermissionError:
        print(f"There was a permission error when trying to save the presentation. Make sure you have write access to the directory.")
        logging.error(f"Permission error when trying to save the presentation to {output_filename}.")


