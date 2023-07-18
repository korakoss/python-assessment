from pptx import Presentation
from pptx.util import Inches
import json
from pptx.chart.data import XySeriesData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE
import matplotlib.pyplot as plt
import tempfile


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
    body = slide.placeholders[1]
    title.text = title_text
    body.text = ""
    for item in list_json:
        level = item["level"]
        text = item["text"]
        body.text += f"\n{' ' * (4 * (level - 1))}• {text}"

def addImgSlide(presentation, title_text, img_path):  #TODO do alignment correctly, the img is covering the title
    slide_layout = presentation.slide_layouts[1]  
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    title.text = title_text
    content.text = ""
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
                        print(f"Warning: Ignoring line '{line}'. Invalid float values.")
                #else raise something maybe?
    return data

def createPlotImage(datapoints, x_label, y_label):  #TODO order the points so they are connected in the right order
    plt.figure(figsize=(6,4))
    plt.plot([x for x, y in datapoints], [y for x, y in datapoints], marker='o') 
    plt.xlabel(x_label)
    plt.ylabel(y_label)
    img_path = 'temp_plot.png'
    plt.savefig(img_path)
    plt.close()
    return img_path

def addChartToPlotSlide(slide, datapoints, x_label, y_label):
    image_path = createPlotImage(datapoints, x_label, y_label)
    slide.shapes.add_picture(image_path, Inches(1), Inches(2))

def addPlotSlide(presentation, title_text, data_path, x_label, y_label):
    slide_layout = presentation.slide_layouts[1]
    slide = presentation.slides.add_slide(slide_layout)
    title = slide.shapes.title
    content = slide.placeholders[1]
    title.text = title_text
    content.text = ""
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

