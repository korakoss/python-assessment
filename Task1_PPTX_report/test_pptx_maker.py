import unittest
from pptx import Presentation
from pptx_maker import readDataFile, addTitleSlide, addTextSlide, addListSlide, addImgSlide, addPlotSlide
import json

class TestPptxMaker(unittest.TestCase):
    def test_readDataFile(self):
        data = [
            (1.0, 2.0),
            (3.0, 4.0)
        ]
        file_path = 'data.dat'
        with open(file_path, 'w') as file:
            for line in data:
                string_line = [str(value) for value in line]
                file.write(';'.join(string_line) + '\n')
        read = readDataFile(file_path)
        self.assertEqual(read, data)

    def test_AddTitleSlide(self):
        presentation = Presentation()
        addTitleSlide(presentation, "Title", "Subtitle")
        slides = presentation.slides
        self.assertEqual(len(slides),1)
        slide = slides[0]  
        title = slide.shapes.title.text
        subtitle = slide.placeholders[1].text
        self.assertEqual(title, "Title")
        self.assertEqual(subtitle, "Subtitle")

    def test_AddTextSlide(self):
        presentation = Presentation()
        addTextSlide(presentation, "Title", "Long text content")
        slides = presentation.slides
        self.assertEqual(len(slides),1)
        slide = slides[0]  
        title = slide.shapes.title.text
        content = slide.placeholders[1].text
        self.assertEqual(title, "Title")
        self.assertEqual(content, "Long text content")

    def test_AddListSlide(self):
        presentation = Presentation()
        json_string = '[{ "level" : 1, "text" : "The Level 1 Text"},{ "level" : 2, "text" : "The Level 2 Text"},{ "level" : 2, "text" : "The Level 2 Text"},{ "level" : 1, "text" : "The Level 1 Text"}]'
        json_data = json.loads(json_string)
        addListSlide(presentation, "Title", json_data)
        slides = presentation.slides
        self.assertEqual(len(slides),1)
        slide = slides[0]  
        title = slide.shapes.title.text
        content = slide.placeholders[1].text
        self.assertEqual(title, "Title")
        print(content)

    def test_AddImgSlide(self):  #the result of the tested function can be evaluated visually
        presentation = Presentation()
        addImgSlide(presentation, "Title", "test.png")
        slides = presentation.slides
        self.assertEqual(len(slides),1)
        slide = slides[0]  
        title = slide.shapes.title.text
        content = slide.placeholders[1].text
        self.assertEqual(title, "Title")
        presentation.save("img_test_pres.pptx")

    def test_AddPlotSlide(self):
        presentation = Presentation()
        addPlotSlide(presentation, "Title", "data.dat", "xlabel", "ylabel")
        slides = presentation.slides
        self.assertEqual(len(slides),1)
        slide = slides[0]  
        title = slide.shapes.title.text
        content = slide.placeholders[1].text
        self.assertEqual(title, "Title")
        presentation.save("plot_test_pres.pptx")


if __name__ == "__main__":
    unittest.main()
