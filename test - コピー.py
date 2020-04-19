import os  
from comtypes import client  
  
if __name__ == "__main__":  
    application = client.CreateObject("Powerpoint.Application")  
    application.Visible = True  
    presentation = application.Presentations.Add()  
    slides = presentation.Slides     
  
    # Reference * https://www.relief.jp/docs/powerpoint-vba-make-new-presentation.html  
    # ppLayoutBlank = 12  
    slide = slides.Add(len(slides) + 1, 12)   
  
    # Reference:  
    # * https://docs.microsoft.com/en-us/office/vba/api/powerpoint.shapes.addtextbox  
    # * https://docs.microsoft.com/en-us/office/vba/api/office.msotextorientation  
  
    shape = slide.shapes.AddTextBox(Orientation=1, Left=100, Top=100, Width=500, Height=50)  
    shape.TextFrame.TextRange.Font.Color.RGB = 0x000000  
    shape.TextFrame.TextRange.text = "Hello, Powerpoint."  
    shape.TextEffect.FontSize = 36  
  
    current_folder = os.getcwd()  
    presentation.SaveAs(os.path.join(current_folder, "sample.pptx"))  
    