# doc-to-ppt
copying the content in the doc file to ppt
Here's a step-by-step process you can follow: 

1.Install the required libraries by running the following command in your command prompt or terminal:

**pip install python-pptx python-docx**

2.Import the necessary modules in your Python script:

**from pptx import Presentation
   from docx import Document**

3.Load the DOC file and create a PPT presentation:

# Load the DOC file
**doc_file = Document("path/to/input.docx")**

# Create a new PPT presentation
**presentation = Presentation()**

4.Iterate through the paragraphs or sections in the DOC file and add them to the PPT slides:

for paragraph in doc_file.paragraphs:
    # Create a new slide
    **slide = presentation.slides.add_slide(presentation.slide_layouts[1])**

    # Add the paragraph content to the slide
    **content = slide.shapes.add_textbox().text_frame
       content.text = paragraph.text**

5.Save the modified PPT presentation to a file:

**presentation.save("path/to/output.pptx")**


Make sure to replace "path/to/input.docx" with the actual path to your input DOC file and "path/to/output.pptx" with the desired path for the output PPT file.

Please note that this basic script assumes that each paragraph in the DOC file should be added as a separate slide in the PPT presentation. You may need to modify the code according to your specific requirements.

Remember to handle exceptions and perform any necessary error checking to ensure the smooth execution of the script.
