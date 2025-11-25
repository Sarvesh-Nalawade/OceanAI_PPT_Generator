# main_script.py

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.chart.data import CategoryChartData, ChartData
from pptx.dml.color import RGBColor

def create_title_slide(prs):
    """Creates the title slide (Slide 1)."""
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "The Evolution of Programming Languages"
    subtitle.text = "A Look at Popularity and Market Share of Java, Python, JavaScript, and C++"

def create_intro_slide(prs):
    """Creates the introduction slide (Slide 2)."""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    body = slide.placeholders[1]

    title.text = "Introduction: The Four Titans"
    tf = body.text_frame
    tf.text = "A brief overview of the languages that have shaped modern software development:"

    p1 = tf.add_paragraph(); p1.text = "Java: The enterprise-level, object-oriented language known for its platform independence."; p1.level = 1
    p2 = tf.add_paragraph(); p2.text = "Python: A high-level, versatile language praised for its simplicity and readability."; p2.level = 1
    p3 = tf.add_paragraph(); p3.text = "JavaScript: The ubiquitous language of the web, essential for front-end development."; p3.level = 1
    p4 = tf.add_paragraph(); p4.text = "C++: A powerful, high-performance language used for systems programming and gaming."; p4.level = 1

    # Add image placeholder
    left, top, width, height = Inches(10.5), Inches(2.5), Inches(4.5), Inches(4)
    slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    
    # Add caption for the placeholder
    caption_box = slide.shapes.add_textbox(left, top + height, width, Inches(0.5))
    caption_p = caption_box.text_frame.paragraphs[0]
    caption_p.text = "Image: Logos of Java, Python, JavaScript, and C++"
    caption_p.font.size = Pt(10)
    caption_p.font.italic = True

def create_java_slide(prs):
    """Creates the 'Java's Journey' slide (Slide 3)."""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    body = slide.placeholders[1]

    title.text = "Java's Journey: The Enterprise Workhorse"
    tf = body.text_frame
    tf.text = "Dominance in Enterprise Applications"

    p1 = tf.add_paragraph(); p1.text = "\"Write Once, Run Anywhere\" philosophy powered by the JVM."; p1.level = 1
    p2 = tf.add_paragraph(); p2.text = "Backbone of large-scale corporate systems and Android mobile development."; p2.level = 1
    p3 = tf.add_paragraph(); p3.text = "Known for stability, security, and a massive ecosystem (Spring, Maven)."; p3.level = 1
    p4 = tf.add_paragraph(); p4.text = "While its growth has slowed, it remains a critical and high-demand skill."; p4.level = 1
    
    # Add image placeholder
    left, top, width, height = Inches(1), Inches(5.5), Inches(4), Inches(3)
    slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    
    # Add caption
    caption_box = slide.shapes.add_textbox(left, top + height, width, Inches(0.5))
    caption_p = caption_box.text_frame.paragraphs[0]
    caption_p.text = "Image: Diagram of the Java Virtual Machine (JVM) architecture."
    caption_p.font.size = Pt(10)
    caption_p.font.italic = True

def create_python_slide_with_pie_chart(prs):
    """Creates the 'Rise of Python' slide with a Pie Chart (Slide 4)."""
    slide_layout = prs.slide_layouts[5]  # Title Only layout
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "The Rise of Python: The Versatile Powerhouse"

    # Add text content in a textbox
    left_text, top_text, width_text, height_text = Inches(0.5), Inches(1.5), Inches(7), Inches(5)
    txBox = slide.shapes.add_textbox(left_text, top_text, width_text, height_text)
    tf = txBox.text_frame
    tf.text = "From scripting to science, Python's growth is unparalleled:"
    p1 = tf.add_paragraph(); p1.text = "Simplicity and readability attract beginners and experts alike."; p1.level = 1
    p2 = tf.add_paragraph(); p2.text = "Dominant force in Data Science, Machine Learning, and AI."; p2.level = 1
    p3 = tf.add_paragraph(); p3.text = "Strong web frameworks like Django and Flask."; p3.level = 1
    p4 = tf.add_paragraph(); p4.text = "Vast library support (NumPy, Pandas, TensorFlow)."; p4.level = 1

    # Add Pie Chart
    chart_data = ChartData()
    chart_data.categories = ['Data Science/ML', 'Web Development', 'Automation', 'Other']
    chart_data.add_series('Primary Use Cases', (45, 25, 20, 10))

    x, y, cx, cy = Inches(8), Inches(2), Inches(7), Inches(5)
    graphic_frame = slide.shapes.add_chart(XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data)
    chart = graphic_frame.chart

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False

    chart.plots[0].has_data_labels = True
    data_labels = chart.plots[0].data_labels
    data_labels.number_format = '0"%"'
    data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    data_labels.font.size = Pt(14)
    chart.chart_title.text_frame.text = "Python's Primary Use Cases (2024)"

def create_javascript_slide_with_table(prs):
    """Creates the 'JavaScript's Dominance' slide with a Table (Slide 5)."""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    body = slide.placeholders[1]

    title.text = "JavaScript's Dominance: King of the Web"
    tf = body.text_frame
    tf.text = "The language that powers virtually every modern website:"
    p1 = tf.add_paragraph(); p1.text = "Evolved from a browser script to a full-stack solution with Node.js."; p1.level = 1
    p2 = tf.add_paragraph(); p2.text = "Vibrant ecosystem of frameworks like React, Angular, and Vue.js."; p2.level = 1

    # Add Table
    rows, cols = 4, 3
    left, top, width, height = Inches(1.5), Inches(4.0), Inches(13.0), Inches(3.0)
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    # Set column widths and write headings
    table.columns[0].width, table.columns[1].width, table.columns[2].width = Inches(3.0), Inches(3.0), Inches(7.0)
    table.cell(0, 0).text, table.cell(0, 1).text, table.cell(0, 2).text = 'Framework', 'Creator / Backer', 'Key Feature'

    # Populate table data
    frameworks = [
        ('React', 'Facebook', 'Component-based UI library, virtual DOM for performance.'),
        ('Angular', 'Google', 'Comprehensive MVC framework with two-way data binding.'),
        ('Vue.js', 'Evan You', 'Progressive framework, easy to learn and integrate.')
    ]
    for i, (fw, creator, feature) in enumerate(frameworks):
        table.cell(i + 1, 0).text, table.cell(i + 1, 1).text, table.cell(i + 1, 2).text = fw, creator, feature

    # Format table headers
    for i in range(cols):
        for para in table.cell(0, i).text_frame.paragraphs:
            para.font.bold = True

def create_cpp_slide(prs):
    """Creates 'The Legacy of C++' slide (Slide 6)."""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title, body = slide.shapes.title, slide.placeholders[1]
    
    title.text = "The Legacy of C++: The Performance King"
    tf = body.text_frame
    tf.text = "When speed and control are paramount, C++ remains undefeated:"
    p1 = tf.add_paragraph(); p1.text = "Direct memory management provides unparalleled performance."; p1.level = 1
    p2 = tf.add_paragraph(); p2.text = "The language of choice for AAA game engines, HFT, and embedded systems."; p2.level = 1
    p3 = tf.add_paragraph(); p3.text = "Modern C++ (C++11 and beyond) adds features for safety and ease of use."; p3.level = 1

    # Add image placeholder
    left, top, width, height = Inches(10.5), Inches(2.5), Inches(4.5), Inches(4)
    slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    
    # Add caption
    caption_box = slide.shapes.add_textbox(left, top + height, width, Inches(0.5))
    caption_p = caption_box.text_frame.paragraphs[0]
    caption_p.text = "Image: Screenshot of C++ code for a high-performance application."
    caption_p.font.size = Pt(10)
    caption_p.font.italic = True

def create_line_chart_slide(prs):
    """Creates the 'Comparative Popularity' slide with a Line Chart (Slide 7)."""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "Comparative Popularity Trends (TIOBE Index-like Data)"

    chart_data = CategoryChartData()
    chart_data.categories = ['2004', '2009', '2014', '2019', '2024']
    chart_data.add_series('Java',       (20.7, 17.3, 16.5, 16.8, 10.4))
    chart_data.add_series('Python',     (2.6,  4.5,  5.8,  9.3,  15.4))
    chart_data.add_series('JavaScript', (1.5,  2.8,  4.2,  7.1,  8.5))
    chart_data.add_series('C++',        (12.2, 10.1, 7.4,  8.2,  11.9))

    x, y, cx, cy = Inches(1), Inches(1.5), Inches(14), Inches(6.5)
    graphic_frame = slide.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data)
    chart = graphic_frame.chart

    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.tick_labels.number_format = '0"%"'
    chart.category_axis.tick_labels.font.size = Pt(12)

def create_bar_chart_slide(prs):
    """Creates the 'Current Market Share' slide with a Bar Chart (Slide 8)."""
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    title.text = "Current Market Share & Demand (2024)"

    chart_data = CategoryChartData()
    chart_data.categories = ['Java', 'Python', 'JavaScript', 'C++']
    chart_data.add_series('Stack Overflow Survey %', (30.5, 43.8, 63.6, 20.5))

    x, y, cx, cy = Inches(2), Inches(1.5), Inches(12), Inches(6.5)
    graphic_frame = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)
    chart = graphic_frame.chart

    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    data_labels = plot.data_labels
    data_labels.number_format = '0.0"%"'
    data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    
    chart.value_axis.maximum_scale = 70.0
    chart.category_axis.tick_labels.font.size = Pt(14)

def create_conclusion_slide(prs):
    """Creates the conclusion slide (Slide 9)."""
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)
    title, body = slide.shapes.title, slide.placeholders[1]

    title.text = "Conclusion and Future Outlook"
    tf = body.text_frame
    tf.text = "Each language has a distinct and vital role in the tech ecosystem:"
    
    p1 = tf.add_paragraph(); p1.text = "Java remains the bedrock of enterprise, though facing competition."; p1.level = 1
    p2 = tf.add_paragraph(); p2.text = "Python's growth trajectory in AI/ML shows no signs of slowing."; p2.level = 1
    p3 = tf.add_paragraph(); p3.text = "JavaScript is irreplaceable on the web, with its ecosystem constantly evolving."; p3.level = 1
    p4 = tf.add_paragraph(); p4.text = "C++ continues to be the undisputed champion for performance-critical applications."; p4.level = 1
    
    p5 = tf.add_paragraph()
    p5.text = "The future is polyglot: developers will increasingly need multiple languages to succeed."
    p5.level = 0
    p5.font.bold = True

def main():
    """Main function to generate the presentation."""
    prs = Presentation()
    # Use 16:9 aspect ratio
    prs.slide_width = Inches(16)
    prs.slide_height = Inches(9)

    # Create all slides
    create_title_slide(prs)
    create_intro_slide(prs)
    create_java_slide(prs)
    create_python_slide_with_pie_chart(prs)
    create_javascript_slide_with_table(prs)
    create_cpp_slide(prs)
    create_line_chart_slide(prs)
    create_bar_chart_slide(prs)
    create_conclusion_slide(prs)

    # Save the presentation
    output_filename = 'output.pptx'
    prs.save(output_filename)
    print(f"Presentation '{output_filename}' created successfully.")

if __name__ == "__main__":
    main()