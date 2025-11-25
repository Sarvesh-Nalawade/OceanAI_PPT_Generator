# Getting Started — python-pptx 1.0.0 documentation
A quick way to get started is by trying out some of the examples below to get a feel for how to use python-pptx.

The API documentation can help you with the fine details of calling signatures and behaviors.

Hello World! example
-----

!../_images/hello-world.png
```
from pptx import Presentation

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "Hello, World!"
subtitle.text = "python-pptx was here!"

prs.save('test.pptx')
```

* * *

Bullet slide example
------

!../_images/bullet-slide.png
```
from pptx import Presentation

prs = Presentation()
bullet_slide_layout = prs.slide_layouts[1]

slide = prs.slides.add_slide(bullet_slide_layout)
shapes = slide.shapes

title_shape = shapes.title
body_shape = shapes.placeholders[1]

title_shape.text = 'Adding a Bullet Slide'

tf = body_shape.text_frame
tf.text = 'Find the bullet slide layout'

p = tf.add_paragraph()
p.text = 'Use _TextFrame.text for first bullet'
p.level = 1

p = tf.add_paragraph()
p.text = 'Use _TextFrame.add_paragraph() for subsequent bullets'
p.level = 2

prs.save('test.pptx')
```

Not all shapes can contain text, but those that do always have at least one paragraph, even if that paragraph is empty and no text is visible within the shape. `_BaseShape.has_text_frame` can be used to determine whether a shape can contain text. (All shapes subclass `_BaseShape`.) When `_BaseShape.has_text_frame` is `True`, `_BaseShape.text_frame.paragraphs[0]` returns the first paragraph. The text of the first paragraph can be set using `text_frame.paragraphs[0].text`. As a shortcut, the writable properties `_BaseShape.text` and `_TextFrame.text` are provided to accomplish the same thing. Note that these last two calls delete all the shape’s paragraphs except the first one before setting the text it contains.

* * *

`add_textbox()` example
--------

!../_images/add-textbox.png
```
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)

left = top = width = height = Inches(1)
txBox = slide.shapes.add_textbox(left, top, width, height)
tf = txBox.text_frame

tf.text = "This is text inside a textbox"

p = tf.add_paragraph()
p.text = "This is a second paragraph that's bold"
p.font.bold = True

p = tf.add_paragraph()
p.text = "This is a third paragraph that's big"
p.font.size = Pt(40)

prs.save('test.pptx')
```

* * *

`add_picture()` example
--------

!../_images/add-picture.png
```
from pptx import Presentation
from pptx.util import Inches

img_path = 'monty-truth.png'

prs = Presentation()
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)

left = top = Inches(1)
pic = slide.shapes.add_picture(img_path, left, top)

left = Inches(5)
height = Inches(5.5)
pic = slide.shapes.add_picture(img_path, left, top, height=height)

prs.save('test.pptx')
```

* * *

`add_shape()` example
-----

!../_images/add-shape.png
```
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

prs = Presentation()
title_only_slide_layout = prs.slide_layouts[5]
slide = prs.slides.add_slide(title_only_slide_layout)
shapes = slide.shapes

shapes.title.text = 'Adding an AutoShape'

left = Inches(0.93)  # 0.93" centers this overall set of shapes
top = Inches(3.0)
width = Inches(1.75)
height = Inches(1.0)

shape = shapes.add_shape(MSO_SHAPE.PENTAGON, left, top, width, height)
shape.text = 'Step 1'

left = left + width - Inches(0.4)
width = Inches(2.0)  # chevrons need more width for visual balance

for n in range(2, 6):
    shape = shapes.add_shape(MSO_SHAPE.CHEVRON, left, top, width, height)
    shape.text = 'Step %d' % n
    left = left + width - Inches(0.4)

prs.save('test.pptx')
```

Constants representing each of the available auto shapes (like MSO\_SHAPE.ROUNDED\_RECT, MSO\_SHAPE.CHEVRON, etc.) are listed on the autoshape-types page.

* * *

`add_table()` example
-----

!../_images/add-table.png
```
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
title_only_slide_layout = prs.slide_layouts[5]
slide = prs.slides.add_slide(title_only_slide_layout)
shapes = slide.shapes

shapes.title.text = 'Adding a Table'

rows = cols = 2
left = top = Inches(2.0)
width = Inches(6.0)
height = Inches(0.8)

table = shapes.add_table(rows, cols, left, top, width, height).table

# set column widths
table.columns[0].width = Inches(2.0)
table.columns[1].width = Inches(4.0)

# write column headings
table.cell(0, 0).text = 'Foo'
table.cell(0, 1).text = 'Bar'

# write body cells
table.cell(1, 0).text = 'Baz'
table.cell(1, 1).text = 'Qux'

prs.save('test.pptx')
```

* * *

# Working with Presentations — python-pptx 1.0.0 documentation
python-pptx allows you to create new presentations as well as make changes to existing ones. Actually, it only lets you make changes to existing presentations; it’s just that if you start with a presentation that doesn’t have any slides, it feels at first like you’re creating one from scratch.

However, a lot of how a presentation looks is determined by the parts that are left when you delete all the slides, specifically the theme, the slide master, and the slide layouts that derive from the master. Let’s walk through it a step at a time using examples, starting with the two things you can do with a presentation, open it and save it.

Opening a presentation
-----

The simplest way to get started is to open a new presentation without specifying a file to open:
```
from pptx import Presentation

prs = Presentation()
prs.save('test.pptx')
```

This creates a new presentation from the built-in default template and saves it unchanged to a file named ‘test.pptx’. A couple things to note:

*   The so-called “default template” is actually just a PowerPoint file that doesn’t have any slides in it, stored with the installed python-pptx package. It’s the same as what you would get if you created a new presentation from a fresh PowerPoint install, a 4x3 aspect ratio presentation based on the “White” template. Well, except it won’t contain any slides. PowerPoint always adds a blank first slide by default.
*   You don’t need to do anything to it before you save it. If you want to see exactly what that template contains, just look in the ‘test.pptx’ file this creates.
*   We’ve called it a _template_, but in fact it’s just a regular PowerPoint file with all the slides removed. Actual PowerPoint template files (.potx files) are something a bit different. More on those later maybe, but you won’t need them to work with python-pptx.

### REALLY opening a presentation

Okay, so if you want any control at all to speak of over the final presentation, or if you want to change an existing presentation, you need to open one with a filename:
```
prs = Presentation('existing-prs-file.pptx')
prs.save('new-file-name.pptx')
```

Things to note:

*   You can open any PowerPoint 2007 or later file this way (.ppt files from PowerPoint 2003 and earlier won’t work). While you might not be able to manipulate all the contents yet, whatever is already in there will load and save just fine. The feature set is still being built out, so you can’t add or change things like Notes Pages yet, but if the presentation has them python-pptx is polite enough to leave them alone and smart enough to save them without actually understanding what they are.
*   If you use the same filename to open and save the file, python-pptx will obediently overwrite the original file without a peep. You’ll want to make sure that’s what you intend.

### Opening a ‘file-like’ presentation

python-pptx can open a presentation from a so-called _file-like_ object. It can also save to a file-like object. This can be handy when you want to get the source or target presentation over a network connection or from a database and don’t want to (or aren’t allowed to) fuss with interacting with the file system. In practice this means you can pass an open file or StringIO/BytesIO stream object to open or save a presentation like so:
```
f = open('foobar.pptx')
prs = Presentation(f)
f.close()

# or

with open('foobar.pptx') as f:
    source_stream = StringIO(f.read())
prs = Presentation(source_stream)
source_stream.close()
...
target_stream = StringIO()
prs.save(target_stream)
```

Okay, so you’ve got a presentation open and are pretty sure you can save it somewhere later. Next step is to get a slide in there …

# Working with Slides — python-pptx 1.0.0 documentation
Every slide in a presentation is based on a slide layout. Not surprising then that you have to specify which slide layout to use when you create a new slide. Let’s take a minute to understand a few things about slide layouts that we’ll need so the slide we add looks the way we want it to.

Slide layout basics
-----

A slide layout is like a template for a slide. Whatever is on the slide layout “shows through” on a slide created with it and formatting choices made on the slide layout are inherited by the slide. This is an important feature for getting a professional-looking presentation deck, where all the slides are formatted consistently. Each slide layout is based on the slide master in a similar way, so you can make presentation-wide formatting decisions on the slide master and layout-specific decisions on the slide layouts. There can actually be multiple slide masters, but I’ll pretend for now there’s only one. Usually there is.

The presentation themes that come with PowerPoint have about nine slide layouts, with names like _Title_, _Title and Content_, _Title Only_, and _Blank_. Each has zero or more placeholders (mostly not zero), preformatted areas into which you can place a title, multi-level bullets, an image, etc. More on those later.

The slide layouts in a standard PowerPoint theme always occur in the same sequence. This allows content from one deck to be pasted into another and be connected with the right new slide layout:

*   Title (presentation title slide)
*   Title and Content
*   Section Header (sometimes called Segue)
*   Two Content (side by side bullet textboxes)
*   Comparison (same but additional title for each side by side content box)
*   Title Only
*   Blank
*   Content with Caption
*   Picture with Caption

In python-pptx, these are `prs.slide_layouts[0]` through `prs.slide_layouts[8]`. However, there’s no rule they have to appear in this order, it’s just a convention followed by the themes provided with PowerPoint. If the deck you’re using as your template has different slide layouts or has them in a different order, you’ll have to work out the slide layout indices for yourself. It’s pretty easy. Just open it up in Slide Master view in PowerPoint and count down from the top, starting at zero.

Now we can get to creating a new slide.

Adding a slide
-----

Let’s use the Title and Content slide layout; a lot of slides do:
```
SLD_LAYOUT_TITLE_AND_CONTENT = 1

prs = Presentation()
slide_layout = prs.slide_layouts[SLD_LAYOUT_TITLE_AND_CONTENT]
slide = prs.slides.add_slide(slide_layout)
```

A few things to note:

*   Using a “constant” value like `SLD_LAYOUT_TITLE_AND_CONTENT` is up to you. If you’re creating many slides it can be handy to have constants defined so a reader can more easily make sense of what you’re doing. There isn’t a set of these built into the package because they can’t be assured to be right for the starting deck you’re using.
*   `prs.slide_layouts` is the collection of slide layouts contained in the presentation and has list semantics, at least for item access which is about all you can do with that collection at the moment. Using `prs` for the Presentation instance is purely conventional, but I like it and use it consistently.
*   `prs.slides` is the collection of slides in the presentation, also has list semantics for item access, and len() works on it. Note that the method to add the slide is on the slide collection, not the presentation. The `add_slide()` method appends the new slide to the end of the collection. At the time of writing it’s the only way to add a slide, but sooner or later I expect someone will want to insert one in the middle, and when they post a feature request for that I expect I’ll add an `insert_slide(idx, ...)` method.

Doing other things with slides
-----

Right now, adding a slide is the only operation on the slide collection. On the backlog at the time of writing is deleting a slide and moving a slide to a different position in the list. Copying a slide from one presentation to another turns out to be pretty hard to get right in the general case, so that probably won’t come until more of the backlog is burned down.

Up next …
-----

Ok, now that we have a new slide, let’s talk about how to put something on it …

# Understanding Shapes — python-pptx 1.0.0 documentation
Pretty much anything on a slide is a shape; the only thing I can think of that can appear on a slide that’s not a shape is a slide background. There are between six and ten different types of shape, depending how you count. I’ll explain some of the general shape concepts you’ll need to make sense of how to work with them and then we’ll jump right into working with the specific types.

Technically there are six and only six different types of shapes that can be placed on a slide:

auto shape

This is a regular shape, like a rectangle, an ellipse, or a block arrow. They come in a large variety of preset shapes, in the neighborhood of 180 different ones. An auto shape can have a fill and an outline, and can contain text. Some auto shapes have adjustments, the little yellow diamonds you can drag to adjust how round the corners of a rounded rectangle are for example. A text box is also an autoshape, a rectangular one, just by default without a fill and without an outline.

picture

A raster image, like a photograph or clip art is referred to as a _picture_ in PowerPoint. It’s its own kind of shape with different behaviors than an autoshape. Note that an auto shape can have a picture fill, in which an image “shows through” as the background of the shape instead of a fill color or gradient. That’s a different thing. But cool.

graphic frame

This is the technical name for the container that holds a table, a chart, a smart art diagram, or media clip. You can’t add one of these by itself, it just shows up in the file when you add a graphical object. You probably won’t need to know anything more about these.

group shape

In PowerPoint, a set of shapes can be _grouped_, allowing them to be selected, moved, resized, and even filled as a unit. When you group a set of shapes a group shape gets created to contain those member shapes. You can’t actually see these except by their bounding box when the group is selected.

line/connector

Lines are different from auto shapes because, well, they’re linear. Some lines can be connected to other shapes and stay connected when the other shape is moved. These aren’t supported yet either so I don’t know much more about them. I’d better get to these soon though, they seem like they’d be very handy.

content part

I actually have only the vaguest notion of what these are. It has something to do with embedding “foreign” XML like SVG in with the presentation. I’m pretty sure PowerPoint itself doesn’t do anything with these. My strategy is to ignore them. Working good so far.

As for real-life shapes, there are these nine types:

*   shape shapes – auto shapes with fill and an outline
*   text boxes – auto shapes with no fill and no outline
*   placeholders – auto shapes that can appear on a slide layout or master and be inherited on slides that use that layout, allowing content to be added that takes on the formatting of the placeholder
*   line/connector – as described above
*   picture – as described above
*   table – that row and column thing
*   chart – pie chart, line chart, etc.
*   smart art – not supported yet, although preserved if present
*   media clip – video or audio

Accessing the shapes on a slide
-----

Each slide has a _shape tree_ that holds its shapes. It’s called a tree because it’s hierarchical in the general case; a node in the shape tree can be a group shape which itself can contain shapes and has the same semantics as the shape tree. For most purposes the shape tree has list semantics. You gain access to it like so:

We’ll see a lot more of the shape tree in the next few sections.

Up next …
-----

Okay. That should be enough noodle work to get started. Let’s move on to working with AutoShapes.

# Working with AutoShapes — python-pptx 1.0.0 documentation
Auto shapes are regular shape shapes. Squares, circles, triangles, stars, that sort of thing. There are 182 different auto shapes to choose from. 120 of these have adjustment “handles” you can use to change the shape, sometimes dramatically.

Many shape types share a common set of properties. We’ll introduce many of them here because several of those shapes are just a specialized form of AutoShape.

Adding an auto shape
------

The following code adds a rounded rectangle shape, one inch square, and positioned one inch from the top-left corner of the slide:
```
from pptx.enum.shapes import MSO_SHAPE

shapes = slide.shapes
left = top = width = height = Inches(1.0)
shape = shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
)
```

See the MSO\_AUTO\_SHAPE\_TYPE enumeration page for a list of all 182 auto shape types.

Understanding English Metric Units
-----

In the prior example we set the position and dimension values to the expression `Inches(1.0)`. What���s that about?

Internally, PowerPoint stores length values in _English Metric Units_ (EMU). This term might be worth a quick Googling, but the short story is EMU is an integer unit of length, 914400 to the inch. Most lengths in Office documents are stored in EMU. 914400 has the great virtue that it is evenly divisible by a great many common factors, allowing exact conversion between inches and centimeters, for example. Being an integer, it can be represented exactly across serializations and across platforms.

As you might imagine, working directly in EMU is inconvenient. To make it easier, python-pptx provides a collection of value types to allow easy specification and conversion into convenient units:
```
>>> from pptx.util import Inches, Pt
>>> length = Inches(1)
>>> length
914400
>>> length.inches
1.0
>>> length.cm
2.54
>>> length.pt
72.0
>>> length = Pt(72)
>>> length
914400
```

More details are available in the API documentation for pptx.util

Shape position and dimensions
-----

All shapes have a position on their slide and have a size. In general, position and size are specified when the shape is created. Position and size can also be read from existing shapes and changed:
```
>>> from pptx.enum.shapes import MSO_SHAPE
>>> left = top = width = height = Inches(1.0)
>>> shape = shapes.add_shape(
>>>     MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
>>> )
>>> shape.left, shape.top, shape.width, shape.height
(914400, 914400, 914400, 914400)
>>> shape.left.inches
1.0
>>> shape.left = Inches(2.0)
>>> shape.left.inches
2.0
```

Fill
-----

AutoShapes have an outline around their outside edge. What appears within that outline is called the shape’s _fill_.

The most common type of fill is a solid color. A shape may also be filled with a gradient, a picture, a pattern (like cross-hatching for example), or may have no fill (transparent).

When a color is used, it may be specified as a specific RGB value or a color from the theme palette.

Because there are so many options, the API for fill is a bit complex. This code sets the fill of a shape to red:
```
>>> fill = shape.fill
>>> fill.solid()
>>> fill.fore_color.rgb = RGBColor(255, 0, 0)
```

This sets it to the theme color that appears as ‘Accent 1 - 25% Darker’ in the toolbar palette:
```
>>> from pptx.enum.dml import MSO_THEME_COLOR
>>> fill = shape.fill
>>> fill.solid()
>>> fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
>>> fill.fore_color.brightness = -0.25
```

This sets the shape fill to transparent, or ‘No Fill’ as it’s called in the PowerPoint UI:
```
>>> shape.fill.background()
```

As you can see, the first step is to specify the desired fill type by calling the corresponding method on fill. Doing so actually changes the properties available on the fill object. For example, referencing `.fore_color` on a fill object after calling its `.background()` method will raise an exception:
```
>>> fill = shape.fill
>>> fill.solid()
>>> fill.fore_color
<pptx.dml.color.ColorFormat object at 0x10ce20910>
>>> fill.background()
>>> fill.fore_color
Traceback (most recent call last):
  ...
TypeError: a transparent (background) fill has no foreground color
```

Line
-----

The outline of an AutoShape can also be formatted, including setting its color, width, dash (solid, dashed, dotted, etc.), line style (single, double, thick-thin, etc.), end cap, join type, and others. At the time of writing, color and width can be set using python-pptx:
```
>>> line = shape.line
>>> line.color.rgb = RGBColor(255, 0, 0)
>>> line.color.brightness = 0.5  # 50% lighter
>>> line.width = Pt(2.5)
```

Theme colors can be used on lines too:
```
>>> line.color.theme_color = MSO_THEME_COLOR.ACCENT_6
```

`Shape.line` has the attribute `.color`. This is essentially a shortcut for:
```
>>> line.fill.solid()
>>> line.fill.fore_color
```

This makes sense for line formatting because a shape outline is most frequently set to a solid color. Accessing the fill directly is required, for example, to set the line to transparent:
```
>>> line.fill.background()
```

### Line width

The shape outline also has a read/write width property:
```
>>> line.width
9525
>>> line.width.pt
0.75
>>> line.width = Pt(2.0)
>>> line.width.pt
2.0
```

Adjusting an autoshape
-----

Many auto shapes have adjustments. In PowerPoint, these show up as little yellow diamonds you can drag to change the look of the shape. They’re a little fiddly to work with via a program, but if you have the patience to get them right, you can achieve some remarkable effects with great precision.

### Shape Adjustment Concepts

There are a few concepts it’s worthwhile to grasp before trying to do serious work with adjustments.

First, adjustments are particular to a specific auto shape type. Each auto shape has between zero and eight adjustments. What each of them does is arbitrary and depends on the shape design.

Conceptually, adjustments are guides, in many ways like the light blue ones you can align to in the PowerPoint UI and other drawing apps. These don’t show, but they operate in a similar way, each defining an x or y value that part of the shape will align to, changing the proportions of the shape.

Adjustment values are large integers, each based on a nominal value of 100,000. The effective value of an adjustment is proportional to the width or height of the shape. So a value of 50,000 for an x-coordinate adjustment corresponds to half the width of the shape; a value of 75,000 for a y-coordinate adjustment corresponds to 3/4 of the shape height.

Adjustment values can be negative, generally indicating the coordinate is to the left or above the top left corner (origin) of the shape. Values can also be subject to limits, meaning their effective value cannot be outside a prescribed range. In practice this corresponds to a point not being able to extend beyond the left side of the shape, for example.

Spending some time fooling around with shape adjustments in PowerPoint is time well spent to build an intuitive sense of how they behave. You also might want to have `opc-diag` installed so you can look at the XML values that are generated by different adjustments as a head start on developing your adjustment code.

The following code formats a callout shape using its adjustments:
```
callout_sp = shapes.add_shape(
    MSO_SHAPE.LINE_CALLOUT_2_ACCENT_BAR, left, top, width, height
)

# get the callout line coming out of the right place
adjs = callout_sp.adjustments
adjs[0] = 0.5   # vert pos of junction in margin line, 0 is top
adjs[1] = 0.0   # horz pos of margin ln wrt shape width, 0 is left side
adjs[2] = 0.5   # vert pos of elbow wrt margin line, 0 is top
adjs[3] = -0.1  # horz pos of elbow wrt shape width, 0 is margin line
adjs[4] = 3.0   # vert pos of line end wrt shape height, 0 is top
a5 = adjs[3] - (adjs[4] - adjs[0]) * height/width
adjs[5] = a5    # horz pos of elbow wrt shape width, 0 is margin line

# rotate 45 degrees counter-clockwise
callout_sp.rotation = -45.0
```

# Understanding placeholders — python-pptx 1.0.0 documentation
Intuitively, a placeholder is a pre-formatted container into which content can be placed. By providing pre-set formatting to its content, it places many of the formatting choices in the hands of the template designer while allowing the end-user to concentrate on the actual content. This speeds the presentation development process while encouraging visual consistency in slides created from the same template.

While their typical end-user behaviors are relatively simple, the structures that support their operation are more complex. This page is for those who want to better understand the architecture of the placeholder subsystem and perhaps be less prone to confusion at its sometimes puzzling behavior. If you don’t care why they work and just want to know how to work with them, you may want to skip forward to the following page Working with placeholders.

A placeholder is a shape
-----

Placeholders are an orthogonal category of shape, which is to say multiple shape types can be placeholders. In particular, the auto shape (p:sp element), picture (p:pic element), and graphic frame (p:graphicFrame) shape types can be a placeholder. The group shape (p:grpSp), connector (p:cxnSp), and content part (p:contentPart) shapes cannot be a placeholder. A graphic frame placeholder can contain a table, a chart, or SmartArt.

Placeholder types
-----

There are 18 types of placeholder.

Title, Center Title, Subtitle, Body

These placeholders typically appear on a conventional “word chart” containing text only, often organized as a title and a series of bullet points. All of these placeholders can accept text only.

Content

This multi-purpose placeholder is the most commonly used for the body of a slide. When unpopulated, it displays 6 buttons to allow insertion of a table, a chart, SmartArt, a picture, clip art, or a media clip.

Picture, Clip Art

These both allow insertion of an image. The insert button on a clip art placeholder brings up the clip art gallery rather than an image file chooser, but otherwise these behave the same.

Chart, Table, Smart Art

These three allow the respective type of rich graphical content to be inserted.

Media Clip

Allows a video or sound recording to be inserted.

Date, Footer, Slide Number

These three appear on most slide masters and slide layouts, but do not behave as most users would expect. These also commonly appear on the Notes Master and Handout Master.

Header

Only valid on the Notes Master and Handout Master.

Vertical Body, Vertical Object, Vertical Title

Used with vertically oriented languages such as Japanese.

Unpopulated vs. populated
-----

A placeholder on a slide can be empty or filled. This is most evident with a picture placeholder. When unpopulated, a placeholder displays customizable prompt text. A rich content placeholder will also display one or more content insertion buttons when empty.

A text-only placeholder enters “populated” mode when the first character of text is entered and returns to “unpopulated” mode when the last character of text is removed. A rich-content placeholder enters populated mode when content such as a picture is inserted and returns to unpopulated mode when that content is deleted. In order to delete a populated placeholder, the shape must be deleted _twice_. The first delete removes the content and restores the placeholder to unpopulated mode. An additional delete will remove the placeholder itself. A deleted placeholder can be restored by reapplying the layout.

Placholders inherit
-----

A placeholder appearing on a slide is only part of the overall placeholder mechanism. Placeholder behavior requires three different categories of placeholder shape; those that exist on a slide master, those on a slide layout, and those that ultimately appear on a slide in a presentation.

These three categories of placeholder participate in a property inheritance hierarchy, either as an inheritor, an inheritee, or both. Placeholder shapes on masters are inheritees only. Conversely placeholder shapes on slides are inheritors only. Placeholders on slide layouts are both, a possible inheritor from a slide master placeholder and an inheritee to placeholders on slides linked to that layout.

A layout inherits from its master differently than a slide inherits from its layout. A layout placeholder inherits from the master placeholder sharing the same type. A slide placeholder inherits from the layout placeholder having the same idx value.

In general, all formatting properties are inherited from the “parent” placeholder. This includes position and size as well as fill, line, and font. Any directly applied formatting overrides the corresponding inherited value. Directly applied formatting can be removed be reapplying the layout.

Glossary
-----

placeholder shape

A shape on a slide that inherits from a layout placeholder.

layout placeholder

a shorthand name for the placeholder shape on the slide layout from which a particular placeholder on a slide inherits shape properties

master placeholder

the placeholder shape on the slide master which a layout placeholder inherits from, if any.

---

# Working with placeholders — python-pptx 1.0.0 documentation
Placeholders can make adding content a lot easier. If you’ve ever added a new textbox to a slide from scratch and noticed how many adjustments it took to get it the way you wanted you understand why. The placeholder is in the right position with the right font size, paragraph alignment, bullet style, etc., etc. Basically you can just click and type in some text and you’ve got a slide.

A placeholder can be also be used to place a rich-content object on a slide. A picture, table, or chart can each be inserted into a placeholder and so take on the position and size of the placeholder, as well as certain of its formatting attributes.

Access a placeholder
------

Every placeholder is also a shape, and so can be accessed using the `shapes` property of a slide. However, when looking for a particular placeholder, the `placeholders` property can make things easier.

The most reliable way to access a known placeholder is by its `idx` value. The `idx` value of a placeholder is the integer key of the slide layout placeholder it inherits properties from. As such, it remains stable throughout the life of the slide and will be the same for any slide created using that layout.

It’s usually easy enough to take a look at the placeholders on a slide and pick out the one you want:
```
>>> prs = Presentation()
>>> slide = prs.slides.add_slide(prs.slide_layouts[8])
>>> for shape in slide.placeholders:
...     print('%d %s' % (shape.placeholder_format.idx, shape.name))
...
0  Title 1
1  Picture Placeholder 2
2  Text Placeholder 3
```

… then, having the known index in hand, to access it directly:
```
>>> slide.placeholders[1]
<pptx.parts.slide.PicturePlaceholder object at 0x10d094590>
>>> slide.placeholders[2].name
'Text Placeholder 3'
```

Note

Item access on the placeholders collection is like that of a dictionary rather than a list. While the key used above is an integer, the lookup is on idx values, not position in a sequence. If the provided value does not match the idx value of one of the placeholders, `KeyError` will be raised. idx values are not necessarily contiguous.

In general, the `idx` value of a placeholder from a built-in slide layout (one provided with PowerPoint) will be between 0 and 5. The title placeholder will always have `idx` 0 if present and any other placeholders will follow in sequence, top to bottom and left to right. A placeholder added to a slide layout by a user in PowerPoint will receive an `idx` value starting at 10.

Identify and Characterize a placeholder
-----

A placeholder behaves differently that other shapes in some ways. In particular, the value of its `shape_type` attribute is unconditionally `MSO_SHAPE_TYPE.PLACEHOLDER` regardless of what type of placeholder it is or what type of content it contains:
```
>>> prs = Presentation()
>>> slide = prs.slides.add_slide(prs.slide_layouts[8])
>>> for shape in slide.shapes:
...     print('%s' % shape.shape_type)
...
PLACEHOLDER (14)
PLACEHOLDER (14)
PLACEHOLDER (14)
```

To find out more, it’s necessary to inspect the contents of the placeholder’s `placeholder_format` attribute. All shapes have this attribute, but accessing it on a non-placeholder shape raises `ValueError`. The `is_placeholder` attribute can be used to determine whether a shape is a placeholder:
```
>>> for shape in slide.shapes:
...     if shape.is_placeholder:
...         phf = shape.placeholder_format
...         print('%d, %s' % (phf.idx, phf.type))
...
0, TITLE (1)
1, PICTURE (18)
2, BODY (2)
```

Another way a placeholder acts differently is that it inherits its position and size from its layout placeholder. This inheritance is overridden if the position and size of a placeholder are changed.

Insert content into a placeholder
-----

Certain placeholder types have specialized methods for inserting content. In the current release, the picture, table, and chart placeholders have content insertion methods. Text can be inserted into title and body placeholders in the same way text is inserted into an auto shape.

### `PicturePlaceholder.insert_picture()`

The picture placeholder has an `insert_picture()` method:
```
>>> prs = Presentation()
>>> slide = prs.slides.add_slide(prs.slide_layouts[8])
>>> placeholder = slide.placeholders[1]  # idx key, not position
>>> placeholder.name
'Picture Placeholder 2'
>>> placeholder.placeholder_format.type
PICTURE (18)
>>> picture = placeholder.insert_picture('my-image.png')
```

Note

A reference to a picture placeholder becomes invalid after its `insert_picture()` method is called. This is because the process of inserting a picture replaces the original p:sp XML element with a new p:pic element containing the picture. Any attempt to use the original placeholder reference after the call will raise `AttributeError`. The new placeholder is the return value of the `insert_picture()` call and may also be obtained from the placeholders collection using the same idx key.

A picture inserted in this way is stretched proportionately and cropped to fill the entire placeholder. Best results are achieved when the aspect ratio of the source image and placeholder are the same. If the picture is taller in aspect than the placeholder, its top and bottom are cropped evenly to fit. If it is wider, its left and right sides are cropped evenly. Cropping can be adjusted using the crop properties on the placeholder, such as `crop_bottom`.

### `TablePlaceholder.insert_table()`

The table placeholder has an `insert_table()` method. The built-in template has no layout containing a table placeholder, so this example assumes a starting presentation named `having-table-placeholder.pptx` having a table placeholder with idx 10 on its second slide layout:
```
>>> prs = Presentation('having-table-placeholder.pptx')
>>> slide = prs.slides.add_slide(prs.slide_layouts[1])
>>> placeholder = slide.placeholders[10]  # idx key, not position
>>> placeholder.name
'Table Placeholder 1'
>>> placeholder.placeholder_format.type
TABLE (12)
>>> graphic_frame = placeholder.insert_table(rows=2, cols=2)
>>> table = graphic_frame.table
>>> len(table.rows), len(table.columns)
(2, 2)
```

A table inserted in this way has the position and width of the original placeholder. Its height is proportional to the number of rows.

Like all rich-content insertion methods, a reference to a table placeholder becomes invalid after its `insert_table()` method is called. This is because the process of inserting rich content replaces the original p:sp XML element with a new element, a p:graphicFrame in this case, containing the rich-content object. Any attempt to use the original placeholder reference after the call will raise `AttributeError`. The new placeholder is the return value of the `insert_table()` call and may also be obtained from the placeholders collection using the original idx key, 10 in this case.

Note

The return value of the `insert_table()` method is a `PlaceholderGraphicFrame` object, which has all the properties and methods of a `GraphicFrame` object along with those specific to placeholders. The inserted table is contained in the graphic frame and can be obtained using its `table` property.

### `ChartPlaceholder.insert_chart()`

The chart placeholder has an `insert_chart()` method. The presentation template built into python-pptx has no layout containing a chart placeholder, so this example assumes a starting presentation named `having-chart-placeholder.pptx` having a chart placeholder with idx 10 on its second slide layout:
```
>>> from pptx.chart.data import ChartData
>>> from pptx.enum.chart import XL_CHART_TYPE

>>> prs = Presentation('having-chart-placeholder.pptx')
>>> slide = prs.slides.add_slide(prs.slide_layouts[1])

>>> placeholder = slide.placeholders[10]  # idx key, not position
>>> placeholder.name
'Chart Placeholder 9'
>>> placeholder.placeholder_format.type
CHART (12)

>>> chart_data = ChartData()
>>> chart_data.categories = ['Yes', 'No']
>>> chart_data.add_series('Series 1', (42, 24))

>>> graphic_frame = placeholder.insert_chart(XL_CHART_TYPE.PIE, chart_data)
>>> chart = graphic_frame.chart
>>> chart.chart_type
PIE (5)
```

A chart inserted in this way has the position and size of the original placeholder.

Note the return value from `insert_chart()` is a `PlaceholderGraphicFrame` object, not the chart itself. A `PlaceholderGraphicFrame` object has all the properties and methods of a `GraphicFrame` object along with those specific to placeholders. The inserted chart is contained in the graphic frame and can be obtained using its `chart` property.

Like all rich-content insertion methods, a reference to a chart placeholder becomes invalid after its `insert_chart()` method is called. This is because the process of inserting rich content replaces the original p:sp XML element with a new element, a p:graphicFrame in this case, containing the rich-content object. Any attempt to use the original placeholder reference after the call will raise `AttributeError`. The new placeholder is the return value of the `insert_chart()` call and may also be obtained from the placeholders collection using the original idx key, 10 in this case.

Setting the slide title
-----

Almost all slide layouts have a title placeholder, which any slide based on the layout inherits when the layout is applied. Accessing a slide’s title is a common operation and there’s a dedicated attribute on the shape tree for it:
```
title_placeholder = slide.shapes.title
title_placeholder.text = 'Air-speed Velocity of Unladen Swallows'
```

# Working with text — python-pptx 1.0.0 documentation
Auto shapes and table cells can contain text. Other shapes can’t. Text is always manipulated the same way, regardless of its container.

Text exists in a hierarchy of three levels:

*   `Shape.text_frame`
*   `TextFrame.paragraphs`
*   `_Paragraph.runs`

All the text in a shape is contained in its _text frame_. A text frame has vertical alignment, margins, wrapping and auto-fit behavior, a rotation angle, some possible 3D visual features, and can be set to format its text into multiple columns. It also contains a sequence of paragraphs, which always contains at least one paragraph, even when empty.

A paragraph has line spacing, space before, space after, available bullet formatting, tabs, outline/indentation level, and horizontal alignment. A paragraph can be empty, but if it contains any text, that text is contained in one or more runs.

A run exists to provide character level formatting, including font typeface, size, and color, an optional hyperlink target URL, bold, italic, and underline styles, strikethrough, kerning, and a few capitalization styles like all caps.

Let’s run through these one by one. Only features available in the current release are shown.

Accessing the text frame
-----

As mentioned, not all shapes have a text frame. So if you’re not sure and you don’t want to catch the possible exception, you’ll want to check before attempting to access it:
```
for shape in slide.shapes:
    if not shape.has_text_frame:
        continue
    text_frame = shape.text_frame
    # do things with the text frame
    ...
```

Accessing paragraphs
------

A text frame always contains at least one paragraph. This causes the process of getting multiple paragraphs into a shape to be a little clunkier than one might like. Say for example you want a shape with three paragraphs:
```
paragraph_strs = [
    'Egg, bacon, sausage and spam.',
    'Spam, bacon, sausage and spam.',
    'Spam, egg, spam, spam, bacon and spam.'
]

text_frame = shape.text_frame
text_frame.clear()  # remove any existing paragraphs, leaving one empty one

p = text_frame.paragraphs[0]
p.text = paragraph_strs[0]

for para_str in paragraph_strs[1:]:
    p = text_frame.add_paragraph()
    p.text = para_str
```

Adding text
-----

Only runs can actually contain text. Assigning a string to the `.text` attribute on a shape, text frame, or paragraph is a shortcut method for placing text in a run contained by those objects. The following two snippets produce the same result:
```
shape.text = 'foobar'

# is equivalent to ...

text_frame = shape.text_frame
text_frame.clear()
p = text_frame.paragraphs[0]
run = p.add_run()
run.text = 'foobar'
```

Applying text frame-level formatting
-----

The following produces a shape with a single paragraph, a slightly wider bottom than top margin (these default to 0.05”), no left margin, text aligned top, and word wrapping turned off. In addition, the auto-size behavior is set to adjust the width and height of the shape to fit its text. Note that vertical alignment is set on the text frame. Horizontal alignment is set on each paragraph:
```
from pptx.util import Inches
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE

text_frame = shape.text_frame
text_frame.text = 'Spam, eggs, and spam'
text_frame.margin_bottom = Inches(0.08)
text_frame.margin_left = 0
text_frame.vertical_anchor = MSO_ANCHOR.TOP
text_frame.word_wrap = False
text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
```

The possible values for `TextFrame.auto_size` and `TextFrame.vertical_anchor` are specified by the enumeration MSO\_AUTO\_SIZE and MSO\_VERTICAL\_ANCHOR respectively.

Applying paragraph formatting
-----

The following produces a shape containing three left-aligned paragraphs, the second and third indented (like sub-bullets) under the first:
```
from pptx.enum.text import PP_ALIGN

paragraph_strs = [
    'Egg, bacon, sausage and spam.',
    'Spam, bacon, sausage and spam.',
    'Spam, egg, spam, spam, bacon and spam.'
]

text_frame = shape.text_frame
text_frame.clear()

p = text_frame.paragraphs[0]
p.text = paragraph_strs[0]
p.alignment = PP_ALIGN.LEFT

for para_str in paragraph_strs[1:]:
    p = text_frame.add_paragraph()
    p.text = para_str
    p.alignment = PP_ALIGN.LEFT
    p.level = 1
```

Applying character formatting
-----

Character level formatting is applied at the run level, using the `.font` attribute. The following formats a sentence in 18pt Calibri Bold and applies the theme color Accent 1.
```
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Pt

text_frame = shape.text_frame
text_frame.clear()  # not necessary for newly-created shape

p = text_frame.paragraphs[0]
run = p.add_run()
run.text = 'Spam, eggs, and spam'

font = run.font
font.name = 'Calibri'
font.size = Pt(18)
font.bold = True
font.italic = None  # cause value to be inherited from theme
font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
```

If you prefer, you can set the font color to an absolute RGB value. Note that this will not change color when the theme is changed:
```
font.color.rgb = RGBColor(0xFF, 0x7F, 0x50)
```

A run can also be made into a hyperlink by providing a target URL:
```
run.hyperlink.address = 'https://github.com/scanny/python-pptx'
```

# Working with charts — python-pptx 1.0.0 documentation
python-pptx supports adding charts and modifying existing ones. Most chart types other than 3D types are supported.

Adding a chart
-----

The following code adds a single-series column chart in a new presentation:
```
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

# create presentation with 1 slide ------
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[5])

-----
chart_data = CategoryChartData()
chart_data.categories = ['East', 'West', 'Midwest']
chart_data.add_series('Series 1', (19.2, 21.4, 16.7))

-----
x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
)

prs.save('chart-01.pptx')
```

!../_images/chart-01.png

Customizing things a bit
-----

The remaining code will leave out code we’ve already seen and only show imports, for example, when they’re used for the first time, just to keep the focus on the new bits. Let’s create a multi-series chart to use for these examples:
```
chart_data = ChartData()
chart_data.categories = ['East', 'West', 'Midwest']
chart_data.add_series('Q1 Sales', (19.2, 21.4, 16.7))
chart_data.add_series('Q2 Sales', (22.3, 28.6, 15.2))
chart_data.add_series('Q3 Sales', (20.4, 26.3, 14.2))

graphic_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
)

chart = graphic_frame.chart
```

!../_images/chart-02.png

Notice that we captured the shape reference returned by the `add_chart()` call as `graphic_frame` and then extracted the chart object from the graphic frame using its `chart` property. We’ll need the chart reference to get to the properties we’ll need in the next steps. The `add_chart()` method doesn’t directly return the chart object. That’s because a chart is not itself a shape. Rather it’s a graphical (DrawingML) object _contained_ in the graphic frame shape. Tables work this way too, also being contained in a graphic frame shape.

XY and Bubble charts
------

The charts so far use a _discrete_ set of values for the independent variable (the X axis, roughly speaking). These are perfect when your values fall into a well-defined set of categories. However, there are many cases, particularly in science and engineering, where the independent variable is a continuous value, such as temperature or frequency. These are supported in PowerPoint by XY (aka. scatter) charts. A bubble chart is essentially an XY chart where the marker size is used to reflect an additional value, effectively adding a third dimension to the chart.

Because the independent variable is continuous, in general, the series do not all share the same X values. This requires a somewhat different data structure and that is provided for by distinct `XyChartData` and `BubbleChartData` objects used to specify the data behind charts of these types:
```
chart_data = XyChartData()

series_1 = chart_data.add_series('Model 1')
series_1.add_data_point(0.7, 2.7)
series_1.add_data_point(1.8, 3.2)
series_1.add_data_point(2.6, 0.8)

series_2 = chart_data.add_series('Model 2')
series_2.add_data_point(1.3, 3.7)
series_2.add_data_point(2.7, 2.3)
series_2.add_data_point(1.6, 1.8)

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.XY_SCATTER, x, y, cx, cy, chart_data
).chart
```

!../_images/chart-08.png

Creation of a bubble chart is very similar, having an additional value for each data point that specifies the bubble size:
```
chart_data = BubbleChartData()

series_1 = chart_data.add_series('Series 1')
series_1.add_data_point(0.7, 2.7, 10)
series_1.add_data_point(1.8, 3.2, 4)
series_1.add_data_point(2.6, 0.8, 8)

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.BUBBLE, x, y, cx, cy, chart_data
).chart
```

!../_images/chart-09.png

Axes
-----

Let’s change up the category and value axes a bit:
```
from pptx.enum.chart import XL_TICK_MARK
from pptx.util import Pt

category_axis = chart.category_axis
category_axis.has_major_gridlines = True
category_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
category_axis.tick_labels.font.italic = True
category_axis.tick_labels.font.size = Pt(24)

value_axis = chart.value_axis
value_axis.maximum_scale = 50.0
value_axis.minor_tick_mark = XL_TICK_MARK.OUTSIDE
value_axis.has_minor_gridlines = True

tick_labels = value_axis.tick_labels
tick_labels.number_format = '0"%"'
tick_labels.font.bold = True
tick_labels.font.size = Pt(14)
```

!../_images/chart-03.png

Okay, that was probably going a bit too far. But it gives us an idea of the kinds of things we can do with the value and category axes. Let’s undo this part and go back to the version we had before.

Data Labels
-----

Let’s add some data labels so we can see exactly what the value for each bar is:
```
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_LABEL_POSITION

plot = chart.plots[0]
plot.has_data_labels = True
data_labels = plot.data_labels

data_labels.font.size = Pt(13)
data_labels.font.color.rgb = RGBColor(0x0A, 0x42, 0x80)
data_labels.position = XL_LABEL_POSITION.INSIDE_END
```

!../_images/chart-04.png

Here we needed to access a Plot object to gain access to the data labels. A plot is like a sub-chart, containing one or more series and drawn as a particular chart type, like column or line. This distinction is needed for charts that combine more than one type, like a line chart appearing on top of a column chart. A chart like this would have two plot objects, one for the series appearing as columns and the other for the lines. Most charts only have a single plot and python-pptx doesn’t yet support creating multi-plot charts, but you can access multiple plots on a chart that already has them.

In the Microsoft API, the name _ChartGroup_ is used for this object. I found that term confusing for a long time while I was learning about MS Office charts so I chose the name Plot for that object in python-pptx.

Legend
-----

A legend is often useful to have on a chart, to give a name to each series and help a reader tell which one is which:
```
from pptx.enum.chart import XL_LEGEND_POSITION

chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.RIGHT
chart.legend.include_in_layout = False
```

!../_images/chart-05.png

Nice! Okay, let’s try some other chart types.

Line Chart
-----

A line chart is added pretty much the same way as a bar or column chart, the main difference being the chart type provided in the `add_chart()` call:
```
chart_data = ChartData()
chart_data.categories = ['Q1 Sales', 'Q2 Sales', 'Q3 Sales']
chart_data.add_series('West',    (32.2, 28.4, 34.7))
chart_data.add_series('East',    (24.3, 30.6, 20.2))
chart_data.add_series('Midwest', (20.4, 18.3, 26.2))

x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4.5)
chart = slide.shapes.add_chart(
    XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data
).chart

chart.has_legend = True
chart.legend.include_in_layout = False
chart.series[0].smooth = True
```

!../_images/chart-06.png

I switched the categories and series data here to better suit a line chart. You can see the line for the “West” region is _smoothed_ into a curve while the other two have their points connected with straight line segments.

Pie Chart
-----

A pie chart is a little special in that it only ever has a single series and doesn’t have any axes:
```
chart_data = ChartData()
chart_data.categories = ['West', 'East', 'North', 'South', 'Other']
chart_data.add_series('Series 1', (0.135, 0.324, 0.180, 0.235, 0.126))

chart = slide.shapes.add_chart(
    XL_CHART_TYPE.PIE, x, y, cx, cy, chart_data
).chart

chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM
chart.legend.include_in_layout = False

chart.plots[0].has_data_labels = True
data_labels = chart.plots[0].data_labels
data_labels.number_format = '0%'
data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
```

!../_images/chart-07.png

Odds & Ends
-----

This should be enough to get you started with adding charts to your presentation with python-pptx. There are more details in the API documentation for charts here: Charts

### About colors

By default, the colors assigned to each series in a chart are the theme colors Accent 1 through Accent 6, in that order. If you have more than six series, darker and lighter versions of those same colors are used. While it’s possible to assign specific colors to data points (bar, line, pie segment, etc.) for at least some chart types, the best strategy to start with is changing the theme colors in your starting “template” presentation.

# Working with tables — python-pptx 1.0.0 documentation
PowerPoint allows text and numbers to be presented in tabular form (aligned rows and columns) in a reasonably flexible way. A PowerPoint table is not nearly as functional as an Excel spreadsheet, and is definitely less powerful than a table in Microsoft Word, but it serves well for most presentation purposes.

Concepts
-----

There are a few terms worth reviewing as a basis for understanding PowerPoint tables:

table

A table is a matrix of cells arranged in aligned rows and columns. This orderly arrangement allows a reader to more easily make sense of relatively large number of individual items. It is commonly used for displaying numbers, but can also be used for blocks of text.

![../_images/table-01.png](../_images/table-01.png)

cell

An individual content “container” within a table. A cell has a text-frame in which it holds that content. A PowerPoint table cell can only contain text. I cannot hold images, other shapes, or other tables.

A cell has a background fill, borders, margins, and several other formatting settings that can be customized on a cell-by-cell basis.

row

A side-by-side sequence of cells running across the table, all sharing the same top and bottom boundary.

column

A vertical sequence of cells spanning the height of the table, all sharing the same left and right boundary.

table grid, also cell grid

The underlying cells in a PowerPoint table are strictly regular. In a three-by-three table there are nine grid cells, three in each row and three in each column. The presence of merged cells can obscure portions of the cell grid, but not change the number of cells in the grid. Access to a table cell in python-pptx is always via that cell’s coordinates in the cell grid, which may not conform to its visual location (or lack thereof) in the table.

merged cell

A cell can be “merged” with adjacent cells, horizontally, vertically, or both, causing the resulting cell to look and behave like a single cell that spans the area formerly occupied by those individual cells.

![../_images/table-02.png](../_images/table-02.png)

merge-origin cell

The top-left grid-cell in a merged cell has certain special behaviors. The content of that cell is what appears on the slide; content of any “spanned” cells is hidden. In python-pptx a merge-origin cell can be identified with the `_Cell.is_merge_origin` property. Such a cell can report the size of the merged cell with its `span_height` and `span_width` properties, and can be “unmerged” back to its underlying grid cells using its `split()` method.

spanned-cell

A grid-cell other than the merge-origin cell that is “occupied” by a merged cell is called a _spanned cell_. Intuitively, the merge-origin cell “spans” the other grid cells within its area. A spanned cell can be identified with its `_Cell.is_spanned` property. A merge-origin cell is not itself a spanned cell.

Adding a table
-----

The following code adds a 3-by-3 table in a new presentation:
```
>>> from pptx import Presentation
>>> from pptx.util import Inches

>>> # ---create presentation with 1 slide---
>>> prs = Presentation()
>>> slide = prs.slides.add_slide(prs.slide_layouts[5])

>>> # ---add table to slide---
>>> x, y, cx, cy = Inches(2), Inches(2), Inches(4), Inches(1.5)
>>> shape = slide.shapes.add_table(3, 3, x, y, cx, cy)

>>> shape
<pptx.shapes.graphfrm.GraphicFrame object at 0x1022816d0>
>>> shape.has_table
True
>>> table = shape.table
>>> table
<pptx.table.Table object at 0x1096f8d90>
```

![../_images/table-03.png](../_images/table-03.png)

A couple things to note:

*   `SlideShapes.add_table()` returns a shape that contains the table, not the table itself. In PowerPoint, a table is contained in a graphic-frame shape, as is a chart or SmartArt. You can determine whether a shape contains a table using its `has_table` property and you access the table object using the shape’s `table` property.

Inserting a table into a table placeholder
-----

A placeholder allows you to specify the position and size of a shape as part of the presentation “template”, and to place a shape of your choosing into that placeholder when authoring a presentation based on that template. This can lead to a better looking presentation, with objects appearing in a consistent location from slide-to-slide.

Placeholders come in different types, one of which is a _table placeholder_. A table placeholder behaves like other placeholders except it can only accept insertion of a table. Other placeholder types accept text bullets or charts.

There is a subtle distinction between a _layout placeholder_ and a _slide placeholder_. A layout placeholder appears in a slide layout, and defines the position and size of the placeholder “cloned” from it onto each slide created with that layout. As long as you don’t adjust the position or size of the slide placeholder, it will inherit it’s position and size from the layout placeholder it derives from.

To insert a table into a table placeholder, you need a slide layout that includes a table placeholder, and you need to create a slide using that layout. These examples assume that the third slide layout in template.pptx includes a table placeholder:
```
>>> prs = Presentation('template.pptx')
>>> slide = prs.slides.add_slide(prs.slide_layouts[2])
```

_Accessing the table placeholder._ Generally, the easiest way to access a placeholder shape is to know its position in the slide.shapes collection. If you always use the same template, it will always show up in the same position:
```
>>> table_placeholder = slide.shapes[1]
```

_Inserting a table._ A table is inserted into the placeholder by calling its `insert_table()` method and providing the desired number of rows and columns:
```
>>> shape = table_placeholder.insert_table(rows=3, cols=4)
```

The return value is a `GraphicFrame` shape containing the new table, not the table object itself. Use the `table` property of that shape to access the table object:

The containing shape controls the position and size. Everything else, like accessing cells and their contents, is done from the table object.

Accessing a cell
-----

All content in a table is in a cell, so getting a reference to one of those is a good place to start:
```
>>> cell = table.cell(0, 0)
>>> cell.text
''
>>> cell.text = 'Unladen Swallow'
```

![../_images/table-04.png](../_images/table-04.png)

The cell is specified by its row, column coordinates as zero-based offsets. The top-left cell is at row, column (0, 0).

Like an auto-shape, a cell has a text-frame and can contain arbitrary text divided into paragraphs and runs. Any desired character formatting can be applied individually to each run.

Often however, cell text is just a simple string. For these cases the read/write `_Cell.text` property can be the quickest way to set cell contents.

Merging cells
-----

A merged cell is produced by specifying two diagonal cells. The merged cell will occupy all the grid cells in the rectangular region specified by that diagonal:

![../_images/table-05.png](../_images/table-05.png)
```
>>> cell = table.cell(0, 0)
>>> other_cell = table.cell(1, 1)
>>> cell.is_merge_origin
False
>>> cell.merge(other_cell)
>>> cell.is_merge_origin
True
>>> cell.is_spanned
False
>>> other_cell.is_spanned
True
>>> table.cell(0, 1).is_spanned
True
```

![../_images/table-06.png](../_images/table-06.png)

A few things to observe:

*   The merged cell appears as a single cell occupying the space formerly occupied by the other grid cells in the specified rectangular region.
*   The formatting of the merged cell (background color, font etc.) is taken from the merge origin cell, the top-left cell of the table in this case.
*   Content from the merged cells was migrated to the merge-origin cell. That content is no longer present in the spanned grid cells (although you can’t see those at the moment). The content of each cell appears as a separate paragraph in the merged cell; it isn’t concatenated into a single paragraph. Content is migrated in left-to-right, top-to-bottom order of the original cells.
*   Calling `other_cell.merge(cell)` would have the exact same effect. The merge origin is always the top-left cell in the specified rectangular region. There are four distinct ways to specify a given rectangular region (two diagonals, each having two orderings).

Un-merging a cell
-----

A merged cell can be restored to its underlying grid cells by calling the `split()` method on its merge-origin cell. Calling `split()` on a cell that is not a merge-origin raises `ValueError`:
```
>>> cell = table.cell(0, 0)
>>> cell.is_merge_origin
True
>>> cell.split()
>>> cell.is_merge_origin
False
>>> table.cell(0, 1).is_spanned
False
```

![../_images/table-07.png](../_images/table-07.png)

Note that the content migration performed as part of the .merge() operation was not reversed.

A few snippets that might be handy
-----

Use Case: Interrogate table for merged cells:
```
def iter_merge_origins(table):
    """Generate each merge-origin cell in *table*.

    Cell objects are ordered by their position in the table,
    left-to-right, top-to-bottom.
    """
    return (cell for cell in table.iter_cells() if cell.is_merge_origin)

def merged_cell_report(cell):
    """Return str summarizing position and size of merged *cell*."""
    return (
        'merged cell at row %d, col %d, %d cells high and %d cells wide'
        % (cell.row_idx, cell.col_idx, cell.span_height, cell.span_width)
    )

# ---Print a summary line for each merged cell in *table*.---
for merge_origin_cell in iter_merge_origins(table):
    print(merged_cell_report(merge_origin_cell))
```

prints a report like:
```
merged cell at row 0, col 0, 2 cells high and 2 cells wide
merged cell at row 3, col 2, 1 cells high and 2 cells wide
merged cell at row 4, col 0, 2 cells high and 1 cells wide
```

Use Case: Access only cells that display text (are not spanned):
```
def iter_visible_cells(table):
    return (cell for cell in table.iter_cells() if not cell.is_spanned)
```

Use Case: Determine whether table contains merged cells:
```
def has_merged_cells(table):
    for cell in table.iter_cells():
        if cell.is_merge_origin:
            return True
    return False
```
