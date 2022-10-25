import os
from pptx import Presentation
from pptx.util import Inches, Mm, Cm
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT as PP_ALIGN
from pptx.util import Pt
from pptx.dml.color import RGBColor
import xml.etree.ElementTree as et


# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def add_verse_slide(pres, verse):
    # Title and content layout
    layout = pres.slide_layouts[5]
    slide = pres.slides.add_slide(layout)

    # Add verse text

    # Split string on newlines and strip leading (and trailing) whitespace
    verse_paragraphs = list(map(lambda v: v.strip(), verse.split('\n')))

    verse_text = ''

    # Re-create the verse text for the slide paragraph
    for i, vp in enumerate(verse_paragraphs):
        verse_text += vp

        if i != len(verse_paragraphs) - 1:
            verse_text += '\n'

    slide.shapes.add_textbox(Inches(0.125), Inches(3), Inches(13.125), Inches(4))

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        text_frame = shape.text_frame

        # Clear out the default empty paragraph
        text_frame.clear()
        text_frame.left = Inches(0)
        text_frame.width = Inches(13)
        p = text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.level = 0
        # p.text = verse_text

        run = p.add_run()
        run.text = verse_text
        font = run.font
        font.name = 'Lucida Sans'
        font.size = Pt(40)
        font.color.rgb = RGBColor(255, 255, 255)

    # Add song title as slide title
    slide.shapes.title.text = f'{title} #{song_number}'
    tf = slide.shapes.title.text_frame
    tfp = tf.paragraphs[0]
    tfp.alignment = PP_ALIGN.CENTER
    tfpr = tfp.runs[0]
    tfprf = tfpr.font
    tfprf.name = 'Lucida Sans'
    tfprf.size = Pt(48)
    tfprf.color.rgb = RGBColor(255, 255, 255)
    tfprf.bold = True
    tfprf.italic = True

    # Add Anthem logo to right of slide
    slide.shapes.add_picture('anthem-logo.png', Inches(11.75), Inches(0))


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print('Generating powerpoint presentations...')

    src_dir = 'examples'
    hymn_xml_files = os.listdir(src_dir)

    for f in hymn_xml_files:
        if not f.endswith('.xml'):
            continue

        print(f'\tCreating pptx for file: {os.path.join(src_dir, f)}')

        tree = et.parse(f'{os.path.join(src_dir, f)}')
        root = tree.getroot()

        title = None
        song_number = None
        chorus = None
        verses = []
        author = None

        for child in root:
            if child.tag == 'Title':
                title = child.text
            elif child.tag == 'SongNumber':
                song_number = child.text
            elif child.tag == 'Text' and child.attrib['section'] == 'Chorus':
                chorus = child.text.strip()
            elif child.tag == 'Text' and child.attrib['section'] != 'Chorus':
                verses.append(child.text.strip())
            elif child.tag == 'Author':
                author = child.text.strip()

        # Create a pptx file using the base pptx so we can use its theme
        pres = Presentation('base.pptx')

        # Delete the first slide (yuck). See https://stackoverflow.com/a/70969310
        for i in range(len(pres.slides) - 1, -1, -1):
            rId = pres.slides._sldIdLst[i].rId
            pres.part.drop_rel(rId)
            del pres.slides._sldIdLst[i]

        for verse in verses:
            add_verse_slide(pres, verse)
            add_verse_slide(pres, chorus)

        # Add the last slide

        # Blank layout
        layout = pres.slide_layouts[6]

        end_slide = pres.slides.add_slide(layout)
        end_slide.shapes.add_textbox(Inches(0.125), Inches(7), Inches(13.125), Inches(0.5))

        for shape in end_slide.shapes:
            if not shape.has_text_frame:
                continue

            text_frame = shape.text_frame

            # Clear out the default empty paragraph
            text_frame.clear()
            text_frame.left = Inches(0)
            text_frame.width = Inches(13)
            p = text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            p.level = 0
            # p.text = verse_text

            run = p.add_run()
            run.text = author
            font = run.font
            font.name = 'Calibri'
            font.size = Pt(14)
            font.color.rgb = RGBColor(255, 255, 255)

        filename = f'{song_number.zfill(3)} {title}.pptx'
        filepath = os.path.join('powerpoints', filename)
        pres.save(filepath)

    print('Done!')
