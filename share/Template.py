from enum import Enum

class Template:
    layout_index = -1
    def parse(self, item, lyric):
        pass
    def print(ppt, lyric):
        pass
    def convert_back(slide):
        pass


class TemplateType(Enum):
    Blank = 0
    Title = 1
    Title_With_Author = 2

    Single = 3
    ZH_EN = 4
    EN_ZH = 5

    Poem_Single = 6
    Poem_Translate = 7
    Poem_Translate2 = 8

    Undefined = -1

class Manager:
    

    class Blank(Template):
        layout_index = TemplateType.Blank.value
        def parse(item, lyric):
            pass
        def print(ppt, lyric):
            ppt.slides.add_slide(ppt.slide_layouts[TemplateType.Blank.value])
        def convert_back(slide):
            return {"template": "Blank", "items": [{}]}
    
    class Title(Template):
        layout_index = TemplateType.Title.value
        def parse(item, lyric):
            lyric.title = item['title']
        def print(ppt, lyric):
            slide = ppt.slides.add_slide(ppt.slide_layouts[TemplateType.Title.value])
            slide.placeholders[0].text = lyric.title
        def convert_back(slide):
            return {"template": "Title", "items": [{"title": slide.placeholders[0].text}]}

    class Title_With_Author(Template):
        layout_index = TemplateType.Title_With_Author.value
        def parse(item, lyric):
            lyric.title = item['title']
            lyric.author = item['author']
        def print(ppt, lyric):
            slide = ppt.slides.add_slide(ppt.slide_layouts[TemplateType.Title_With_Author.value])
            slide.placeholders[13].text = lyric.title
            slide.placeholders[14].text = lyric.author
        def convert_back(slide):
            return {"template": "Title_With_Author", "items": [{"title": slide.placeholders[13].text, "author": slide.placeholders[14].text}]}


    class Single(Template):
        layout_index = TemplateType.Single.value
        def parse(item, lyric):
            lyric.text = item['text']
        def print(ppt, lyric):
            slide = ppt.slides.add_slide(ppt.slide_layouts[TemplateType.Single.value])
            slide.placeholders[13].text = lyric.text
        def convert_back(slide):
            return {"template": "Single", "items": [{"text": slide.placeholders[13].text}]}

    class ZH_EN(Template):
        layout_index = TemplateType.ZH_EN.value
        def parse(item, lyric):
            lyric.zh = item['zh']
            lyric.en = item['en']
        def print(ppt, lyric):
            slide = ppt.slides.add_slide(ppt.slide_layouts[TemplateType.ZH_EN.value])
            slide.placeholders[13].text = lyric.zh
            slide.placeholders[14].text = lyric.en
        def convert_back(slide):
            return {"template": "ZH_EN", "items": [{"zh": slide.placeholders[13].text, "en": slide.placeholders[14].text}]}
    
    class EN_ZH(Template):
        layout_index = TemplateType.EN_ZH.value
        def parse(item, lyric):
            lyric.en = item['en']
            lyric.zh = item['zh']
        def print(ppt, lyric):
            slide = ppt.slides.add_slide(ppt.slide_layouts[TemplateType.EN_ZH.value])
            slide.placeholders[13].text = lyric.en
            slide.placeholders[14].text = lyric.zh
        def convert_back(slide):
            return {"template": "EN_ZH", "items": [{"en": slide.placeholders[13].text, "zh": slide.placeholders[14].text}]}

    class Poem_Single(Template):
        layout_index = TemplateType.Poem_Single.value
        def parse(item, lyric):
            lyric.text = item['text']
        def print(ppt, lyric):
            slide = ppt.slides.add_slide(ppt.slide_layouts[TemplateType.Poem_Single.value])
            slide.placeholders[13].text = lyric.text
        def convert_back(slide):
            return {"template": "Poem_Single", "items": [{"text": slide.placeholders[13].text}]}

    class Poem_Translate(Template):
        layout_index = TemplateType.Poem_Translate.value
        def parse(item, lyric):
            lyric.text = item['text']
            lyric.translate = item['translate']
        def print(ppt, lyric):
            slide = ppt.slides.add_slide(ppt.slide_layouts[TemplateType.Poem_Translate.value])
            slide.placeholders[13].text = lyric.text
            slide.placeholders[14].text = lyric.translate
        def convert_back(slide):
            return {"template": "Poem_Translate", "items": [{"text": slide.placeholders[13].text, "translate": slide.placeholders[14].text}]}

    class Poem_Translate2(Template):
        layout_index = TemplateType.Poem_Translate2.value
        def parse(item, lyric):
            lyric.text = item['text']
            lyric.translate = item['translate']
        def print(ppt, lyric):
            slide = ppt.slides.add_slide(ppt.slide_layouts[TemplateType.Poem_Translate2.value])
            slide.placeholders[13].text = lyric.text
            slide.placeholders[14].text = lyric.translate
        def convert_back(slide):
            return {"template": "Poem_Translate2", "items": [{"text": slide.placeholders[13].text, "translate": slide.placeholders[14].text}]}

    def __init__(self):
        pass

    def parse(self, slide_type, item, lyric):
        try:
            getattr(self, slide_type).parse(item, lyric)
        except AttributeError:
            lyric.template = TemplateType.Undefined
            print("Error: No parser for template type "+slide_type)
            print("item:", item)
            input()
            exit(0)
        except KeyError:
            lyric.template = TemplateType.Undefined
            print("Error: No key for template type "+self.type.name)
            print("item:", item)
            input()
            exit(0)
        
    def print(self, ppt, lyric):
        try:
            getattr(self, lyric.template).print(ppt, lyric)
        except AttributeError:
            print("Error: No printer for template type "+lyric.template)
            print("lyric:", lyric)
            input()
            exit(0)

    def convert_back(self, slide, template_index):
        try:
            templatename = TemplateType(template_index).name
            return getattr(self, templatename).convert_back(slide)
        except AttributeError:
            print("Error: No converter for template type "+templatename)
            print("slide:", slide)
            input()
            exit(0)