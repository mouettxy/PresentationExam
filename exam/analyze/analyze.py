import csv
from pathlib import Path

from win32com.client import Dispatch

from .images import Images
from ..config import get_analyze
from ..constants import msoOrientationHorizontal
from ..utils import (layouts, layout_to_dict, pt_to_px, is_text, is_image, is_title, check_collision_between_shapes,
                     get_shape_dimensions, get_shape_crop_values, get_download_path, dict_to_string)

Application = Dispatch("PowerPoint.Application")
config = get_analyze()


class Analyze:
    def __init__(self, presentation_path):
        super().__init__()
        self._Images = Images(presentation_path)
        self._Presentation = Application.Presentations.Open(presentation_path, WithWindow=False)

    def which_layout(self):
        for layout in layouts:
            layout_positions = layout_to_dict(
                pt_to_px(self._Presentation.PageSetup.SlideWidth),
                pt_to_px(self._Presentation.PageSetup.SlideHeight),
                layout
            )
            elements, collision = set(), set()
            for Slide in self._Presentation.Slides:
                if Slide.SlideIndex == 2 or Slide.SlideIndex == 3:
                    for Shape in Slide.Shapes:
                        elements.add(Shape.Name)
                        layout_position = layout_positions[Slide.SlideIndex]
                        shape_dims = get_shape_dimensions(Shape)
                        cur_layout_position = []
                        if is_title(Shape) or is_text(Shape):
                            cur_layout_position = layout_position["title"] + layout_position["text"]
                        elif is_image(Shape):
                            cur_layout_position = layout_position["images"]
                        for pos in cur_layout_position:
                            if check_collision_between_shapes(shape_dims, pos):
                                collision.add(Shape.Name)
            if len(elements) == len(collision):
                return layout
        return False

    def __analyze_count_of_slides(self):
        if self._Presentation.Slides.Count == int(config['slides']):
            return True
        return False

    def __analyze_slides_aspect_ratio(self):
        PageSetup = self._Presentation.PageSetup
        aspect_ratio = config['aspect_ratio'].split('/')
        if (PageSetup.SlideWidth / PageSetup.SlideHeight) / (int(aspect_ratio[0]) / int(aspect_ratio[1])):
            return True
        return False

    def __analyze_typefaces(self):
        typefaces = set()
        for Slide in self._Presentation.Slides:
            for Shape in Slide.Shapes:
                if is_text(Shape):
                    typefaces.add(Shape.TextFrame.TextRange.Font.Name)

        if len(typefaces) == 1:
            return True
        else:
            tmp_typefaces = set()
            for typeface in list(typefaces):
                tmp_typefaces.add(typeface.split()[0])
            if len(tmp_typefaces) == 1:
                return True
        return False

    def __collisions_between_slide_elements(self, slide):
        overlaps = set()
        shapes_1, shapes_2 = ([Shape for Shape in self._Presentation.Slides(slide).Shapes],
                              [Shape for Shape in self._Presentation.Slides(slide).Shapes])
        for k in range(len(shapes_1) - 1):
            for j in range(1, len(shapes_2)):
                dims_1, dims_2 = get_shape_dimensions(shapes_1[k]), get_shape_dimensions(shapes_2[j])
                if shapes_1[k].Name != shapes_2[j].Name:
                    collision = check_collision_between_shapes(dims_1, dims_2)
                    overlaps.add(collision)
        if len(overlaps) == 1:
            return True
        return False

    def __analyze_slide_text_image_blocks(self, slide):
        text, images, title, subtitle = 0, 0, False, False
        for Shape in self._Presentation.Slides(slide).Shapes:
            if is_title(Shape):
                if slide == 1:
                    if not title and not subtitle:
                        title = True
                    elif title and not subtitle:
                        subtitle = True
                else:
                    if not title:
                        title = True
            elif is_text(Shape):
                text += 1
            elif is_image(Shape):
                images += 1
        a_text, a_images, a_title, a_subtitle = False, False, False, False
        if slide == 1:
            if text == int(config[f'text_blocks_{slide}']) and not title and not subtitle:
                a_text, a_title, a_subtitle = True, True, True
            elif text == int(config[f'text_blocks_{slide}']) - 1 and title and not subtitle:
                a_text, a_title, a_subtitle = True, True, False
            elif text == int(config[f'text_blocks_{slide}']) - 2 and title and subtitle:
                a_text, a_title, a_subtitle = True, True, True
            if images == int(config[f'images_{slide}']):
                a_images = True
            return a_text, a_images, a_title, a_subtitle
        else:
            if text == int(config[f'text_blocks_{slide}']) and not title:
                a_text, a_title = True, True
            elif text == int(config[f'text_blocks_{slide}']) - 1 and title:
                a_text = True
            elif text == int(config[f'text_blocks_{slide}']) - 1 and not title:
                a_text = True
            if images == int(config[f'images_{slide}']):
                a_images = True
            return a_text, a_images, a_title, a_subtitle

    def __analyze_slide_font_sizes(self, slide):
        font_sizes, correct_counter = [], 0
        for Shape in self._Presentation.Slides(slide).Shapes:
            if is_text(Shape):
                font_sizes.append(Shape.TextFrame.TextRange.Font.Size)
        required_font_sizes = config[f'font_sizes_{slide}'].split(",")
        for f in range(len(required_font_sizes)):
            required_font_sizes[f] = float(required_font_sizes[f])
        if len(required_font_sizes) == len(font_sizes) == int(config[f'text_blocks_{slide}']):
            if required_font_sizes == font_sizes:
                return True
        elif len(required_font_sizes) - 1 == len(font_sizes) == int(config[f'text_blocks_{slide}']) - 1:
            required_font_sizes.pop(0)
            if required_font_sizes == font_sizes:
                return True
        return False

    def presentation(self):
        analyze = {}
        layout = self.which_layout()
        analyze[0] = self.__analyze_count_of_slides()
        analyze[1] = self.__analyze_slides_aspect_ratio()
        analyze[2] = self._Presentation.PageSetup.SlideOrientation == msoOrientationHorizontal
        analyze[3] = self.__analyze_typefaces()
        analyze[4] = self._Images.compare()
        analyze[5], analyze[6] = True if layout else False, layout if layout else None
        analyze[13] = self._Images.distorted_images()
        return analyze

    def slide_1(self):
        analyze = {}
        text, images, title, subtitle = self.__analyze_slide_text_image_blocks(1)
        analyze[7], analyze[8] = title, subtitle
        analyze[9] = self.__collisions_between_slide_elements(1)
        analyze[10] = self.__analyze_slide_font_sizes(1)
        return analyze

    def slide_2(self):
        analyze = {}
        text, images, title, subtitle = self.__analyze_slide_text_image_blocks(2)
        analyze[7] = title
        analyze[9] = self.__collisions_between_slide_elements(2)
        analyze[10] = self.__analyze_slide_font_sizes(2)
        analyze[11], analyze[12] = text, images
        return analyze

    def slide_3(self):
        analyze = {}
        text, images, title, subtitle = self.__analyze_slide_text_image_blocks(3)
        analyze[7] = title
        analyze[9] = self.__collisions_between_slide_elements(3)
        analyze[10] = self.__analyze_slide_font_sizes(3)
        analyze[11], analyze[12] = text, images
        return analyze

    @staticmethod
    def __translate(analyze, grade):
        presentation_analyze = {
            "Соотношение сторон 16:9": "Выполнено" if analyze[1] else "Не выполнено",
            "Горизонтальная ориентация": "Выполнено" if analyze[2] else "Не выполнено"
        }

        structure_analyze = {
            "Три слайда в презентации": "Выполнено" if analyze[0] else "Не выполнено",
            "Соответствует макету": "Выполнено" if analyze[5] else "Не выполнено",
            "Заголовки на слайдах": "Выполнено" if analyze[7] else "Не выполнено",
            "Подзаголовок на первом слайде": "Выполнено" if analyze[8] else "Не выполнено",
            "Элементы не перекрывают друг друга": "Выполнено" if analyze[9] else "Не выполнено",
            "Текстовые блоки на 2, 3 слайде": "Выполнено" if analyze[11] else "Не выполнено",
            "Картинки на 2, 3 слайде": "Выполнено" if analyze[12] else "Не выполнено"
        }

        fonts_analyze = {
            "Единый тип шрифта": "Выполнено" if analyze[3] else "Не выполнено",
            "Размер шрифта": "Выполнено" if analyze[10] else "Не выполнено"
        }

        images_analyze = {
            "Оригинальные картинки": "Выполнено" if analyze[4] else "Не выполнено",
            "Картинки не искажены": "Выполнено" if not analyze[13] else "Не выполнено"
        }
        return presentation_analyze, structure_analyze, fonts_analyze, images_analyze, analyze[6], grade

    def __summary(self):
        """
        0 - does prs have 3 slides
        1 - right aspect ratio of presentation
        2 - horizontal orientation
        3 - right typefaces
        4 - original photos
        5 - contatins layout
        6 - which layout
        7 - title
        8 - subtitle
        9 - overlaps
        10 - font sizes
        11 - text blocks
        12 - image blocks
        13 - distorted images

        Presentation:
            1, 2
        Structure:
            0, 5, 7, 8, 9, 11, 12
        Fonts:
            3, 10
        Images:
            4 13
        """
        # 16 HOURS HERE
        presentation_info, first_slide, second_slide = (self.presentation(),
                                                        self.slide_1(),
                                                        self.slide_2())
        if self._Presentation.Slides.Count >= 3:
            third_slide = self.slide_3()
            data = {**presentation_info, **first_slide, **second_slide, **third_slide}
            err_structure = [data[k] for k in data if k in [0, 5, 7, 8, 9, 11, 12]].count(False)
            err_fonts = [data[k] for k in data if k in [3, 10]].count(False)
            err_images = [data[k] for k in data if k in [4, 13]].count(False)
            # compute grade
            r_grade = 0
            # check if we can give grade 2 (max)
            if all(value for value in data.values()):
                r_grade = 2
                return self.__translate(data, r_grade)
            # or if we can give 1
            if err_structure == 1 and not err_fonts and not err_images:
                r_grade = 1
            elif not err_structure and err_fonts == 1 and not err_images:
                r_grade = 1
            elif not err_structure and not err_fonts and err_images == 1:
                r_grade = 1
            return self.__translate(data, r_grade)
        elif self._Presentation.Slides.Count == 2:
            data = {**presentation_info, **first_slide, **second_slide}
            err_structure = [data[k] for k in data if k in [0, 5, 7, 8, 9, 11, 12]].count(False)
            err_fonts = [data[k] for k in data if k in [3, 10]].count(False)
            err_images = [data[k] for k in data if k in [4, 13]].count(False)
            # compute grade
            r_grade = 0
            if err_structure == 1 and not err_fonts and not err_images:
                r_grade = 1
            return self.__translate(data, r_grade)

    def get(self, typeof="analyze"):
        if typeof == "analyze":
            return self.__summary()
        elif typeof == "thumb":
            return self._Images.get("thumb")
        elif typeof == "slides":
            return self._Presentation.Slides.Count

    def __del__(self):
        Application.Quit()

    def __exit__(self):
        Application.Quit()

    @property
    def warnings(self):
        warnings = {0: [], 1: [], 2: [], 3: []}
        shape_animations, slide_1_text_blocks = 0, 0
        for Slide in self._Presentation.Slides:
            # count slide animations or entry effects
            if Slide.TimeLine.MainSequence.Count >= 1:
                shape_animations += Slide.TimeLine.MainSequence.Count
            if Slide.SlideShowTransition.EntryEffect:
                warnings[0].append(f"Анимация перехода на слайде {Slide.SlideIndex}.")

            for Shape in Slide.Shapes:
                crop = get_shape_crop_values(Shape)
                if crop:
                    crop_warning = (f'Объект {Shape.Name}, {Shape.Id} обрезан {crop["left"]}:{crop["right"]}:'
                                    f':{crop["top"]}:{crop["bottom"]}')
                if Slide.SlideIndex == 1:
                    if is_image(Shape):
                        warnings[1].append(f"Изображение {Shape.Name} с ID {Shape.Id}")
                    elif is_text(Shape) is True:
                        slide_1_text_blocks += 1
                    elif is_text(Shape) is None:
                        warnings[1].append(f"Пустой текстовый блок {Shape.Name}, {Shape.Id}")
                    elif not is_text(Shape):
                        warnings[1].append(f"Неизвестный объект {Shape.Name}, {Shape.Id}")
                    if slide_1_text_blocks > 2:
                        warnings[1].append(f"Больше двух текстовых элементов на слайде.")
                elif Slide.SlideIndex == 2:
                    if is_text(Shape) is None:
                        warnings[2].append(f"Пустой текстовый блок {Shape.Name}, {Shape.Id}")
                    elif not is_text(Shape) and not is_image(Shape):
                        warnings[2].append(f"Неизвестный объект {Shape.Name}, {Shape.Id}")
                    if crop:
                        warnings[2].append(crop_warning)
                elif Slide.SlideIndex == 3:
                    if is_text(Shape) is None:
                        warnings[3].append(f"Пустой текстовый блок {Shape.Name}, {Shape.Id}")
                    elif not is_text(Shape) and not is_image(Shape):
                        warnings[3].append(f"Неизвестный объект {Shape.Name} с ID {Shape.Id}")
                    if crop:
                        warnings[3].append(crop_warning)
        if shape_animations:
            warnings[0].append(f"Анимации в объектах: {shape_animations}.")
        return warnings

    def export_csv(self):
        warnings = self.warnings
        presentation, structure, fonts, images, layout, grade = self.get()
        warn_0, warn_1, warn_2, warn_3 = warnings[0], warnings[1], warnings[2], warnings[3]
        path = Path.joinpath(Path(get_download_path()), self._Presentation.Name + ".csv")
        fieldnames = ['Презентация', 'Структура', 'Шрифты', 'Картинки', 'Предупреждения', 'Слайд 1', 'Слайд 2',
                      'Слайд 3']
        with open(path, "w", newline='', encoding="windows-1251") as fCsv:
            writer = csv.writer(fCsv, delimiter=',')
            writer.writerow(fieldnames)
            writer.writerow([
                dict_to_string(presentation),
                dict_to_string(structure),
                dict_to_string(fonts),
                dict_to_string(images),
                '\n'.join(warn_0),
                '\n'.join(warn_1),
                '\n'.join(warn_2),
                '\n'.join(warn_3),
            ])
        return path
