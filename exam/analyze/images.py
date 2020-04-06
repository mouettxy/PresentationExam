import shutil
import zipfile
from pathlib import Path

from PIL import Image, ImageDraw
from imagehash import average_hash
from win32com.client import Dispatch

from ..constants import ppShapeFormatJPG
from ..utils import is_text, is_image, pt_to_px, get_shape_dimensions, layout_to_dict, get_shape_percentage_width_height

Application = Dispatch("PowerPoint.Application")


class Images:
    def __init__(self, presentation_path):
        super().__init__()
        self._path = presentation_path
        self._Presentation = Application.Presentations.Open(presentation_path, WithWindow=False)
        Path("temp").mkdir(exist_ok=True, parents=True)
        self.destination = Path(f"temp/{self._Presentation.Name}").resolve()
        self.destination.mkdir(exist_ok=True, parents=True)

    @staticmethod
    def __draw_rectangle(draw, shape_dimensions, color, outline="red"):
        draw.rectangle(
            [shape_dimensions['left'],
             shape_dimensions['top'],
             shape_dimensions['width'] + shape_dimensions['left'],
             shape_dimensions['height'] + shape_dimensions['top']],
            fill=color,
            outline=outline
        )

    def skeleton(self):
        paths = []
        for Slide in self._Presentation.Slides:
            path = Path.joinpath(self.destination, f"skeleton_{Slide.SlideIndex}.jpg")
            image = Image.new(
                "RGB",
                color="white",
                size=(pt_to_px(self._Presentation.PageSetup.SlideWidth),
                      pt_to_px(self._Presentation.PageSetup.SlideHeight))
            )
            skeleton_draw = ImageDraw.Draw(image)
            for Shape in Slide.Shapes:
                shape_dimensions = get_shape_dimensions(Shape)
                if is_text(Shape) is True:
                    self.__draw_rectangle(skeleton_draw, shape_dimensions, "yellow", "red")
                elif is_text(Shape) is None:
                    self.__draw_rectangle(skeleton_draw, shape_dimensions, "orange", "yellow")
                elif is_image(Shape) is True:
                    self.__draw_rectangle(skeleton_draw, shape_dimensions, "blue", "yellow")
                else:
                    self.__draw_rectangle(skeleton_draw, shape_dimensions, "red", "yellow")
            image.save(path)
            paths.append(path)
        return paths

    def layout(self, lt="DEFAULT"):
        """
        Experimental
        """
        paths = []
        color = (250, 250, 250, 1)
        layout = layout_to_dict(pt_to_px(self._Presentation.PageSetup.SlideWidth),
                                pt_to_px(self._Presentation.PageSetup.SlideHeight),
                                lt)
        for slide in layout:
            path = Path.joinpath(self.destination, f"layout_{slide}.png")
            image = Image.new(
                "RGB",
                (pt_to_px(self._Presentation.PageSetup.SlideWidth), pt_to_px(self._Presentation.PageSetup.SlideHeight)),
                "white"
            )
            draw = ImageDraw.Draw(image, "RGBA")
            for block_type in layout[slide]:
                if block_type == "title":
                    color = (27, 94, 32, 100)
                elif block_type == "images":
                    color = (245, 127, 23, 120)
                elif block_type == "text":
                    color = (26, 35, 126, 175)
                for dims in layout[slide][block_type]:
                    self.__draw_rectangle(draw, dims, color)
            image.save(path)
            paths.append(path)
        return paths

    def get_shape_images(self):
        """
        Reserved for internal use
        """
        destination = Path.joinpath(self.destination, 'shapes')
        destination.mkdir(exist_ok=True, parents=True)
        paths = []
        for Slide in self._Presentation.Slides:
            for Shape in Slide.Shapes:
                if is_image(Shape):
                    path = Path.joinpath(destination, f"{Slide.SlideIndex}_{Shape.Id}.jpg")
                    Shape.Export(path, ppShapeFormatJPG)
                    paths.append(path)
        return paths

    def save_original_images(self):
        paths = []
        file = zipfile.ZipFile(self._path)
        destination = Path.joinpath(self.destination, "media")
        destination.mkdir(parents=True, exist_ok=True)
        for f in file.namelist():
            if f.startswith('ppt/media'):
                path = Path.joinpath(destination, Path(f).name)
                shutil.copyfileobj(file.open(f), open(path, "wb"))
                paths.append(path)
        return paths

    def compare(self, path='original_images'):
        if Path(path).exists():
            original_images, compare_counter, shape_images, images_counter = [], 0, self.save_original_images(), 0
            for extension in ['*.png', '*.jpg', '*.jpeg']:
                original_images.extend(Path(path).resolve().glob(extension))
            for o_path in original_images:
                for s_path in shape_images:
                    o_image = Image.open(o_path)
                    s_image = Image.open(s_path)
                    if average_hash(o_image) == average_hash(s_image):
                        compare_counter += 1
                        break
            for Slide in self._Presentation.Slides:
                for Shape in Slide.Shapes:
                    if is_image(Shape):
                        images_counter += 1
            if compare_counter == images_counter:
                return True
        return False

    def distorted_images(self):
        for Slide in self._Presentation.Slides:
            for Shape in Slide.Shapes:
                if is_image(Shape):
                    w, h = get_shape_percentage_width_height(Shape)
                    if abs(w - h) > 10:
                        return True
        return False

    def get(self, thumb=False):
        paths = []
        for Slide in self._Presentation.Slides:
            path = Path.joinpath(self.destination, f"screenshot_{Slide.SlideIndex}.jpg")
            Slide.Export(path, "JPG")
            paths.append(path)
            if thumb:
                image = Image.open(path)
                image.thumbnail((200, 200), Image.ANTIALIAS)
                image.save(Path(self.destination, "thumb.jpg"))
                return Path(self.destination, "thumb.jpg")
        return paths
