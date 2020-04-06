import os
import shutil
from pathlib import Path

from win32com.client import Dispatch

import exam.config as configuration
from .constants import (msoTrue, msoPicture, msoLinkedPicture, msoPlaceholder, ppPlaceholderCenterTitle,
                        ppPlaceholderTitle, ppPlaceholderSubtitle, ppPlaceholderPicture, msoScaleFromTopLeft)

Application = Dispatch("PowerPoint.Application")
config, layouts = configuration.get_constants(), configuration.get_layouts()
text_out_of_bounds = int(config['text out of bounds'])
text_dimensions_average = int(config['text dimensions average'])


def pt_to_px(value):
    return round(value / 72 * 96)


def is_text(Shape):
    if Shape.HasTextFrame and Shape.Visible == msoTrue:
        if Shape.TextFrame.HasText:
            dims = get_shape_dimensions(Shape)
            if dims["left"] < text_out_of_bounds or dims['top'] < text_out_of_bounds:
                return None
            return True
        else:
            return None
    return False


def is_image(Shape):
    if Shape.Type == msoPicture or Shape.Type == msoLinkedPicture:
        return True
    if Shape.Type == msoPlaceholder:
        if Shape.PlaceholderFormat.Type == ppPlaceholderPicture:
            return True
    return False


def is_title(Shape):
    if Shape.Type == msoPlaceholder:
        if (Shape.PlaceholderFormat.Type == ppPlaceholderCenterTitle or
                Shape.PlaceholderFormat.Type == ppPlaceholderSubtitle or
                Shape.PlaceholderFormat.Type == ppPlaceholderTitle):
            return True
    return False


def get_shape_dimensions(Shape):
    if Shape.HasTextFrame:
        if Shape.TextFrame.HasText:
            movement_counter = 0
            text_length = Shape.TextFrame.TextRange.Length
            # if shape has text, we delete all end characters like enter, vertical tab, spaces, etc
            while (Shape.TextFrame.TextRange.Characters(text_length, 1).Text == chr(13) or
                   Shape.TextFrame.TextRange.Characters(text_length, 1).Text == chr(11) or
                   Shape.TextFrame.TextRange.Characters(text_length, 1).Text == chr(10) or
                   Shape.TextFrame.TextRange.Characters(text_length, 1).Text == chr(32)):
                Shape.TextFrame.TextRange.Characters(text_length, 1).Delete()
                text_length -= 1
                movement_counter += 1
            else:
                Range = Shape.TextFrame.TextRange
                TextFrame = Shape.TextFrame
                # int 7 that is experimental average value
                shape_t = pt_to_px(Range.BoundTop) - pt_to_px(TextFrame.MarginTop) + text_dimensions_average
                shape_l = pt_to_px(Range.BoundLeft) - pt_to_px(TextFrame.MarginLeft) + text_dimensions_average
                shape_w = pt_to_px(Range.BoundWidth) - pt_to_px(TextFrame.MarginRight) - text_dimensions_average
                shape_h = pt_to_px(Range.BoundHeight) - pt_to_px(TextFrame.MarginBottom) - text_dimensions_average
                # undo all what we deleted, because for some reason, this changes saved in presentation
                for i in range(movement_counter):
                    Application.StartNewUndoEntry()
                return {
                    'left': shape_l,
                    'top': shape_t,
                    'width': shape_w,
                    'height': shape_h
                }
    return {
        'top': pt_to_px(Shape.Top),
        'left': pt_to_px(Shape.Left),
        'width': pt_to_px(Shape.Width),
        'height': pt_to_px(Shape.Height),
    }


def get_shape_crop_values(Shape):
    if is_image(Shape):
        result = {
            'left': pt_to_px(Shape.PictureFormat.CropLeft),
            'top': pt_to_px(Shape.PictureFormat.CropTop),
            'right': pt_to_px(Shape.PictureFormat.CropRight),
            'bottom': pt_to_px(Shape.PictureFormat.CropBottom),
        }
        if result["left"] or result["top"] or result["right"] or result["bottom"]:
            return result
    else:
        return None


def get_shape_percentage_width_height(Shape, original_w_h=False):
    shape_width, shape_height = Shape.Width, Shape.Height
    Shape.ScaleWidth(1, msoTrue, msoScaleFromTopLeft)
    Shape.ScaleHeight(1, msoTrue, msoScaleFromTopLeft)
    original_width, original_height = (Shape.Width,
                                       Shape.Height)
    percentage_width, percentage_height = (shape_width / original_width * 100,
                                           shape_height / original_height * 100)
    Shape.ScaleWidth(percentage_width / 100, msoTrue)
    Shape.ScaleHeight(percentage_height / 100, msoTrue)
    if original_w_h:
        return round(original_width), round(original_height)
    return round(percentage_width), round(percentage_height)


def dict_to_list(dictionary, key=None):
    if key is None:
        for d in dictionary:
            if type(dictionary[d]) == dict:
                for d2 in dictionary[d]:
                    yield dictionary[d][d2]
            else:
                yield dictionary[d]
    else:
        for d in dictionary[key]:
            yield dictionary[key][d]


def check_collision_between_shapes(first_shape, second_shape):
    if (first_shape['left'] + first_shape['width'] > second_shape['left'] and
            first_shape['left'] < second_shape['left'] + second_shape['width'] and
            first_shape['top'] + first_shape['height'] > second_shape['top'] and
            first_shape['top'] < second_shape['top'] + second_shape['height']):
        return True
    return False


def get_download_path():
    """Returns the default downloads path for linux or windows"""
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')


def dict_to_string(dictionary, qml_color_wrongs=None):
    if qml_color_wrongs:
        qml_result = []
        for key in dictionary:
            if dictionary[key] == "Не выполнено":
                qml_result.append(f"{key}: <font color='red'>{dictionary[key]}</font>.")
            else:
                qml_result.append(f"{key}: <font color='green'>{dictionary[key]}</font>.")
        return qml_result
    else:
        result = []
        for key in dictionary:
            result.append(f"{key}: {dictionary[key]}")
        return '\n'.join(result)


def open_presentation(path):
    Application.Presentations.Open(path)


def upload_images(from_directory):
    f_dir = Path(from_directory).resolve()
    if f_dir.is_dir():
        path = Path('original_images').resolve()
        if path.exists():
            shutil.rmtree(path, ignore_errors=True)
        path.mkdir(parents=True)
        for filename in f_dir.iterdir():
            if filename.suffix in ['.png', '.jpg', '.jpeg']:
                shutil.copy(filename, path, follow_symlinks=True)
        return True
    return False  # TODO generate expression here


def layout_to_dict(width, height, lt="DEFAULT"):
    result = {
        2: {"title": [], "images": [], "text": []},
        3: {"title": [], "images": [], "text": []},
    }
    for layout in layouts:
        if layout == lt:
            for place in layouts[layout]:
                slide, properties = int(place[-1]), layouts[layout][place].split("|")
                for prop in properties:
                    prop, temp, = prop.split(","), []
                    for p in prop:
                        if "e" in p:
                            temp.append(0)
                        elif "w" in p:
                            if "*" in p:
                                temp.append(float((width / int(p[0])) * int(p[2])))
                            else:
                                temp.append(float(width / int(p[0])))
                        elif "h" in p:
                            if "*" in p:
                                temp.append(float((height / int(p[0])) * int(p[2])))
                            else:
                                temp.append(float(height / int(p[0])))
                    result[slide][place.split("_")[0]].append({
                        "left": temp[0],
                        "top": temp[1],
                        "width": temp[2],
                        "height": temp[3],
                    })
            return result
    return False
