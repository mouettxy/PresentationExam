import configparser
from pathlib import Path


def get_config():
    config = configparser.ConfigParser()
    config.read(Path.joinpath(Path(__file__).parent, 'config.ini'))
    return config


def get_constants():
    config = configparser.ConfigParser()
    config.read(Path.joinpath(Path(__file__).parent, 'config.ini'))
    return config['CONSTANTS']


def get_layouts():
    layouts = configparser.ConfigParser()
    layouts.read(Path.joinpath(Path(__file__).parent, 'layouts.ini'))
    return layouts


def get_analyze():
    analyze = configparser.ConfigParser()
    analyze.read(Path.joinpath(Path(__file__).parent, 'config.ini'))
    return analyze['ANALYZE']


def modify_analyze(what_to_modify):
    cfg = get_config()
    for k in what_to_modify:
        cfg.set("ANALYZE", k, what_to_modify[k])
    with open(Path.joinpath(Path(__file__).parent, 'config.ini'), "w") as config_file:
        cfg.write(config_file)


def add_layout(name, layout_props):
    '''
    EXAMPLE LAYOUT_PROPS
    {
    "title_2": [[("0", "e"), ("0", "e"), ("1", "w"), ("4", "h")]],
    "images_2": [[("0", "e"), ("0", "e"), ("2", "w"), ("1", "h")]],
    "text_2": [[("2", "w"), ("0", "e"), ("2", "w"), ("1", "h")]],
    "title_3": [[("0", "e"), ("0", "e"), ("1", "w"), ("4", "h")]],
    "images_3": [[("0", "e"), ("0", "e"), ("3", "w"), ("2", "h")],
                 [("3", "w"), ("2", "h"), ("3", "w"), ("2", "h")],
                 [("3*2", "w"), ("0", "e"), ("3", "w"), ("2", "h")]],
    "text_3": [[("0", "e"), ("2", "h"), ("3", "w"), ("2", "h")],
               [("3", "w"), ("0", "e"), ("3", "w"), ("2", "h")],
               [("3*2", "w"), ("2", "h"), ("3", "w"), ("2", "h")]],
    }
    '''
    cfg = get_layouts()
    try:
        cfg.add_section(name)
    except configparser.DuplicateSectionError:
        pass
    for k in layout_props:
        result_str = ""
        for j in layout_props[k]:
            result_str += f"{j[0][0]}-{j[0][1]},"
            result_str += f"{j[1][0]}-{j[1][1]},"
            result_str += f"{j[2][0]}-{j[2][1]},"
            result_str += f"{j[3][0]}-{j[3][1]}"
            if layout_props[k].index(j) != len(layout_props[k]) - 1:
                result_str += "|"
        cfg.set(name, k, result_str)
    with open(Path.joinpath(Path(__file__).parent, 'layouts.ini'), "w") as config_file:
        cfg.write(config_file)
