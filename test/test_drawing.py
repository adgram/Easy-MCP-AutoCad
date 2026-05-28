import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

from tools import *

def test_create_new_drawing():
    print(create_new_drawing(None, "TestDrawing"))



def test_create_layer():
    print(create_layer(None, "TestLayer"))

def test_draw_line():
    print(draw_line(None, 0, 0, 100, 100, "TestLayer"))

def test_draw_circle():
    print(draw_circle(None, 50, 50, 25, "TestLayer"))

def test_draw_polyline():
    print(draw_polyline(None, [(0, 0), (100, 100), (200, 0)], "TestLayer"))

def test_draw_rectangle():
    print(draw_rectangle(None, 0, 0, 100, 50, "TestLayer"))

def test_draw_text():
    print(draw_text(None, 50, 50, "Hello, AutoCAD!", "TestLayer"))

def test_autocad_get_layers():
    print(autocad_get_layers(None))

def test_autocad_get_blocks():
    print(autocad_get_blocks(None))

def test_autocad_get_linetypes():
    print(autocad_get_linetypes(None))


if __name__ == "__main__":
    # test_create_new_drawing()
    test_create_layer()
    test_draw_line()
    test_draw_circle()
    test_draw_polyline()
    test_draw_rectangle()
    test_draw_text()
    test_autocad_get_layers()
    test_autocad_get_blocks()
    test_autocad_get_linetypes()