import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent))

from tools import *

def test_execute_command():
    print(execute_command(None, "LINE 0,0 100,100 "))

def test_create_block():
    print(create_block(None, "TestBlock", ["ABC123", "DEF456"], [0, 0, 0], True))

def test_autocad_select_objects():
    print(autocad_select_objects(None, None, "X", None))

def test_autocad_eval_lisp():
    print(autocad_eval_lisp(None, "(+ 1 2 3)"))

def test_autocad_get_sysvar():
    print(autocad_get_sysvar(None, "DWGNAME"))

def test_autocad_set_sysvar():
    print(autocad_set_sysvar(None, "CMDECHO", 0))


if __name__ == "__main__":
    test_execute_command()
    test_create_block()
    test_autocad_select_objects()
    test_autocad_eval_lisp()
    test_autocad_get_sysvar()
    test_autocad_set_sysvar()
