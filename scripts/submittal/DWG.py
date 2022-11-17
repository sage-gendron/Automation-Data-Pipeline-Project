# DWG.py
import re
"""
author: Sage Gendron
Object representing an engineering drawing to help aid in submittal generation.

This object was created in a full refactor of the product_submittal.py file and in seeking to abstract/accelerate
the submittal generation process as initially it was quite slow. The creation of this object would have eventually led 
to a full refactor of the product_quote.py file along with implementing some object-oriented practices there as well,
but it was not destined to be.
"""


class DWG:
    def __init__(self, name, pkg, ctrl_model, ctrl_size, signal, sm=False, lg=False, sp_case_1=False, sp_case_2=False):
        self.name: str = name
        self.pkg: str = pkg
        self.ctrl_model: str = ctrl_model
        self.ctrl_size: str = ctrl_size
        self.signal: str = signal
        self.sm: bool = sm
        self.lg: bool = lg
        self.sp_case_1: bool = sp_case_1
        self.sp_case_2: bool = sp_case_2

        self.fpath = None

    def __eq__(self, other):
        if isinstance(other, DWG):
            return self.name == other.name and self.pkg == other.pkg
        return False

    def __str__(self):
        return self.name

    def set_fpath(self, fpath):
        self.fpath = fpath

    def set_pkg(self, pkg):
        self.pkg = pkg

    def set_sm(self):
        self.sm = True

    def set_lg(self):
        self.lg = True

    def set_sp_case_1(self):
        self.sp_case_1 = True

    def set_sp_case_2(self):
        self.sp_case_2 = True

    def parts(self):
        return re.findall(r"[\w'+=]+", self.name)

