# DWG.py
import re


class DWG:
    def __init__(self, name, pkg, ctrl_model, ctrl_size, act_signal, compact=False, sweat=False, large=False, ss=False,
                 press=False, stacked=False, isolation=False):
        self.name: str = name
        self.pkg: str = pkg
        self.ctrl_model: str = ctrl_model
        self.ctrl_size: str = ctrl_size
        self.act_signal: str = act_signal
        self.compact: bool = compact
        self.sweat: bool = sweat
        self.large: bool = large
        self.ss: bool = ss
        self.press: bool = press
        self.stacked: bool = stacked
        self.isolation: bool = isolation

        self.fpath = None

    def __eq__(self, other):
        if isinstance(other, DWG):
            return self.name == other.name and self.pkg == other.pkg
        return False

    def __str__(self):
        return self.name

    def setfilepath(self, fpath):
        self.fpath = fpath

    def setpkgkey(self, pkg):
        self.pkg = pkg

    def setcompact(self):
        self.compact = True

    def setsweat(self):
        self.sweat = True

    def setlarge(self):
        self.large = True

    def setss(self):
        self.ss = True

    def setpress(self):
        self.press = True

    def setstacked(self):
        self.stacked = True

    def setiso(self):
        self.isolation = True

    def parts(self):
        return re.findall(r"[\w'+=]+", self.name)

