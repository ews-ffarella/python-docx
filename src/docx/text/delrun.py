# encoding: utf-8

"""
Run-related proxy objects for python-docx, Run in particular.
"""

from __future__ import absolute_import, print_function, unicode_literals

from docx.shared import StoryChild
from docx.text.run import Run


class Del(StoryChild):
    """
    A delRun object
    """

    def __init__(self, d, parent):
        super(Del, self).__init__(parent)
        self._d = self._element = self.element = d

    def add_run(self, text, style):
        """
        Append a run to this w:del containing `text` and having character
        style identified by style ID `style`. `text` can contain tab
        (`\\t`) characters, which are converted to the appropriate XML form
        for a tab. `text` can also include newline (`\\n`) or carriage
        return (`\\r`) characters, each of which is converted to a line
        break.
        """
        r = self._d.add_r()
        run = Run(r, self)
        if text:
            run.deltext = text
        if style:
            run.style = style
        return run

    def add_text(self, text):
        t = self._d.add_dt(text)
        return _Text(t)

    def text(self, text):
        self._d.text = text

    @property
    def rpr(self):
        return self._d.r_lst[0].rpr

    @rpr.setter
    def rpr(self, value):
        for r in self._d.r_lst:
            r.rpr = value

    @property
    def all_runs(self):
        return [Run(r, self) for r in self._d.xpath("./w:r")]


class _Text(object):
    """
    Proxy object wrapping ``<w:delText>`` element.
    """

    def __init__(self, t_elm):
        super(_Text, self).__init__()
        self._dt = t_elm
