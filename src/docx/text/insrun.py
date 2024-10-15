# encoding: utf-8

"""
Run-related proxy objects for python-docx, Run in particular.
"""

from __future__ import absolute_import, print_function, unicode_literals

from docx.shared import StoryChild
from docx.styles.style import CharacterStyle
from docx.text.run import Run


class Ins(StoryChild):
    """
    An insRun object
    """

    def __init__(self, i, parent):
        super(Ins, self).__init__(parent)
        self._i = self._element = self.element = i

    def add_run(self, text: str | None, style: str | CharacterStyle | None = None) -> Run:
        """
        Append a run to this w:ins containing `text` and having character
        style identified by style ID `style`. `text` can contain tab
        (`\\t`) characters, which are converted to the appropriate XML form
        for a tab. `text` can also include newline (`\\n`) or carriage
        return (`\\r`) characters, each of which is converted to a line
        break.
        """
        r = self._i.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    def add_text(self, text):
        t = self._i.add_t(text)
        return _Text(t)

    def text(self, text):
        self._i.text = text

    @property
    def rpr(self):
        return self._i.r_lst[0].rpr

    @rpr.setter
    def rpr(self, value):
        for r in self._i.r_lst:
            r.rpr = value

    @property
    def all_runs(self):
        return [Run(r, self) for r in self._i.xpath(".//w:r")]


class _Text(object):
    """
    Proxy object wrapping ``<w:t>`` element.
    """

    def __init__(self, t_elm):
        super(_Text, self).__init__()
        self._dt = t_elm
