"""Paragraph-related proxy types."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List, cast
from datetime import datetime

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.text.run import CT_R
from docx.shared import StoryChild
from docx.styles.style import ParagraphStyle
from docx.text.hyperlink import Hyperlink
from docx.text.pagebreak import RenderedPageBreak
from docx.text.parfmt import ParagraphFormat
from docx.text.run import Run
from docx.text.delrun import Del
from docx.text.insrun import Ins

if TYPE_CHECKING:
    import docx.types as t
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.text.paragraph import CT_P
    from docx.styles.style import CharacterStyle


class Paragraph(StoryChild):
    """Proxy object wrapping a `<w:p>` element."""

    def __init__(self, p: CT_P, parent: t.ProvidesStoryPart):
        super(Paragraph, self).__init__(parent)
        self._p = self._element = self.element = p

    def add_run(self, text: str | None = None, style: str | CharacterStyle | None = None) -> Run:
        """Append run containing `text` and having character-style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break. When `text` is `None`, the new run is empty.
        """
        r = self._p.add_r()
        run = Run(r, self)
        if text:
            run.text = text
        if style:
            run.style = style
        return run

    def add_ins(
            self,
            text: str | None = None,
            style: str | CharacterStyle | None = None
    ) -> Ins:
        i = self._p.add_i()

        irun = Ins(i, self)
        run = None
        if text:
            run = irun.add_run(text)
        if run is not None and style is not None:
            run.style = style

        return irun

    def add_del(
            self, deltext=None
    ) -> Del:

        d = self._p.add_d()

        drun = Del(d, self)
        if deltext:
            drun.add_text(deltext)

        return drun

    def delete(self):
        """
        delete the content of the paragraph
        """
        self._p.getparent().remove(self._p)
        self._p = self._element = None

    def add_comment(self, text, author='python-docx', initials='pd', dtime=None, rangeStart=0, rangeEnd=0,
                    comment_part=None):
        if comment_part is None:
            comment_part = self.part._comments_part.element
        if dtime is None:
            dtime = str(datetime.now()).replace(' ', 'T')
        comment = self._p.add_comm(author, comment_part, initials, dtime, text, rangeStart, rangeEnd)

        return comment

    def add_footnote(self, text):
        footnotes_part = self.part._footnotes_part.element
        footnote = self._p.add_fn(text, footnotes_part)

        return footnote

    def merge_paragraph(self, otherParagraph):
        r_lst = otherParagraph.runs
        self.append_runs(r_lst)

    def append_runs(self, runs):
        self.add_run(' ')
        for run in runs:
            self._p.append(run._r)

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """A member of the :ref:`WdParagraphAlignment` enumeration specifying the
        justification setting for this paragraph.

        A value of |None| indicates the paragraph has no directly-applied alignment
        value and will inherit its alignment value from its style hierarchy. Assigning
        |None| to this property removes any directly-applied alignment value.
        """
        return self._p.alignment

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT):
        self._p.alignment = value

    def clear(self):
        """Return this same paragraph after removing all its content.

        Paragraph-level formatting, such as style, is preserved.
        """
        self._p.clear_content()
        return self

    @property
    def contains_page_break(self) -> bool:
        """`True` when one or more rendered page-breaks occur in this paragraph."""
        return bool(self._p.lastRenderedPageBreaks)

    @property
    def hyperlinks(self) -> List[Hyperlink]:
        """A |Hyperlink| instance for each hyperlink in this paragraph."""
        return [Hyperlink(hyperlink, self) for hyperlink in self._p.hyperlink_lst]

    def insert_paragraph_before(
        self, text: str | None = None, style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return a newly created paragraph, inserted directly before this paragraph.

        If `text` is supplied, the new paragraph contains that text in a single run. If
        `style` is provided, that style is assigned to the new paragraph.
        """
        paragraph = self._insert_paragraph_before()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def insert_paragraph_after(self, text=None, style=None):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph. If *text* is supplied, the new paragraph contains that
        text in a single run. If *style* is provided, that style is assigned
        to the new paragraph.
        """
        paragraph = self._insert_paragraph_after()
        if text:
            paragraph.add_run(text)
        if style is not None:
            paragraph.style = style
        return paragraph

    def iter_inner_content(self) -> Iterator[Run | Hyperlink]:
        """Generate the runs and hyperlinks in this paragraph, in the order they appear.

        The content in a paragraph consists of both runs and hyperlinks. This method
        allows accessing each of those separately, in document order, for when the
        precise position of the hyperlink within the paragraph text is important. Note
        that a hyperlink itself contains runs.
        """
        for r_or_hlink in self._p.inner_content_elements:
            yield (
                Run(r_or_hlink, self)
                if isinstance(r_or_hlink, CT_R)
                else Hyperlink(r_or_hlink, self)
            )

    @property
    def paragraph_format(self):
        """The |ParagraphFormat| object providing access to the formatting properties
        for this paragraph, such as line spacing and indentation."""
        return ParagraphFormat(self._element)

    @property
    def rendered_page_breaks(self) -> List[RenderedPageBreak]:
        """All rendered page-breaks in this paragraph.

        Most often an empty list, sometimes contains one page-break, but can contain
        more than one is rare or contrived cases.
        """
        return [RenderedPageBreak(lrpb, self) for lrpb in self._p.lastRenderedPageBreaks]

    @property
    def runs(self) -> List[Run]:
        """Sequence of |Run| instances corresponding to the <w:r> elements in this
        paragraph."""
        return [Run(r, self) for r in self._p.r_lst]

    @property
    def all_runs(self):
        return [Run(r, self) for r in self._p.xpath('.//w:r[not(ancestor::w:r)]')]

    @property
    def ins(self) -> List[Ins]:
        """
        Sequence of |Ins| instances corresponding to the <w:ins> elements in this paragraph
        """
        return [Ins(i, self) for i in self._p.i_lst]

    @property
    def dels(self) -> List[Del]:
        """
        Sequence of |Del| instances corresponding to the <w:del> elements in this paragraph
        """
        return [Del(d, self) for d in self._p.d_lst]

    @property
    def style(self) -> ParagraphStyle | None:
        """Read/Write.

        |_ParagraphStyle| object representing the style assigned to this paragraph. If
        no explicit style is assigned to this paragraph, its value is the default
        paragraph style for the document. A paragraph style name can be assigned in lieu
        of a paragraph style object. Assigning |None| removes any applied style, making
        its effective value the default paragraph style for the document.
        """
        style_id = self._p.style
        style = self.part.get_style(style_id, WD_STYLE_TYPE.PARAGRAPH)
        return cast(ParagraphStyle, style)

    @style.setter
    def style(self, style_or_name: str | ParagraphStyle | None):
        style_id = self.part.get_style_id(style_or_name, WD_STYLE_TYPE.PARAGRAPH)
        self._p.style = style_id

    @property
    def text(self) -> str:
        """The textual content of this paragraph.

        The text includes the visible-text portion of any hyperlinks in the paragraph.
        Tabs and line breaks in the XML are mapped to ``\\t`` and ``\\n`` characters
        respectively.

        Assigning text to this property causes all existing paragraph content to be
        replaced with a single run containing the assigned text. A ``\\t`` character in
        the text is mapped to a ``<w:tab/>`` element and each ``\\n`` or ``\\r``
        character is mapped to a line break. Paragraph-level formatting, such as style,
        is preserved. All run-level formatting, such as bold or italic, is removed.
        """
        return self._p.text

    @property
    def header_level(self):
        '''
        input Paragraph Object
        output Paragraph level in case of header or returns None
        '''
        headerPattern = re.compile(".*Heading (\d+)$")
        level = 0
        if headerPattern.match(self.style.name):
            level = int(self.style.name.lower().split('heading')[-1].strip())
        return level

    @property
    def NumId(self):
        '''
        returns NumId val in case of paragraph has numbering
        else: return None
        '''
        try:
            return self._p.pPr.numPr.numId.val
        except:
            return None

    @property
    def list_lvl(self):
        '''
        returns ilvl val in case of paragraph has a numbering level
        else: return None
        '''
        try:
            return self._p.pPr.numPr.ilvl.val
        except:
            return None

    @property
    def list_info(self):
        '''
        returns tuple (has numbering info, numId value, ilvl value)
        '''
        if self.NumId and self.list_lvl:
            return True, self.NumId, self.list_lvl
        else:
            return False, 0, 0

    @property
    def is_heading(self):
        return True if self.header_level else False

    @property
    def full_text(self):
        return u"".join([r.text for r in self.all_runs])

    @property
    def footnotes(self):
        if self._p.footnote_ids is not None:
            return True
        else:
            return False

    @property
    def comments(self):
        runs_comments = [run.comments for run in self.all_runs]
        return [comment for comments in runs_comments for comment in comments]

    @text.setter
    def text(self, text: str | None):
        self.clear()
        self.add_run(text)

    def _insert_paragraph_before(self):
        """Return a newly created paragraph, inserted directly before this paragraph."""
        p = self._p.add_p_before()
        return Paragraph(p, self._parent)

    def _insert_paragraph_after(self):
        """
        Return a newly created paragraph, inserted directly before this
        paragraph.
        """
        p = self._p.add_p_after()
        return Paragraph(p, self._parent)
