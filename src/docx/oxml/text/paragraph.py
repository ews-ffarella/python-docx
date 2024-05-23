# pyright: reportPrivateUsage=false

"""Custom element classes related to paragraphs (CT_P)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.parser import OxmlElement
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.section import CT_SectPr
    from docx.oxml.text.hyperlink import CT_Hyperlink
    from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak
    from docx.oxml.text.parfmt import CT_PPr
    from docx.oxml.text.run import CT_R


class CT_P(BaseOxmlElement):
    """`<w:p>` element, containing the properties and text for a paragraph."""

    add_r: Callable[[], CT_R]
    add_d: Callable[[], CT_R]
    add_i: Callable[[], CT_R]
    get_or_add_pPr: Callable[[], CT_PPr]
    hyperlink_lst: List[CT_Hyperlink]
    r_lst: List[CT_R]

    pPr: CT_PPr | None = ZeroOrOne("w:pPr")  # pyright: ignore[reportAssignmentType]
    hyperlink = ZeroOrMore("w:hyperlink")
    r = ZeroOrMore("w:r")
    d = ZeroOrMore("w:del")
    i = ZeroOrMore("w:ins")

    def add_p_before(self) -> CT_P:
        """Return a new `<w:p>` element inserted directly prior to this one."""
        new_p = cast(CT_P, OxmlElement("w:p"))
        self.addprevious(new_p)
        return new_p

    def add_p_after(self) -> CT_P:
        """
        Return a new ``<w:p>`` element inserted directly after this one.
        """
        new_p = cast(CT_P, OxmlElement('w:p'))
        self.addnext(new_p)
        return new_p

    def link_comment(self, _id, rangeStart=0, rangeEnd=0):
        rStart = OxmlElement('w:commentRangeStart')
        rStart._id = _id
        rEnd = OxmlElement('w:commentRangeEnd')
        rEnd._id = _id
        if rangeStart == 0 and rangeEnd == 0:
            self.insert(0,rStart)
            self.append(rEnd)
        else:
            self.insert(rangeStart,rStart)
            if rangeEnd == len(self.getchildren() ) - 1 :
                self.append(rEnd)
            else:
                self.insert(rangeEnd+1, rEnd)

    def add_comm(self, author, comment_part, initials, dtime, comment_text, rangeStart, rangeEnd):

        comment = comment_part.add_comment(author, initials, dtime)
        comment._add_p(comment_text)
        _r = self.add_r()
        _r.add_comment_reference(comment._id)
        self.link_comment(comment._id, rangeStart= rangeStart, rangeEnd=rangeEnd)

        return comment

    def add_fn(self, text, footnotes_part):
        footnote = footnotes_part.add_footnote()
        footnote._add_p(' '+text)
        _r = self.add_r()
        _r.add_footnote_reference(footnote._id)

        return footnote

    def footnote_style(self):
        pPr = self.get_or_add_pPr()
        rstyle = pPr.get_or_add_pStyle()
        rstyle.val = 'FootnoteText'

        return self

    @property
    def alignment(self) -> WD_PARAGRAPH_ALIGNMENT | None:
        """The value of the `<w:jc>` grandchild element or |None| if not present."""
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.jc_val

    @alignment.setter
    def alignment(self, value: WD_PARAGRAPH_ALIGNMENT):
        pPr = self.get_or_add_pPr()
        pPr.jc_val = value

    def clear_content(self):
        """Remove all child elements, except the `<w:pPr>` element if present."""
        for child in self.xpath("./*[not(self::w:pPr)]"):
            self.remove(child)

    @property
    def inner_content_elements(self) -> List[CT_R | CT_Hyperlink]:
        """Run and hyperlink children of the `w:p` element, in document order."""
        return self.xpath("./w:r | ./w:hyperlink")

    @property
    def lastRenderedPageBreaks(self) -> List[CT_LastRenderedPageBreak]:
        """All `w:lastRenderedPageBreak` descendants of this paragraph.

        Rendered page-breaks commonly occur in a run but can also occur in a run inside
        a hyperlink. This returns both.
        """
        return self.xpath(
            "./w:r/w:lastRenderedPageBreak | ./w:hyperlink/w:r/w:lastRenderedPageBreak"
        )

    def set_sectPr(self, sectPr: CT_SectPr):
        """Unconditionally replace or add `sectPr` as grandchild in correct sequence."""
        pPr = self.get_or_add_pPr()
        pPr._remove_sectPr()
        pPr._insert_sectPr(sectPr)

    @property
    def style(self) -> str | None:
        """String contained in `w:val` attribute of `./w:pPr/w:pStyle` grandchild.

        |None| if not present.
        """
        pPr = self.pPr
        if pPr is None:
            return None
        return pPr.style

    @property
    def comment_id(self):
        _id = self.xpath('./w:commentRangeStart/@w:id')
        if len(_id) > 1 or len(_id) == 0:
            return None
        else:
            return int(_id[0])

    @property
    def footnote_ids(self):
        _id = self.xpath('./w:r/w:footnoteReference/@w:id')
        if  len(_id) == 0 :
            return None
        else:
            return _id

    def copy_ppr(self,pprCopy):
        pPr = self.get_or_add_pPr()
        if not pprCopy is None:
            for p in pprCopy[:]:
                pPr.append(p)

    @style.setter
    def style(self, style: str | None):
        pPr = self.get_or_add_pPr()
        pPr.style = style

    @property
    def ppr(self):
        return self.pPr

    @ppr.setter
    def ppr(self, value):
        self.copy_ppr(value)

    @property
    def text(self):  # pyright: ignore[reportIncompatibleVariableOverride]
        """The textual content of this paragraph.

        Inner-content child elements like `w:r` and `w:hyperlink` are translated to
        their text equivalent.
        """
        return "".join(e.text for e in self.xpath("w:r | w:hyperlink"))

    def _insert_pPr(self, pPr: CT_PPr) -> CT_PPr:
        self.insert(0, pPr)
        return pPr
