"""Custom element classes related to text runs (CT_R)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Iterator, List

from docx.oxml.drawing import CT_Drawing
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.simpletypes import ST_BrClear, ST_BrType
from docx.oxml.text.font import CT_RPr
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrMore, ZeroOrOne
from docx.shared import TextAccumulator

if TYPE_CHECKING:
    from docx.oxml.shape import CT_Anchor, CT_Inline
    from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak
    from docx.oxml.text.parfmt import CT_TabStop

# ------------------------------------------------------------------------------------
# Run-level elements


class CT_R(BaseOxmlElement):
    """`<w:r>` element, containing the properties and text for a run."""

    add_br: Callable[[], CT_Br]
    add_tab: Callable[[], CT_TabStop]
    get_or_add_rPr: Callable[[], CT_RPr]
    _add_drawing: Callable[[], CT_Drawing]
    _add_t: Callable[..., CT_Text]

    rPr: CT_RPr | None = ZeroOrOne("w:rPr")  # pyright: ignore[reportAssignmentType]
    br = ZeroOrMore("w:br")
    cr = ZeroOrMore("w:cr")
    drawing = ZeroOrMore("w:drawing")
    t = ZeroOrMore("w:t")
    dt = ZeroOrOne("w:delText")
    tab = ZeroOrMore("w:tab")

    def add_t(self, text: str) -> CT_Text:
        """Return a newly added `<w:t>` element containing `text`."""
        t = self._add_t(text=text)
        if len(text.strip()) < len(text):
            t.set(qn("xml:space"), "preserve")
        return t

    def add_dt(self, text):
        """
        Return a newly added ``<w:delText>`` element containing *text*.
        """
        dt = self._add_dt(text=text)
        if len(text.strip()) < len(text):
            dt.set(qn("xml:space"), "preserve")
        return dt

    def add_drawing(self, inline_or_anchor: CT_Inline | CT_Anchor) -> CT_Drawing:
        """Return newly appended `CT_Drawing` (`w:drawing`) child element.

        The `w:drawing` element has `inline_or_anchor` as its child.
        """
        drawing = self._add_drawing()
        drawing.append(inline_or_anchor)
        return drawing

    def clear_content(self) -> None:
        """Remove all child elements except a `w:rPr` element if present."""
        # -- remove all run inner-content except a `w:rPr` when present. --
        for e in self.xpath("./*[not(self::w:rPr)]"):
            self.remove(e)

    def add_comment(self, author, comment_part, initials, dtime, comment_text):
        comment = comment_part.add_comment(author, initials, dtime)
        comment._add_p(comment_text)
        # _r = self.add_r()
        self.link_comment(comment._id)

        return comment

    def link_comment(self, _id):
        self.mark_comment_start(_id)
        self.mark_comment_end(_id)

    def mark_comment_start(self, _id):
        """Mark run as start of comment with CRS

        <w:commentRangeStart w:id=_id/>
        <self run element>
        """
        comm_start = OxmlElement("w:commentRangeStart")
        comm_start._id = _id
        self.addprevious(comm_start)

    def mark_comment_end(self, _id):
        """Mark run as end of a comment with CRE and ref.

        <self run element>
        <w:commentRangeEnd w:id=_id/>
        <w:r>
            <w:commentReference w:id=_id/>
        </w:r>
        """
        comm_end = OxmlElement("w:commentRangeEnd")
        comm_end._id = _id
        self.addnext(comm_end)
        reference = OxmlElement("w:commentReference")
        reference._id = _id
        ref_run_elem = OxmlElement("w:r")
        ref_run_elem.append(reference)
        comm_end.addnext(ref_run_elem)

    def add_footnote_reference(self, _id):
        rPr = self.get_or_add_rPr()
        rstyle = rPr.get_or_add_rStyle()
        rstyle.val = "FootnoteReference"
        reference = OxmlElement("w:footnoteReference")
        reference._id = _id
        self.append(reference)
        return reference

    def add_footnoteRef(self):
        ref = OxmlElement("w:footnoteRef")
        self.append(ref)

        return ref

    def footnote_style(self):
        rPr = self.get_or_add_rPr()
        rstyle = rPr.get_or_add_rStyle()
        rstyle.val = "FootnoteReference"

        self.add_footnoteRef()
        return self

    @property
    def footnote_id(self):
        _id = self.xpath("./w:footnoteReference/@w:id")
        if len(_id) > 1 or len(_id) == 0:
            return None
        else:
            return int(_id[0])

    def clear_content(self):
        """
        Remove all child elements except the ``<w:rPr>`` element if present.
        """
        content_child_elms = self[1:] if self.rPr is not None else self[:]
        for child in content_child_elms:
            self.remove(child)

    def add_comment_reference(self, _id):
        reference = OxmlElement("w:commentReference")
        reference._id = _id
        self.append(reference)
        return reference

    def copy_rpr(self, rprCopy):
        rPr = self.get_or_add_rPr()
        if rprCopy is not None:
            for p in rprCopy[:]:
                rPr.append(p)

    @property
    def inner_content_items(self) -> List[str | CT_Drawing | CT_LastRenderedPageBreak]:
        """Text of run, possibly punctuated by `w:lastRenderedPageBreak` elements."""
        from docx.oxml.text.pagebreak import CT_LastRenderedPageBreak

        accum = TextAccumulator()

        def iter_items() -> Iterator[str | CT_Drawing | CT_LastRenderedPageBreak]:
            for e in self.xpath(
                "w:br"
                " | w:cr"
                " | w:drawing"
                " | w:lastRenderedPageBreak"
                " | w:noBreakHyphen"
                " | w:ptab"
                " | w:t"
                " | w:tab"
            ):
                if isinstance(e, (CT_Drawing, CT_LastRenderedPageBreak)):
                    yield from accum.pop()
                    yield e
                else:
                    accum.push(str(e))

            # -- don't forget the "tail" string --
            yield from accum.pop()

        return list(iter_items())

    @property
    def lastRenderedPageBreaks(self) -> List[CT_LastRenderedPageBreak]:
        """All `w:lastRenderedPageBreaks` descendants of this run."""
        return self.xpath("./w:lastRenderedPageBreak")

    @property
    def rpr(self):
        return self.rPr

    @rpr.setter
    def rpr(self, value):
        self.copy_rpr(value)

    @property
    def style(self) -> str | None:
        """String contained in `w:val` attribute of `w:rStyle` grandchild.

        |None| if that element is not present.
        """
        rPr = self.rPr
        if rPr is None:
            return None
        return rPr.style

    @style.setter
    def style(self, style: str | None):
        """Set character style of this `w:r` element to `style`.

        If `style` is None, remove the style element.
        """
        rPr = self.get_or_add_rPr()
        rPr.style = style

    @property
    def text(self) -> str:
        """The textual content of this run.

        Inner-content child elements like `w:tab` are translated to their text
        equivalent.
        """
        return "".join(
            str(e) for e in self.xpath("w:br | w:cr | w:noBreakHyphen | w:ptab | w:t | w:tab")
        )

    @text.setter
    def text(self, text: str):  # pyright: ignore[reportIncompatibleMethodOverride]
        self.clear_content()
        _RunContentAppender.append_to_run_from_text(self, text)

    def _insert_rPr(self, rPr: CT_RPr) -> CT_RPr:
        self.insert(0, rPr)
        return rPr

    @property
    def deltext(self):
        """
        A string representing the textual content of this run, with content
        child elements like ``<w:tab/>`` translated to their Python
        equivalent.
        """
        text = ""
        for child in self:
            if child.tag == qn("w:delText"):
                t_text = child.text
                text += t_text if t_text is not None else ""
            elif child.tag == qn("w:tab"):
                text += "\t"
            elif child.tag in (qn("w:br"), qn("w:cr")):
                text += "\n"
        return text

    @deltext.setter
    def deltext(self, text):
        self.clear_content()
        _DelRunContentAppender.append_to_run_from_text(self, text)


# ------------------------------------------------------------------------------------
# Run inner-content elements


class CT_Br(BaseOxmlElement):
    """`<w:br>` element, indicating a line, page, or column break in a run."""

    type: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "w:type", ST_BrType, default="textWrapping"
    )
    clear: str | None = OptionalAttribute("w:clear", ST_BrClear)  # pyright: ignore

    def __str__(self) -> str:
        """Text equivalent of this element. Actual value depends on break type.

        A line break is translated as "\n". Column and page breaks produce the empty
        string ("").

        This allows the text of run inner-content to be accessed in a consistent way
        for all run inner-context text elements.
        """
        return "\n" if self.type == "textWrapping" else ""


class CT_Cr(BaseOxmlElement):
    """`<w:cr>` element, representing a carriage-return (0x0D) character within a run.

    In Word, this represents a "soft carriage-return" in the sense that it does not end
    the paragraph the way pressing Enter (aka. Return) on the keyboard does. Here the
    text equivalent is considered to be newline ("\n") since in plain-text that's the
    closest Python equivalent.

    NOTE: this complex-type name does not exist in the schema, where `w:tab` maps to
    `CT_Empty`. This name was added to give it distinguished behavior. CT_Empty is used
    for many elements.
    """

    def __str__(self) -> str:
        """Text equivalent of this element, a single newline ("\n")."""
        return "\n"


class CT_NoBreakHyphen(BaseOxmlElement):
    """`<w:noBreakHyphen>` element, a hyphen ineligible for a line-wrap position.

    This maps to a plain-text dash ("-").

    NOTE: this complex-type name does not exist in the schema, where `w:noBreakHyphen`
    maps to `CT_Empty`. This name was added to give it behavior distinguished from the
    many other elements represented in the schema by CT_Empty.
    """

    def __str__(self) -> str:
        """Text equivalent of this element, a single dash character ("-")."""
        return "-"


class CT_PTab(BaseOxmlElement):
    """`<w:ptab>` element, representing an absolute-position tab character within a run.

    This character advances the rendering position to the specified position regardless
    of any tab-stops, perhaps for layout of a table-of-contents (TOC) or similar.
    """

    def __str__(self) -> str:
        """Text equivalent of this element, a single tab ("\t") character.

        This allows the text of run inner-content to be accessed in a consistent way
        for all run inner-context text elements.
        """
        return "\t"


# -- CT_Tab functionality is provided by CT_TabStop which also uses `w:tab` tag. That
# -- element class provides the __str__() method for this empty element, unconditionally
# -- returning "\t".


class CT_Text(BaseOxmlElement):
    """`<w:t>` element, containing a sequence of characters within a run."""

    def __str__(self) -> str:
        """Text contained in this element, the empty string if it has no content.

        This property allows this run inner-content element to be queried for its text
        the same way as other run-content elements are. In particular, this never
        returns None, as etree._Element does when there is no content.
        """
        return self.text or ""


class CT_DelText(BaseOxmlElement):
    """
    ``<w:delText>`` element, containing a sequence of characters within a run
    """


class _DelRunContentAppender(object):
    """
    Service object that knows how to translate a Python string into run
    content elements appended to a specified ``<w:r>`` element. Contiguous
    sequences of regular characters are appended in a single ``<w:delText>``
    element. Each tab character ('\t') causes a ``<w:tab/>`` element to be
    appended. Likewise a newline or carriage return character ('\n', '\r')
    causes a ``<w:cr>`` element to be appended.
    """

    def __init__(self, r):
        self._r = r
        self._bfr = []

    @classmethod
    def append_to_run_from_text(cls, r, text):
        """
        Create a "one-shot" ``_RunContentAppender`` instance and use it to
        append the run content elements corresponding to *text* to the
        ``<w:r>`` element *r*.
        """
        appender = cls(r)
        appender.add_deltext(text)

    def add_deltext(self, text):
        """
        Append the run content elements corresponding to *text* to the
        ``<w:r>`` element of this instance.
        """
        for char in text:
            self.add_char(char)
        self.flush()

    def add_char(self, char):
        """
        Process the next character of input through the translation finite
        state maching (FSM). There are two possible states, buffer pending
        and not pending, but those are hidden behind the ``.flush()`` method
        which must be called at the end of text to ensure any pending
        ``<w:t>`` element is written.
        """
        if char == "\t":
            self.flush()
            self._r.add_tab()
        elif char in "\r\n":
            self.flush()
            self._r.add_br()
        else:
            self._bfr.append(char)

    def flush(self):
        text = "".join(self._bfr)
        if text:
            self._r.add_dt(text)
        del self._bfr[:]


# ------------------------------------------------------------------------------------
# Utility


class _RunContentAppender:
    """Translates a Python string into run content elements appended in a `w:r` element.

    Contiguous sequences of regular characters are appended in a single `<w:t>` element.
    Each tab character ('\t') causes a `<w:tab/>` element to be appended. Likewise a
    newline or carriage return character ('\n', '\r') causes a `<w:cr>` element to be
    appended.
    """

    def __init__(self, r: CT_R):
        self._r = r
        self._bfr: List[str] = []

    @classmethod
    def append_to_run_from_text(cls, r: CT_R, text: str):
        """Append inner-content elements for `text` to `r` element."""
        appender = cls(r)
        appender.add_text(text)

    def add_text(self, text: str):
        """Append inner-content elements for `text` to the `w:r` element."""
        for char in text:
            self.add_char(char)
        self.flush()

    def add_char(self, char: str):
        """Process next character of input through finite state maching (FSM).

        There are two possible states, buffer pending and not pending, but those are
        hidden behind the `.flush()` method which must be called at the end of text to
        ensure any pending `<w:t>` element is written.
        """
        if char == "\t":
            self.flush()
            self._r.add_tab()
        elif char in "\r\n":
            self.flush()
            self._r.add_br()
        else:
            self._bfr.append(char)

    def flush(self):
        text = "".join(self._bfr)
        if text:
            self._r.add_t(text)
        self._bfr.clear()
