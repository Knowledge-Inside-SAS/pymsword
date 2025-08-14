"""Various inserters that use COM interface to manipulate documsnts."""
import os
from datetime import datetime
from typing import List, Callable


def table_inserter(data:List[List[str]])->Callable[[object], object]:
    """This function creates an "Inserter" - a function that takes Word.Range object
    And inserts a table into the document at the specified location
    :param data: A list of lists, where each inner list represents a row in the table.
    :return: A function that takes a Word.Range object and inserts a table into the document.
    """
    nrows = len(data)
    ncols = 0 if nrows == 0 else len(data[0])

    def inserter(word_range):
        #THis function accetps a Word.Range object and inserts a table into the document
        if nrows == 0 or ncols == 0:
            return
        table = word_range.Document.Tables.Add(word_range,nrows,ncols,1,2) # wdWord9TableBehavior, wdAutoFitWindow
        for i, row in enumerate(data):
            for j, cell in enumerate(row):
                # Set the text of the cell
                table.Cell(i + 1, j + 1).Range.Text = str(cell)
        return

    return inserter


def _insert_document_inline(word_range, filename:str, set_style=None):
    """Inserts content of the RTF file into specified range of document.
    After operation, range points after the inserted RTF."""
    word_range.Collapse()  # wdCollapseStart
    rngEnd = word_range.Duplicate
    rngEnd.InsertParagraph()
    word_range.InsertFile(os.path.abspath(filename), ConfirmConversions=False)
    if set_style is not None:
        # Set the style of the inserted content
        word_range.Style = word_range.Document.Styles(set_style)
    rngEnd.Characters.Last.Delete()
    rngEnd.Characters.Last.Delete()
    word_range.SetRange(rngEnd.Start, rngEnd.End)

    word_range.Collapse(0)  # wdCollapseEnd

    return word_range

def _insert_picture_inline(insert_range, filename):
    """Inserts picture from file <filename> into specified range of document.
    After operation, range points after the inserted picture.
    Function returns a new Word.InlineShape object."""
    pic = insert_range.InlineShapes.AddPicture (os.path.abspath(filename))
    if not pic:
        raise ValueError("AddPicture returned None, check the file path")
    pic_range = pic.Range
    insert_range.SetRange(pic_range.Start, pic_range.End)
    insert_range.Collapse(0) #wdCollapseEnd
    return pic

def image_inserter(picture_path:str)->Callable[[object], object]:
    """Returns a function that inserts an image into the document at the specified range."""
    def insert(word_range):
        if not os.path.exists(picture_path): raise ValueError("Image file was not found")
        _insert_picture_inline(word_range, picture_path)
    return insert

def document_inserter(rtf_path:str, set_style=None)->Callable[[object], object]:
    """Returns a function that inserts RTF content into the document at the specified range."""
    def insert(word_range):
        if not os.path.exists(rtf_path): raise ValueError("RTF file was not found")
        _insert_document_inline(word_range, rtf_path, set_style = set_style)
    return insert


def anchor_inserter(text:str, anchor:str):
    """Returns a function that inserts a bookmark into the document at the specified range.
    :param text: The text to insert at the bookmark.
    :param anchor: The name of the bookmark to create.
    """
    def insert(word_range):
        word_range.Text = text
        bmk = word_range.Document.Bookmarks.Add(anchor, word_range)
        return bmk
    return insert


def heading_inserter(text:str, level:int=1)->Callable[[object], object]:
    """Returns a function that inserts a heading into the document at the specified range.
    :param text: The text of the heading.
    :param level: The heading level (1 for Heading 1, 2 for Heading 2, etc.).
    """
    def insert(word_range):
        word_range.text = text
        word_range.ParagraphFormat.Style = -2 - level
    return insert


def update_document_toc(document):
    """Post-processing function that updates table of content in the document."""
    document.BuiltinDocumentProperties("Creation Date").Value = datetime.now()
    for toc in document.TablesOfContents:
        toc.Update()
