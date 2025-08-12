# coding=utf-8
import sys
from typing import Dict, Any, Callable
from pymsword.xml_template import XMLFragment, XMLRunner
from pymsword.docx_template import DocxTemplate

__all__=["DocxComTemplate"]

class DocxComTemplate(DocxTemplate):
    """Version of the DocxTemplate generator that supports COM-based postprocessing"""
    PLACEHODER_TEMPLATE = "EIWOSCVY{:05d}HFBJSYXO"
    PLACEHOLDER_WORD_SEARCH_PATTERN = "EIWOSCVY[0-9]{5}HFBJSYXO"

    def __init__(self, docx_file):
        super(DocxComTemplate, self).__init__(docx_file)
        self.insert_counter = 0

    def prepare_data(self, data):
        """Data can contain string values, lists and inserters.
        We can't run inserters in the DOCX mode, so instead of running them immediately,
        we replace them with placeholder text,
        store inserters in a array, then load the DOCX in COM mode and run all inserters
        """
        placeholder2inserter = {}
        def new_placeholder():
            self.insert_counter += 1
            return self.PLACEHODER_TEMPLATE.format(self.insert_counter)

        def walk_data(data): #type: (Dict[str, Any])->Dict[str, Any]
            prepared = {}
            for key, value in data.items():
                ukey = str(key)
                if isinstance(value, list):
                    mapped_value = [walk_data(item) for item in value]
                elif isinstance(value, XMLRunner):
                    mapped_value = value
                elif callable(value):
                    #store the inserter
                    placeholder = new_placeholder()
                    placeholder2inserter[placeholder] = value
                    mapped_value = placeholder
                elif isinstance(value, XMLFragment):
                    mapped_value = value
                else:
                    mapped_value = str(value)
                prepared[ukey] = mapped_value
            return prepared
        return walk_data(data), placeholder2inserter

    def generate(self, data, output):
        prepared_data, inserters = self.prepare_data(data)
        DocxTemplate.generate(self, prepared_data, output)
        if inserters:
            self._run_inserters(output, inserters)

    def _run_inserters(self, docx_file, placeholder2inserter): #type: (str, Dict[str, Callable])->None
        """Run inserters in the DOCX file"""
        # Initialize the Word Application
        import pythoncom
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        doc = None

        def find_placeholder_in_range(range_obj):
            placeholders_with_ranges = []
            find = range_obj.Find
            find.ClearFormatting()
            find.Text = self.PLACEHOLDER_WORD_SEARCH_PATTERN
            #accept wildcards
            find.MatchWildcards = True
            while True:
                if not find.Execute():
                    break
                found_range = find.Parent.Duplicate
                matched_text = str(found_range.Text)
                placeholders_with_ranges.append((matched_text, found_range))
            return placeholders_with_ranges

        def find_and_process_markers(range_getter):
            #we need getter because for document it is doc.Range(), for headers and footers it is header.Range
            range_obj = range_getter()
            assert not isinstance(range_obj, str)
            #let's do it in 2 passes.
            # Pass 1
            #First, we find all placeholders and hold Ranges for them in a list
            placeholders_with_ranges = find_placeholder_in_range(range_obj)
            # Pass 2
            for matched_text, found_range in placeholders_with_ranges[::-1]:
                try:
                    inserter = placeholder2inserter.get(matched_text)
                    if inserter is not None:
                        inserter(found_range)
                    else:
                        pass
                except Exception as err:
                    print("Error running inserter:", err, file=sys.stderr)
                self.insert_counter += 1
            #pass 3: remove all placeholders, if they are still there
            range_obj = range_getter()
            placeholders_with_ranges = find_placeholder_in_range(range_obj)
            #remove all found ranges
            for matched_text, found_range in placeholders_with_ranges[::-1]:
                try:
                    found_range.Delete()
                except Exception as err:
                    print("Error replacing placeholder:", err, file=sys.stderr)

        try:
            word.Visible = False
            # Open the DOCX file
            word.DisplayAlerts = 0
            word.ScreenUpdating = False
            doc = word.Documents.Open(docx_file)
            # Search for the marker string
            #find_and_process_marders(lambda : doc.Range())
            # Search for the marker string in headers and footers
            for section in doc.Sections:
                find_and_process_markers(lambda : section.Range)
                for header in section.Headers:
                    find_and_process_markers(lambda : header.Range)
                for footer in section.Footers:
                    find_and_process_markers(lambda : footer.Range)
        finally:
            word.ScreenUpdating = True
            if doc is not None:
                doc.Close(SaveChanges=True)
            # Quit Word
            word.Quit()
