# coding=utf-8
import os
import shutil
import zipfile
from copy import deepcopy
from random import randint
from typing import Dict, Any
from pymsword.xml_template import parse, fill, RE_INSTRUCTION, XMLFragment, XMLRunner
import xml.etree.ElementTree as ET

__all__ = [
    "DocxTemplate", "DocxImageInserter"
]

EXPAND_TAGS = {
    #Container tag to table row
    "row": {"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr"},
    #Container tag to table cell (not really usable)
    "cell": {"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc"},
    #For list item and paragraph
    "p": {"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"},
    "image": {"{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"},
}

def connect_template_ranges(tree):#type: (ET.ElementTree)->None
    """in WOrd XML, the text is split into runs, and each run can have different formatting.
    This interfers with our template engine, so we need to merge all runs that belong to the same template item.

    To do this:
     - take  paragraph tag
     - find all runs that are inside this tag
     - take their unified text
     - find template tags in the unified run text
     - find runs it belongs to them
     - move template code completely to the first run
    """
    #iterate over all w:p tags
    for p in tree.iterfind(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
        #iterate over all text runs of the paragraph
        runs = []
        parts = []
        character_origins = []
        for irun, run in enumerate(p.iterfind("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")):
            text = run.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
            if text is not None:
                runs.append(run)
                parts.append(text.text)
                #this can be done more effectively using binary search, but whatever.
                character_origins.extend([irun]*len(text.text))
        if len(runs) <= 1:
            #nothing to merge, only one run
            continue
        unified_text = "".join(parts)

        #find all RE_INSTRUCTION regexp matches in the unified text
        needs_glue = False
        for match in RE_INSTRUCTION.finditer(unified_text):
            match_start = match.start()
            match_end = match.end()
            first_origin = character_origins[match_start]
            runs_in_range = character_origins[match_start:match_end]
            unique_runs = sorted(set(runs_in_range))
            if len(unique_runs) != 1:
                #mark origins of the matched text to be the same
                character_origins[match_start:match_end] = [first_origin]*(match_end-match_start)
                needs_glue = True
        if needs_glue:
            run_start = 0
            for run_index, run in enumerate(runs):
                cur_pos = run_start
                while cur_pos < len(character_origins) and character_origins[cur_pos] == run_index:
                    cur_pos += 1
                part = unified_text[run_start:cur_pos]
                if part != parts[run_index]:
                    if not part:
                        #when text run became completely empty - just remove it completely
                        p.remove(run)
                    else:
                        run.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t").text = part
                run_start = cur_pos
class DocxContentTypes(object):
    def __init__(self, data):
        self.content_types = {}
        self.default_content_types = {}
        self._orig_data = data
        self._modified = False
        self._parse(data)
    def _parse(self, data):
        # Parse the XML file using ElementTree
        tree = ET.fromstring(data)
        self._root = tree
        # Now you can work with the XML tree, for example:
        for elem in tree.iter():
            if elem.tag == '{http://schemas.openxmlformats.org/package/2006/content-types}Override':
                ctype = elem.attrib['ContentType']
                partname = elem.attrib['PartName']
                if partname[:1] == "/":
                    partname = partname[1:]
                self.content_types[partname] = ctype
            elif elem.tag == '{http://schemas.openxmlformats.org/package/2006/content-types}Default':
                extension = elem.attrib['Extension'].lower()
                ctype = elem.attrib['ContentType']
                self.default_content_types[extension] = ctype
    def get(self, path):
        """For the given path in the archive, returns its content type"""
        ext = path.rsplit('.', 1)[-1].lower()
        ctype = self.default_content_types.get(ext)
        #try to load override, if ext enough is not enough
        ctype = self.content_types.get(path, ctype)
        return ctype

    def add_extension(self, ext, ctype):
        if ext in self.default_content_types:
            return
        self.default_content_types[ext] = ctype
        ET.SubElement(self._root, "{http://schemas.openxmlformats.org/package/2006/content-types}Default",
                        attrib={"Extension":ext.lower(), "ContentType":ctype})
        self._modified = True

    def to_xml(self):
        if not self._modified: return self._orig_data
        # Generate the XML string from the ElementTree
        return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(self._root, encoding="utf-8")


class DocxRels(object):
    def __init__(self, name, data, _isempty=False): #type: (str, str, bool)->None
        self.name = name
        self.root = ET.fromstring(data)
        self._id2rel = {}
        self._target2rel = {}
        self._last_id = 1
        self._modified = False
        self._orig_data = data
        self._isempty = _isempty

        self._parse_ids()

    @staticmethod
    def create_empty(name):
        return DocxRels(name, """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""", _isempty=True)


    def _parse_ids(self):
        for child in self.root.iter():
            if child.tag == "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship":
                rid = child.attrib['Id']
                target = child.attrib['Target']
                self._id2rel[rid] = child
                self._target2rel[target] = child

    def new_id(self): #type: ()->str
        """Create new relation ID"""
        i = self._last_id
        while True:
            newid = "rId{}".format(i)
            i += 1
            if newid in self._id2rel: continue
            self._last_id = i
            return newid

    def new_relationship(self, target, type_):
        if target in self._target2rel: raise ValueError("Already have relationship with target {}".format(target))

        rid = self.new_id()
        child = ET.SubElement(self.root, "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship",
                              attrib = {"Id":rid,
                                        "Type":type_,
                                        "Target":target})
        self._modified = True
        return rid

    def tostr(self):
        if not self._modified:
            if self._isempty:
                return ""
            else:
                return self._orig_data
        else:
            #need to generate XML back
            return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(self.root, encoding="utf-8")

class DocxRelsCollection(object):
    def __init__(self):
        self.path2rels = {}

    def __iter__(self):
        """Iterate over all rels in the collection"""
        return iter(self.path2rels.values())

    def __len__(self):
        """Returns number of rels in the collection"""
        return len(self.path2rels)

    def load_rels(self, path, data):
        """Load rels file from the path. WHen target is
           word/document.xml, then rels path is
           word/_rels/document.xml.rels"""
        if path in self.path2rels:
            raise ValueError("Already have rels for path {}".format(path))
        self.path2rels[path] = DocxRels(path, data)

    def get_rels(self, target, create_new=True):
        """Returns relationships for the given target, or None if not found"""
        #first construct the rels path
        splitted = target.split("/")
        fname = splitted.pop()
        splitted.append("_rels")
        splitted.append(fname + ".rels")
        relspath = "/".join(splitted)

        rels = self.path2rels.get(relspath)
        if rels is None:
            if create_new:
                #create new rels
                rels = DocxRels.create_empty(relspath)
                self.path2rels[relspath] = rels
        return rels

class DocxTemplate(object):
    def __init__(self, docx_file):
        #Open the docx as a ZIP archive and read content of the
        self.docx = docx_file
        with zipfile.ZipFile(self.docx, 'r') as z:
            self._read_content(z)
            self._read_rels(z)

    def _read_rels(self, z):
        self.rels_catalog = DocxRelsCollection()
        for info in z.infolist():
            ctype = self.content_types.get(info.filename)
            if ctype == 'application/vnd.openxmlformats-package.relationships+xml':
                #read the rels file
                rels_data = z.read(info.filename)
                self.rels_catalog.load_rels(info.filename, rels_data)

    def _read_content(self, z):
        # Open the DOCX file as a ZIP archive
        # Extract the [Content_Types].xml file from the DOCX file
        content_types_xml = z.read('[Content_Types].xml')
        self.content_types = DocxContentTypes(content_types_xml)


    def generate(self, data, output):
        #copy all files from the source archive to the output, except the main document
        docx = DocxDocument(self)
        docx.generate(data, output)

class DocxDocument(object):
    """Instance of the generated document. Stores state that changes during generation"""
    def __init__(self, template): #type: (DocxTemplate)->None
        self.template = template
        self.media = []
        self.temporary_media_dir = None
        self.rels_catalog = deepcopy(template.rels_catalog) #we will modify them, so make a copy
        self._current_item = None
        self._next_image_index = 1
        self.content_types = deepcopy(template.content_types)


    def _parse_template(self, document):
        etree = ET.fromstring(document)
        connect_template_ranges(etree)
        #parse the template from the main part
        return parse(etree, expansion_markers=EXPAND_TAGS)


    def generate(self, data, output):
        ctype_to_action = {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml": "template",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml": "template",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml": "template",
            'application/vnd.openxmlformats-package.relationships+xml': "skip"
        }
        with zipfile.ZipFile(output, 'w') as zout:
            with zipfile.ZipFile(self.template.docx, 'r') as zin:
                for item in zin.infolist():
                    fdata = zin.read(item.filename)
                    #first get defailt ctype for
                    item_ctype = self.content_types.get(item.filename)
                    action = ctype_to_action.get(item_ctype, "copy")
                    if item.filename.lower() == "[content_types].xml":
                        action = "skip" #we will write it later
                    if action == "skip":
                        pass
                    elif action == "template":
                        self._current_item = item.filename
                        zout.writestr(item.filename, self._generate_item(fdata, data))
                        self._current_item = None
                    elif action == "copy":
                        zout.writestr(item.filename, fdata)
                    else:
                        raise RuntimeError("Bad action for item {}: {}".format(item.filename, action))
            self.write_content_types(zout)
            self.write_media(zout)
            self.write_rels(zout)
        self.cleanup()

    def write_content_types(self, z): #type: (zipfile.ZipFile)->None
        #write the content types file to the ZIP archive
        data = self.content_types.to_xml()
        z.writestr('[Content_Types].xml', data)

    def write_rels(self, z): #type: (zipfile.ZipFile)->None
        for rel in self.rels_catalog:
            #write the rels file to the ZIP archive
            data = rel.tostr()
            if not data:
                continue
            z.writestr(rel.name, data)

    def write_media(self, z):#type: (zipfile.ZipFile)->None
        for media_file, media_name in self.media:
            #write the media file to the ZIP archive
            z.write(media_file, media_name)

    def cleanup(self):
        if self.temporary_media_dir is not None:
            shutil.rmtree(self.temporary_media_dir, ignore_errors=True)

    def _generate_item(self, xmldata, data): #type: (str, Dict[str, Any])->bytes
        template = self._parse_template(xmldata)
        filled = fill(template, data, context=self)
        try:
            del filled.attrib["{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"]
        except KeyError:
            pass
        split_runs_by_newlines(filled)
        self._postprocess_tree(filled)
        return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(filled)

    def _postprocess_tree(self, tree):
        #self._replace_image_size_formulas(tree) #disabled due to a bug: incorrectly detected page size
        pass

    #disabled
    def _replace_image_size_formulas(self, tree):
        """We search for tags that have attributes cx and cy, whose values start with "=", and evaluate them.
        Evaluating is only possible when inner width and height of the page are known.
        """
        def make_functions(page_inner_width, page_inner_height):
            #creates 2 functions: =fit_width(w,h,scale_up,scale_down) and =fit_height(w,h,scale_up,scale_down)
            #returrning optimal width and height
            def fit_box(w, h, scale_up=False, scale_down=True):
                #"""Returns the optimal size of the box, fitting into the page"""
                if w < page_inner_width and h < page_inner_height:
                    #both smaller, scaling up if needed
                    if scale_up:
                        ratio = min(float(page_inner_width) / w, float(page_inner_height) / h)
                        w = int(w*ratio)
                        h = int(h*ratio)
                elif w > page_inner_width or h > page_inner_height:
                    #at least one is bigger
                    if scale_down:
                        ratio = min(float(page_inner_width) / w, float(page_inner_height) / h)
                        w = int(w*ratio)
                        h = int(h*ratio)
                return w, h

            def fit_width(*args,**kwargs):
                return fit_box(*args,**kwargs)[0]
            def fit_height(*args, **kwargs):
                return fit_box(*args, **kwargs)[1]
            return {"fit_width": fit_width, "fit_height": fit_height}
        #Now we should walk the document tree, keeping track of the page size in every part
        # Default A4-ish content area: ~6.5 x 8.5 inches → 6.5*914400, 8.5*914400 in EMUs
        EMUS_PER_TWIP = 635  # 1 twip = 1/20 point = 1/1440 inch → 914400/1440=635 EMUs per twip

        to_process = [] #list of triples (node, attribute name, formula) to update
        def process_nodes(inner_width_twips, inner_height_twips):
            inner_width_emus = inner_width_twips * EMUS_PER_TWIP
            inner_height_emus = inner_height_twips * EMUS_PER_TWIP
            functions = make_functions(inner_width_emus, inner_height_emus)
            # now updating nodes to process
            for node, attr, formula in to_process:
                # we have a formula, so we need to evaluate it
                float_value = eval(formula, {}, functions)
                value = str(int(float_value))  # convert to emus
                node.attrib[attr] = value
            # clear the to_process list
            del to_process[:]

        def walk_tree(node):
            # Detect section properties → new page size context
            if node.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr":
                page_width = page_height = None
                margin_left = margin_right = margin_top = margin_bottom = 0
                for child in node:
                    if child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgSz":
                        page_width = int(child.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w", "0"))
                        page_height = int(child.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}h", "0"))
                    if child.tag == "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgMar":
                        margin_left = int(child.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left", "0"))
                        margin_right = int(child.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right", "0"))
                        margin_top = int(child.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top", "0"))
                        margin_bottom = int(child.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom", "0"))
                if page_width is not None and page_height is not None:
                    inner_width_twips = page_width - margin_left - margin_right
                    inner_height_twips = page_height - margin_top - margin_bottom
                    process_nodes(inner_width_twips, inner_height_twips)

            elif node.tag == "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent" or \
                node.tag == "{http://schemas.openxmlformats.org/drawingml/2006/main}ext":
                #check if we have cx and cy attributes
                for attr in ("cx", "cy"):
                    if attr in node.attrib:
                        value = node.attrib[attr]
                        if value.startswith("="):
                            formula = value[1:] #remove the "="
                            to_process.append((node, attr, formula))
            for child in node:
                walk_tree(child)
        walk_tree(tree)
        if to_process:
            process_nodes(9026, 13958)  # I found these values in a real document, they are close to A4 size

    def get_media_dir(self):
        if self.temporary_media_dir is not None:
            return self.temporary_media_dir
        import tempfile
        #create temporary media directory
        self.temporary_media_dir = tempfile.mkdtemp(prefix="docxmedia-")
        #also create media subdirectory
        return self.temporary_media_dir

    def image(self, image_file, dpi=96, scale_down=True, scale_up=True, size=None, fit_size=None): #type: (str)->XMLFragment
        """Adds image reference to the media list, and returns XML code that uses it
        scale_down: hen size is not specified, allows to scale image down to fit box size
        scale_up: when size is not specified, allows to scale image up to fit box size
        size: if specified, is a tuple (width, height) in pixels, of exact image size. If height is None or -1, it is detected automatically.
        fit_size: if specified, is a tuple (width, height) in pixels, of the size to fit the image into. Hen fitting, scale_up and scale_down are considered.

        """
        if size is not None and fit_size is not None:
            raise ValueError("Cannot specify both size and fit_size parameters")

        from PIL import Image
        #first get rels for the current item
        im = Image.open(image_file)
        px_width, px_height = im.size

        #convert pixels to EMUs
        emu_per_pixel = 914400.0 / dpi
        emu_width = int(px_width * emu_per_pixel)
        emu_height = int(px_height * emu_per_pixel)


        mediadir = self.get_media_dir()
        ext = os.path.splitext(image_file)[1].lower()
        if not ext: raise ValueError("Image name without extension: {}".format(image_file))
        def get_image_temp_path(ext):
            imagename = "genimage{}{}".format(self._next_image_index, ext)
            self._next_image_index += 1
            return imagename, os.path.join(mediadir, imagename)
        need_copy = True
        if ext in (".jpg", ".jpeg"):
            self.content_types.add_extension("jpeg", "image/jpeg")
            ext = "jpeg"
        elif ext == ".png":
            self.content_types.add_extension("png", "image/png")
        else:
            #convert to PNG using PIL
            self.content_types.add_extension("png", "image/png")
            ext = ".png"
            imagename, image_temp_path = get_image_temp_path(".png")
            im.save(image_temp_path)
            need_copy = False #image is already in the right place

        del im

        if need_copy:
            imagename, image_temp_path = get_image_temp_path(ext)
            #copy the image to the media directory with the new name
            shutil.copy(image_file, image_temp_path)

        path_splitted = self._current_item.split("/")
        path_splitted.pop()
        path_splitted.append("media")
        path_splitted.append(imagename)
        full_image_path = "/".join(path_splitted)
        #create a new relationship for the image
        rels = self.rels_catalog.get_rels(self._current_item, create_new=True)
        relid = rels.new_relationship("media/"+imagename,
                                       "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        self.media.append((image_temp_path, full_image_path))
        #return the XML fragment for the image
        pic_id = randint(1000000, 9999999)
        pic_name = "Image {}".format(pic_id)
        #RIght now we can't write them, because it is too hard to know the size of the page to fit the image
        if size is None and fit_size is None:
            #No fitting, use size as is
            formula_width = str(emu_width)
            formula_height = str(emu_height)
        elif size is not None:
            #use the given size
            w = size[0]
            formula_width = str(int(w * emu_per_pixel))
            h = size[1]
            if h is None or h==-1 or h==0:
                h = w * px_height / px_width
            formula_height = str(int(h * emu_per_pixel))
        elif fit_size is not None:
            fit_w, fit_h = fit_size
            #fit the image to the given size
            scale_w = float(fit_w) / px_width
            scale_h = float(fit_h) / px_height
            scale = min(scale_w, scale_h)
            if px_width < fit_w and px_height < fit_h and scale_up:
                #scale up the image
                w = int(px_width * scale)
                h = int(px_height * scale)
            elif (px_width > fit_w or px_height > fit_h) and scale_down:
                #scale down the image
                w = int(px_width * scale)
                h = int(px_height * scale)
            else:
                w, h = px_width, px_height
            formula_width = str(int(w * emu_per_pixel))
            formula_height = str(int(h * emu_per_pixel))
        else:
            raise RuntimeError("Unreachable code reached")

        return XMLFragment.from_text(""""<ns0:drawing>
          <ns2:inline distB="0" distL="0" distR="0" distT="0" ns3:anchorId="1828AD8F" ns3:editId="64F02F6F">
            <ns2:extent cx="{emu_width}" cy="{emu_height}"/>
            <ns2:docPr id="{pic_id}" name="{pic_name}"/>
            <ns2:cNvGraphicFramePr>
              <ns4:graphicFrameLocks noChangeAspect="1"/>
            </ns2:cNvGraphicFramePr>
            <ns4:graphic>
              <ns4:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <ns5:pic>
                  <ns5:nvPicPr>
                    <ns5:cNvPr id="{pic_id}" name="{pic_name}"/>
                    <ns5:cNvPicPr/>
                  </ns5:nvPicPr>
                  <ns5:blipFill>
                    <ns4:blip ns6:embed="{relid}">
                      <ns4:extLst>
                        <ns4:ext uri="{{28A0092B-C50C-407E-A947-70E740481C1C}}">
                          <ns7:useLocalDpi val="0"/>
                        </ns4:ext>
                      </ns4:extLst>
                    </ns4:blip>
                    <ns4:stretch>
                      <ns4:fillRect/>
                    </ns4:stretch>
                  </ns5:blipFill>
                  <ns5:spPr>
                    <ns4:xfrm>
                      <ns4:off x="0" y="0"/>
                       <ns4:ext cx="{emu_width}" cy="{emu_height}"/>
                    </ns4:xfrm>
                    <ns4:prstGeom prst="rect">
                      <ns4:avLst/>
                    </ns4:prstGeom>
                  </ns5:spPr>
                </ns5:pic>
              </ns4:graphicData>
            </ns4:graphic>
          </ns2:inline>
        </ns0:drawing>""".format(relid=relid, pic_id=pic_id, pic_name=pic_name, emu_width=formula_width, emu_height=formula_height),
            #xmlns:ns0="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            #xmlns:ns1="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:ns2="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:ns3="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:ns4="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:ns5="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:ns6="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:ns7="http://schemas.microsoft.com/office/drawing/2010/main"
            nsmap={"ns0": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                   "ns1": "http://schemas.microsoft.com/office/word/2010/wordml",
                   "ns2": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
                   "ns3": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
                   "ns4": "http://schemas.openxmlformats.org/drawingml/2006/main",
                   "ns5": "http://schemas.openxmlformats.org/drawingml/2006/picture",
                   "ns6": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                   "ns7": "http://schemas.microsoft.com/office/drawing/2010/main"},
                                     parent_tag = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r",)



class DocxImageInserter(XMLRunner):
    def __init__(self, image_path, **kwargs):
        """FOr arguments, see DocxDocument.image() method"""
        self.image_path = image_path
        self.kwargs = kwargs

    def run(self, context):
        assert isinstance(context, DocxDocument)
        return context.image(self.image_path, **self.kwargs)

def split_runs_by_newlines(tree): #type: (ET.Element)->None
    """Find text runs that contain newlines and split them, inserting w:br tags in between"""
    #find all text runs
    for run in tree.iterfind(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r"):
        #get text
        text_tag = run.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
        if text_tag is None:
            continue
        text = text_tag.text
        if text is None:
            continue
        parts = text.split("\n")
        if len(parts) > 1:
            text_tag.text = parts[0]
            for part in parts[1:]:
                ET.SubElement(run, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br")
                #create text node
                textnode = ET.SubElement(run, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t",
                                         attrib={"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
                textnode.text = part

