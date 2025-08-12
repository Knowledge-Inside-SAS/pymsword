import re
from ast import literal_eval
from copy import deepcopy
import xml.etree.ElementTree as ET
from collections import defaultdict
from itertools import groupby
from typing import Optional, List, Dict, Iterable, Any, Set, Callable, Union


class TemplateElementBase(object):
    def __init__(self, parent): #type: (Optional[TemplateContainerBase])->None
        self.parent = parent
        self.valid = True

    def index(self):
        if self.parent is None: raise ValueError("Root node has no index")
        return self.parent.children.index(self)

    def pretty_print(self, indent=""):
        print(indent, str(self))

    def walk(self): #type: ()->Iterable[TemplateElementBase]
        stack = [self]
        while stack:
            node = stack.pop()
            yield node
            if isinstance(node, TemplateContainerBase):
                stack.extend(reversed(node.children))

    def is_static_data(self):
        """Returns True if this element is static data, i.e. it does not contain any placeholders or groups"""
        return all(isinstance(child, (TemplateText, TemplateNode)) for child in self.walk())

class TemplateContainerBase(TemplateElementBase):
    def __init__(self, parent):
        TemplateElementBase.__init__(self, parent)
        self.children = [] #type: List[TemplateElementBase]
    def pretty_print(self, indent=""):
        print(indent, str(self))
        for child in self.children:
            child.pretty_print(indent + "  ")
    def set_children(self, children):
        self.children = children
        for child in children:
            child.parent = self
    def extend_children(self, children):
        for child in children:
            self.children.append(child)
            child.parent = self

class TemplateNode(TemplateContainerBase):
    def __init__(self, xml_node, parent): #type: (ET.Element, Optional[TemplateContainerBase])->None
        TemplateContainerBase.__init__(self, parent)
        self.xml_node = xml_node
        self.sticky_head = False
        self.sticky_tail = False

    def split_by(self, index):
        """Splits the node into two nodes at the given index.
        The first node contains children[:index], the second node contains children[index+1:]
        (this child at index is excluded)
        Split parts are marked as "sticky" so that they can be merged later.
        """
        if self.parent is None: raise ValueError("Cannot split root node")
        head = TemplateNode(self.xml_node, self.parent)
        head.set_children(self.children[:index])
        head.sticky_tail = True
        tail = TemplateNode(self.xml_node, self.parent)
        tail.set_children(self.children[index+1:])
        tail.sticky_head = True
        central_child = self.children[index]
        #reconnect head, center, tail to the parent instead of self
        self_index = self.index()

        new_siblings = self.parent.children[:self_index]# + self.parent.children[self_index+1:]
        new_siblings.extend([head, central_child, tail])
        new_siblings.extend(self.parent.children[self_index+1:])
        self.parent.children = new_siblings
        central_child.parent = self.parent

        self.valid = False

    def __str__(self):
        s = "Node: "+self.xml_node.tag
        if self.sticky_head:
            s += " (sticky head)"
        if self.sticky_tail:
            s += " (sticky tail)"
        if not self.valid:
            s += " (!invalid!)"
        return s

class TemplateNamedElement(TemplateElementBase):
    def __init__(self, name, parent):
        TemplateElementBase.__init__(self, parent)
        self.name = name

    def __str__(self):
        return "{}({})".format(type(self).__name__, self.name)
class TemplatePlaceholder(TemplateNamedElement):
    pass

class TemplateGroupBase(TemplateNamedElement):
    def __init__(self, name, parent):
        TemplateNamedElement.__init__(self, name, parent)
        self.opposite = None #type: Optional[TemplateGroupBase]
    def __str__(self):
        return "{}({}) matched to:{}".format(type(self).__name__, self.name, self.opposite.name)
    def float_up(self):
        self.parent.split_by(self.index())

class CollectedGroupNode(TemplateContainerBase):
    def __init__(self, name, parent):
        TemplateContainerBase.__init__(self, parent)
        self.name = name
    def __str__(self):
        return "Group({})".format(self.name)

class TemplateGroupOpen(TemplateGroupBase):
    def __init__(self, name, parent, expand_to):
        TemplateGroupBase.__init__(self, name, parent)
        self.expand_to = expand_to

class TemplateGroupClose(TemplateGroupBase):
    pass

class TemplateText(TemplateElementBase):
    def __init__(self, text, parent): #type: (str, Optional[TemplateContainerBase])->None
        TemplateElementBase.__init__(self, parent)
        self.text = text
    def __str__(self):
        return "Text: "+repr(self.text)

class XMLFragment(object):
    def __init__(self, data, parent_tag=None):#type: (List[ET.Element], Optional[str])->None
        """Represents a fragment of XML, which can be put instead of a string into the data
        """
        self.data = data
        self.parent_tag = parent_tag  # type: Optional[str]

    @staticmethod
    def from_text(s, nsmap=None, parent_tag=None):  # type: (str, Optonal[Dict[str, str]], Optional[str]) -> XMLFragment
        """Creates an XMLFragment from a string, with optional namespace map."""
        # Build xmlns attributes from nsmap
        ns_attrs = ""
        if nsmap:
            ns_attrs = " " + " ".join(
                'xmlns:{}="{}"'.format(prefix, uri) for prefix, uri in nsmap.items()
            )
        # Wrap the string with root including namespace declarations
        wrapped = "<root{}>{}</root>".format(ns_attrs, s)
        root = ET.fromstring(wrapped)
        #print wrapped
        frag = XMLFragment([child for child in root if isinstance(child, ET.Element)], parent_tag=parent_tag)
        frag.nsmap = nsmap or {}
        return frag
    def __str__(self):
        return "XMLFragment: [{}]".format("".join(ET.tostring(e) for e in self.data))
    def __len__(self):
        return len(self.data)
    def __getitem__(self, index):
        """Returns the i-th element of the fragment"""
        return self.data[index]
    def __iter__(self):
        """Iterates over the elements of the fragment"""
        return iter(self.data)

class TemplateXMLFragment(TemplateElementBase):
    def __init__(self, data, parent):
        #type: (XMLFragment, Optional[TemplateContainerBase])->None
        TemplateElementBase.__init__(self, parent)
        self.data = data
    def __str__(self):
        return  "TemplateXMLFragment: [{}]".format(self.data)

RE_INSTRUCTION = re.compile(r"\{%(.*?)%}|\{\{(.*?)}}")

class XMLRunner(object):
    def __init__(self, callable_): #type: (Optional[Callable[[Any], Union[str, XMLFragment]]])->None
        assert callable(callable_)
        self._callable = callable_

    def __call__(self, context):
        """Runs the XML runner with the given context.
        Should return an XMLFragment or a string.
        """
        return self.run(context)

    def run(self, context):
        return self._callable(context) #type: XMLFragment

def fill_template(template, data, context=None):#type: (TemplateNode, Dict[str, Any], Any)->TemplateNode
    """Fills the template with the given data.
    Data dict must contain string values for placeholders,
    and list-of-dict values for groups.
    context - argument, passed to runners, if they are present.
    """
    def fill(node, filled_parent, data):#type: (TemplateElementBase, Optional[TemplateContainerBase], Dict[str, Any])->List[TemplateElementBase]
        if isinstance(node, TemplateText):
            return [TemplateText(node.text, filled_parent)]
        elif isinstance(node, TemplateXMLFragment):
            raise ValueError("TemplateXMLFragment can only be present in filled data, template is wrong")
        elif isinstance(node, TemplatePlaceholder):
            value = data.get(node.name, '')
            if isinstance(value, XMLRunner):
                #Value is dynamically generated
                value = value(context)
            if isinstance(value, XMLFragment):
                #if value is XMLFragment, then return it as a TemplateXMLFragment
                return [TemplateXMLFragment(value, filled_parent)]
            else:
                return [TemplateText(str(value), filled_parent)]
        elif isinstance(node, (TemplateGroupOpen, TemplateGroupClose)):
            raise ValueError("Cannot fill uncollected groups!")
        elif isinstance(node, CollectedGroupNode):
            #fill the group by repeating its children and exploding its placeholders
            group_data = data.get(node.name)
            if group_data is None: return []#no data for this group
            if not isinstance(group_data, (list, tuple)): raise ValueError("Data for the group {} should be a list or tuple, not {}".format(repr(node.name), type(group_data)))
            generated = []
            for item_data in group_data:
                updated_data = data.copy()
                updated_data.update(item_data)
                for child in node.children:
                    generated.extend(fill(child, filled_parent, updated_data))
            return generated
        elif isinstance(node, TemplateNode):
            #fill the node by filling its children
            filled = TemplateNode(node.xml_node, filled_parent)
            filled.sticky_head = node.sticky_head
            filled.sticky_tail = node.sticky_tail
            generated = []
            for child in node.children:
                generated.extend(fill(child, filled_parent, data))
            filled.set_children(generated)
            return [filled]
        else:
            raise ValueError("Unsupported node type: {}".format(type(node)))

    if isinstance(template, TemplateNode):
        filled_nodes = fill(template, None, data)
        assert len(filled_nodes) == 1 and isinstance(filled_nodes[0], TemplateNode)
        return filled_nodes[0]
    else:
        assert template.is_static_data()
        return template


def parse_template_xml(root, parent): #type: (ET.Element, Optional[TemplateContainerBase])->TemplateNode
    #Each XML node translates to a TemplateNode
    #Each text piece node translates to 1 or more TemplateText and placeholders

    node = TemplateNode(root, parent)
    #First, parse text of the root
    if root.text:
        node.children.extend(parse_text_template(root.text, node))

    #Then, parse children
    for child in root:
        node.children.append(parse_template_xml(child, node))
        if child.tail:
            node.children.extend(parse_text_template(child.tail, node))
    return node

def parse_text_template(text, parent):#type: (str, TemplateContainerBase)->List[TemplateElementBase]
    #parse the text into TemplateText and placeholders
    #find all instructions and split text by them.
    #Each instruction is either a placeholder or a group
    #Each group has a name and a list of instructions

    #find all instructions
    # text between them is transformed to TemplateText
    # instructions are transformed to TemplatePlaceholder or TemplateGroup
    parsed = []
    last_end = 0
    for match in re.finditer(RE_INSTRUCTION, text):
        # Process the string fragment before the match
        fragment = text[last_end:match.start()]
        if fragment:
            parsed.append(TemplateText(fragment, parent))
        # Process the match
        # it is either {% ... %} or {{ ... }}
        if match.group(1):
            instruction = match.group(1).strip()
            if instruction.startswith("end "):
                parsed.append(TemplateGroupClose(instruction[4:], parent))
            else:
                #It is either {% group %} or {% group expand_marker %}
                #try to parse the marker
                instruction_parts = instruction.rsplit(" ", 1)
                expand_to = None
                if len(instruction_parts)==2:
                    expand_to = instruction_parts[1]
                    instruction = instruction_parts[0].strip()
                parsed.append(TemplateGroupOpen(instruction, parent, expand_to))
        else:
            instruction = match.group(2).strip()
            #if instruction is not starting with '"' or "'", then it is a placeholder
            if not instruction.startswith('"') and not instruction.startswith("'"):
                parsed.append(TemplatePlaceholder(instruction, parent))
            else:
                #it is a text
                #evaluate the string literal
                svalue = literal_eval(instruction)
                parsed.append(TemplateText(svalue, parent))
        last_end = match.end()
    # Process the string fragment after the last match
    fragment = text[last_end:]
    if fragment:
        parsed.append(TemplateText(fragment, parent))
    return parsed

#Now find matching open and close group elements
def find_matching_groups(root): #type: (TemplateNode)->None
    name2group = defaultdict(list) #type: Dict[str, List[TemplateGroupOpen]]
    for child in root.walk():
        if isinstance(child, TemplateGroupOpen):
            name2group[child.name].append(child)
        elif isinstance(child, TemplateGroupClose):
            open_groups = name2group.get(child.name)
            if not open_groups:
                raise ValueError("Group {} is closed but not opened".format(child.name))
            open_group = open_groups.pop()
            if not open_groups:
                del name2group[child.name]
            open_group.opposite = child
            child.opposite = open_group
    if name2group:
        raise ValueError("Some groups are not closed: {}".format(", ".join(name2group.keys())))


def expand_group_markers(template, marker2tags): #type: (TemplateNode, Dict[str, Set[str]])->None
    """
    If open group's expand_to is not None, then
    move group open and close markers up by the XML
    hierarchy, until their parent is in the marker2tags[expand_to]

    Needed because in some cases, we can't put textual markers where we want.

    Example:
        <a> <b> <c> {% group exp_a%} {% end group%}</c> </b> </a>
        marker2tags = {"exp_a": {"b"}}

        then after expansion we get:
        <a> {% group exp_a%}<b> <c> </c> </b>{% end group%} </a>
    """
    expansible_nodes = [child for child in template.walk()
                        if isinstance(child, TemplateGroupOpen) \
                            and child.expand_to is not None and \
                            child.expand_to in marker2tags]
    if not expansible_nodes: return

    def pop_tag_up(node_open, wait_tags, direction): #type: (TemplateElementBase, Set[str], int)->None
        #move the open and close markers up
        while node_open.parent is not None:
            last_pop = node_open.parent.xml_node.tag in wait_tags
            parent_node = node_open.parent
            grandparent_node = parent_node.parent
            if grandparent_node is None:
                raise ValueError("Cannot move node {} up, its parent already at the top level".format(node_open))
            #remove the node from its parent
            parent_node.children.remove(node_open)
            #insert it to the grandparent, right before parent_node
            parent_index = parent_node.index()
            grandparent_node.children.insert(parent_index if direction==-1 else parent_index+1, node_open)
            node_open.parent = grandparent_node
            if last_pop: break

    for node_open in expansible_nodes:
        node_close = node_open.opposite
        assert isinstance(node_close, TemplateGroupClose)
        pop_tag_up(node_open, marker2tags[node_open.expand_to], -1)
        pop_tag_up(node_close, marker2tags[node_open.expand_to], 1)


def parents(node):
    while node.parent is not None:
        yield node.parent
        node = node.parent

def last_common_parents(node1, node2): #type: (TemplateElementBase, TemplateElementBase)->TemplateElementBase
    """Find the last common parent of node1 and node2"""
    parents1 = list(parents(node1))
    parents2 = list(parents(node2))
    common = None
    for p1, p2 in zip(parents1[::-1], parents2[::-1]):
        if p1 is p2:
            common = p1
        else:
            break
    assert common is not None #this should never happen, at least root is common
    return common

def push_group_boundaries_up(template):
    """Sometimes group boundaries are inside different tags.
    FOr example, like this:
    <a><b> xxx {% group %} yyy</b> <c> zzz{% end group %}www </c></a>
    In this case, we need to move the boundaries up, splitting the nodes by the boundary, obtaining the following structure:
    <a><b> xxx</b> {% group %} <b>yyy</b> <c> zzz</c> {% end group %}<c>www </c></a>

    This way, both begin and end of the group are inside the same tag <a/>
    Nodes that were split should be marked as "sticky", so that they can be merged later.
    """

    #Find opening tags of the groups
    groups = [child for child in template.walk()
              if isinstance(child, TemplateGroupOpen)]
    assert all(g.opposite is not None for g in groups)
    for group_open in groups:
        group_close = group_open.opposite
        assert isinstance(group_close, TemplateGroupClose)
        #find the last common parent of the group open and close
        common_parent = last_common_parents(group_open, group_close)
        #float open and close tags up to the common parent
        while group_open.parent is not common_parent:
            group_open.float_up()
        while group_close.parent is not common_parent:
            group_close.float_up()
    assert all(g.parent is g.opposite.parent for g in groups)

def collect_groups(template): #type: (TemplateNode)->None
    """Collects all groups into CollectedGroupNode objects"""
    groups = [child for child in template.walk()
                if isinstance(child, TemplateGroupOpen)]
    # replacing ... [open] ... [close] ... with  ... [collected] ...
    for g in groups:
        parent = g.parent
        if not (g.opposite.parent is parent):
            raise ValueError("Node {} and its opposite {} are not in the same parent: {} and {}".format(g, g.opposite, parent, g.opposite.parent))
        open_index = g.index()
        close_index = g.opposite.index()
        collected_group = CollectedGroupNode(g.name, parent)
        collected_group.set_children(parent.children[open_index+1:close_index])
        #replace all children with
        parent.children[open_index:close_index+1] = [collected_group]
        g.valid = False
        g.opposite.valid = False

def glue_texts(tree):
    if isinstance(tree, TemplateNode):
        new_children = []
        had_glues = False
        for is_text, ielements in groupby(tree.children, lambda x: isinstance(x, TemplateText)):
            elements = list(ielements)
            if is_text:
                if len(elements)==1:
                    new_children.append(elements[0])
                else:
                    new_children.append(TemplateText("".join(e.text for e in elements), tree))
                    had_glues = True
            else:
                new_children.extend(elements)
                for elem in elements:
                    glue_texts(elem)
        if had_glues:
            tree.set_children(new_children)

def glue_sticky_elements(template):
    """TempalteNodes that have matching sticky heads and tails AND refer to the same tag
    can be glued together to a single TemplateNode with joined children
    """
    def walk(root):
        modified = False
        #First glue the deeper elements, then the shallower ones
        for child in root.children:
            if isinstance(child, TemplateNode):
                modified = walk(child) or modified
        #Now glue own children
        if isinstance(root, TemplateNode):
            new_children = []
            had_glues = False
            last_node = None #type: Optional[TemplateNode]
            for child in root.children:
                if not isinstance(child, TemplateNode):
                    last_node = None
                    new_children.append(child)
                else:
                    #maybe this node can be glued with the last node?
                    if last_node and last_node.sticky_tail and child.sticky_head and last_node.xml_node is child.xml_node:
                        #glue them!
                        last_node.extend_children(child.children)
                        last_node.sticky_tail = child.sticky_tail
                        had_glues = True
                    else:
                        new_children.append(child)
                        last_node = child
            if had_glues:
                root.set_children(new_children)
            return modified or had_glues

    while walk(template):
        pass

def generate_xml(filled_template):
    """Generate the XML back from the template.
    Node that element references in the TemplateNodes contain data for the template, not for the filled template."""
    assert isinstance(filled_template, TemplateNode)
    #Create root ElementTree element, copying the parameters from filled_template.tag (which is the ElementTree element of the original template)
    parent_map = {} #type: Dict[ET.Element, ET.Element]

    def walk(template_node, xml_parent): #type: (TemplateElementBase, ET.Element)->None
        if isinstance(template_node, TemplateText):
            #if xml_parent has no children - store data to its text
            if len(xml_parent) == 0:
                xml_parent.text = template_node.text
            else:
                #otherwise, put it to the last child's tail
                xml_parent[-1].tail = template_node.text
        elif isinstance(template_node, TemplateXMLFragment):
            #A list of XML tags, copy them to the xml_parent. Create copies!
            # if parent tag is specified, then go up to that parent,
            # and append the XMLFragment to it
            # then go down and recreate the same structure
            if template_node.data:
                tags_up_to_parent = []
                if template_node.data.parent_tag is not None: #we have to go up
                    while xml_parent.tag != template_node.data.parent_tag:
                        tags_up_to_parent.append((xml_parent.tag, xml_parent.attrib))
                        xml_parent = parent_map[xml_parent]
                        if xml_parent is None:
                            raise ValueError("XML fragment was requested to be put under tag {}, but there are no such parents. Parent hierarchy is:" + "\n".join(
                                "{}".format(tag) for tag, _attrs in tags_up_to_parent[::-1]))
                for xml_element in template_node.data:
                    #create a copy of the XML element, with all its attributes and children
                    xml_parent.append(deepcopy(xml_element))
                #now go down to the parent tag, if needed
                for tag, attrs in tags_up_to_parent[::-1]:
                    xml_parent1 = ET.SubElement(xml_parent, tag, attrs)
                    parent_map[xml_parent1] = xml_parent
                    xml_parent = xml_parent1
            if tags_up_to_parent:
                #parent was changed, so need to return it
                return xml_parent

        elif isinstance(template_node, TemplateNode):
            #create sub-element
            xml_node = ET.SubElement(xml_parent, template_node.xml_node.tag, template_node.xml_node.attrib)
            parent_map[xml_node] = xml_parent
            #fill it with children
            for child in template_node.children:
                new_xml_node = walk(child, xml_node)
                if new_xml_node is not None: #happened when walker inserted a TemplateXMLFragment that required going up to the parent tag
                    xml_node = new_xml_node
        else:
            raise ValueError("Unsupported node type: {}".format(type(template_node)))
    root = ET.Element(filled_template.xml_node.tag,
                      filled_template.xml_node.attrib)
    parent_map[root] = None  # root has no parent
    for child in filled_template.children:
        walk(child, root)
    return root

def parse(template_xml, expansion_markers=None):
    template = parse_template_xml(template_xml, None)
    find_matching_groups(template)
    if expansion_markers is not None:
        expand_group_markers(template, expansion_markers)
    #split the template up by group boundaries
    push_group_boundaries_up(template)
    collect_groups(template)
    return template

def fill(template, data, context=None):
    filled = fill_template(template, data, context=context)
    glue_sticky_elements(filled)
    glue_texts(filled)
    return generate_xml(filled)

