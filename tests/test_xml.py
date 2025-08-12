import unittest
import xml.etree.ElementTree as ET
from pymsword.xml_template import parse, XMLFragment, fill


class TestXMLTemplate(unittest.TestCase):
    def setUp(self):
        self.sample_xml = """
<root>
<document>
<paragraph>
<text>Some {{placeholder}} text</text>
<text>Some {% group %} {{placeholder2}} {% end group %}</text>
<text>A group that contains tags within: {% group1 %} <tag> {{placeholder3}} </tag> {% end group1 %}</text>
<text>A group that overlaps with the paragraph boundary: {% group2 %} <tag> {{placeholder4}} </tag> </text>
</paragraph>
<paragraph>
<text> continue text {% end group2 %}</text>
</paragraph>
</document>
<a> <b> <c>{% egrp marker_a%} {% end egrp %}</c> </b> </a>
</root>
"""
        self.data = {
            'placeholder': 'placeholder value',
            'group': [
                {'placeholder2': 'placeholder2 value 1'},
                {'placeholder2': XMLFragment.from_text("<aa><bb>xxx</bb></aa>")},
            ],
            'group1': [
                {'placeholder3': 'placeholder3 value 1'},
                {'placeholder3': 'placeholder3 value 2'},
            ],
            'group2': [
                {'placeholder4': 'placeholder4 value 1'},
                {'placeholder4': 'placeholder4 value 2'},
            ],
            'egrp': [{}, {}]
        }

    def test_exact_xml_structure(self):
        root = ET.fromstring(self.sample_xml)
        template = parse(root, expansion_markers={'marker_a': {'b'}})
        filled_tree = fill(template, self.data)

        # Build the expected XML tree
        expected_xml = """
<root>
<document>
<paragraph>
<text>Some placeholder value text</text>
<text>Some  placeholder2 value 1  <aa><bb>xxx</bb></aa> </text>
<text>A group that contains tags within:  <tag> placeholder3 value 1 </tag>  <tag> placeholder3 value 2 </tag> </text>
<text>A group that overlaps with the paragraph boundary:  <tag> placeholder4 value 1 </tag> </text>
</paragraph>
<paragraph>
<text> continue text </text></paragraph><paragraph><text> <tag> placeholder4 value 2 </tag> </text>
</paragraph>
<paragraph>
<text> continue text </text>
</paragraph>
</document>
<a> <b> <c> </c> </b><b> <c> </c> </b> </a>
</root>
"""
        expected_tree = ET.fromstring(expected_xml)

        # Compare trees structurally
        def compare_elements(e1, e2):
            self.assertEqual(e1.tag, e2.tag, f"Tag mismatch: {e1.tag} != {e2.tag}")
            self.assertEqual(
                (e1.text or "").strip(),
                (e2.text or "").strip(),
                f"Text mismatch in <{e1.tag}>: {e1.text!r} != {e2.text!r}"
            )
            self.assertEqual(len(e1), len(e2), f"Child count mismatch in <{e1.tag}>")
            for c1, c2 in zip(e1, e2):
                compare_elements(c1, c2)

        print(ET.tostring(filled_tree, encoding='unicode'))
        compare_elements(filled_tree, expected_tree)

if __name__ == "__main__":
    unittest.main()
