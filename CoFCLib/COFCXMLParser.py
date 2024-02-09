import xml.etree.ElementTree as ET
import xml.dom.minidom
import os
import logging
import json

class COFCXMLParser:
    def __init__(self, input_xml, output_xml=None, json_input=None, log_path=None):
        self.input_xml: str = os.path.normpath(input_xml)
        self.output_xml: str = os.path.normpath(output_xml) or self.generate_output_filename()
        self.json_input: str = os.path.normpath(json_input)
        self.log_path = os.path.normpath(log_path) or ''
        self.log_file: str = os.path.join(self.log_path, os.path.splitext(os.path.basename(self.input_xml))[0] + '.log')
        pass


    #def element_to_dict(self,elem):
    #    result = {}
    #    if elem.attrib:
    #        result["@attributes"] = elem.attrib
    #    if elem.text:
    #        result["text"] = elem.text.strip()
    #    for child in elem:
    #        cdata = self.element_to_dict(child)
    #        if child.tag in result:
    #            if isinstance(result[child.tag], list):
    #                result[child.tag].append(cdata)
    #            else:
    #                result[child.tag] = [result[child.tag], cdata]
    #        else:
    #            result[child.tag] = cdata
    #    return result

    #def dict_to_element(self, tag, data):

    #    if isinstance(data, list):
    #        elem = ET.Element(tag)
    #        for item in data:
    #            child = self.dict_to_element(tag, item)
    #            elem.append(child)
    #            return elem

    #    elem = ET.Element(tag)
    #    if "text" in data and data["text"] is not None:
    #        elem.text = str(data["text"])
    #    if "@attributes" in data and isinstance(data["@attributes"], dict):
    #        for k, v in data["@attributes"].items():
    #            elem.set(k,v)

    #    for k,v in data.items():
    #        if k not in ["@attributes", "text"]:
    #            child = self.dict_to_element(k,v)
    #            elem.append(child)
    #    return elem


    @staticmethod
    def build_hierarchy_from_xpaths(xpath_list):
        hierarchy = {}
        for xpath in xpath_list:
            nodes = xpath.split('/')
            nodes = nodes[2:]
            current_level = hierarchy
            for node in nodes:
                if node not in current_level:
                    current_level[node] = {}
                current_level = current_level[node]
        return hierarchy

    @staticmethod
    def reduce_and_create_new_xml(root, new_root_name, selected_xpath, xpath_hierarchy):
        new_tree = ET.ElementTree(ET.Element(new_root_name))
        new_root = new_tree.getroot()

        processed_root_node = set()
        for xpath in selected_xpath:
            test_path = root.findall(xpath)
            if test_path is None:
                print("invalid_path")
                continue

            xpath = xpath.replace(".//", "", 1)
            xpath = xpath.replace("//", "", 1)
            xpath_parts = xpath.split('/')
            first_elem = xpath_parts[0] if len(xpath_parts) >= 1 else None
            if first_elem == "." or first_elem == "":
                continue
            print(first_elem)
            if first_elem is not None:
                selected_nodes = root.findall(".//" + first_elem)
                if selected_nodes is None:
                    continue
                if first_elem in processed_root_node:
                    continue
                print("Processing",first_elem)
                for snode in selected_nodes:
                    child_nodes = snode.findall('*')
                    if child_nodes is None:
                        continue
                    selected_xpath_child_nodes = xpath_hierarchy[first_elem].keys()
                    for child in child_nodes:
                        ctag_name = child.tag
                        if ctag_name not in selected_xpath_child_nodes:
                            #print('\tremoving=',first_elem,ctag_name)
                            if snode.find(ctag_name) is not None:
                                snode.remove(snode.find(ctag_name))
                        else:
                            child_of_child_nodes = child.findall('*')
                            if child_of_child_nodes is None:
                                continue
                            selected_coc_xpath_tags = xpath_hierarchy[first_elem][ctag_name].keys()
                            #print(selected_coc_xpath_tags)
                            for coc_node in child_of_child_nodes:
                                coc_tag_name = coc_node.tag
                                #print(coc_tag_name)
                                if coc_tag_name not in selected_coc_xpath_tags:
                                    #print(coc_tag_name,' not in ',selected_coc_xpath_tags)
                                    #print('\t\tremoving=',coc_tag_name)
                                    child.remove(coc_node)
                        pass
                    new_root.append(snode)
                if first_elem not in processed_root_node:
                    processed_root_node.add(first_elem)
        return new_tree

    def generate_output_filename(self):
        base_name = os.path.splitext(os.path.basename(self.input_xml))[0]
        return f"{base_name}_reduced"

    def convert_xml(self):
        # Configure logging
        logging.basicConfig(filename=self.log_file, level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')

        logging.info(f"------------------------------------------------------------")
        logging.info(f"Input XML: {self.input_xml}")

        selected_xpaths = []
        # Parse the JSON input to get selected xpaths
        with open(self.json_input, 'r') as json_file:
            json_data = json.load(json_file)
            for item in json_data:
                if item['Selected'] == 1:
                    selected_xpaths.append(item['AttributeNameinXML'])

        xpath_hierarchy = self.build_hierarchy_from_xpaths(selected_xpaths)
        # Parse the JSON input
        tree = ET.parse(self.input_xml)
        root = tree.getroot()
        new_tree = self.reduce_and_create_new_xml(root, "CDMaskResults", selected_xpaths, xpath_hierarchy)
        new_root = new_tree.getroot()
        xml_string = ET.tostring(new_root, xml_declaration=False, method='xml').decode('utf-8').strip().replace('\n','')
        pretty_xml = xml.dom.minidom.parseString(xml_string).toprettyxml(newl='\n', indent='    ')
        with open(self.output_xml, "w") as file:
            file.write(pretty_xml)
        logging.info(f"Output XML: {self.output_xml}")
        print(f"XML Generated: {self.output_xml}")
        return self.output_xml


test_cofc_xml_parser_test_2 = 0
if test_cofc_xml_parser_test_2 or __name__ == "__main__":
    input_file = r'C:\Users\magalang\Documents\Reference\DNP\CofC_20200501121042.xml'
    output_file = 'xml_to_cd_xml_output.xml'
    config_file = r'C:\Users\magalang\Documents\demo\cofc_transformer\config\OS-006768-01_24.xpaths.json'
    log_path = './logs/'
    if not os.path.exists('./logs/'):
        os.makedirs('./logs/')

    xml_reducer = COFCXMLParser(input_file, output_file, config_file, log_path)
    xml_reducer.convert_xml()

