import xml.etree.ElementTree as ET

# Add missing drawings elements from the original file to the additional file
def add_elements_to_xml(original_file, additional_file):
    # Register the desired prefix for the namespace URI
    ET.register_namespace('', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')

    # Parse the original XML file
    original_tree = ET.parse(original_file)
    original_root = original_tree.getroot()

    # Parse the additional XML file
    additional_tree = ET.parse(additional_file)
    additional_root = additional_tree.getroot()

    # Find the xdr:grpSp element with xdr:cNvPr name="Group 25" in the original XML file
    grp_sp = None
    for element in original_root.iter('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}grpSp'):
        nv_grp_sp_pr = element.find('xdr:nvGrpSpPr', {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'})
        c_nv_pr = nv_grp_sp_pr.find('xdr:cNvPr', {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'})
        if c_nv_pr.get('name') == 'Group 25':
            grp_sp = element
            break

    if grp_sp is None:
        print("Unable to find xdr:grpSp with name='Group 25'")
        return

    # Find all xdr:sp and xdr:cxnSp elements with macro="" attribute in the additional XML file
    elements_to_add = additional_root.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}sp[@macro=""]'
                                               ' | .//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}cxnSp[@macro=""]')

    # Add each element (excluding xdr:pic) to the original XML file within the same xdr:grpSp
    for element in elements_to_add:
        if element.tag != '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}pic':
            grp_sp.append(element)

    # Save the modified XML structure
    original_tree.write('drawing1.xml')

    print('Completed!')


# Example usage
add_elements_to_xml('C:\in.xml', 'C:\out.xml')
