import sys
import os

# See if the filename has been provided as a command line argument and set the file variable otherwise set the file variable to the default file
if len(sys.argv) > 1:
    file_name = sys.argv[1]
else:
    file_name = "c:\\Temp\\SMWSASCA-CPU-SWD-EW300-SD-M3D-263102.ifc"

def extract_csv(bytesio):
    # determine the file name minus the file extension and path from the file_name variable
    #file_name_no_ext = file_name.split("\\")[-1].split(".")[0]

    import ifcopenshell
    import ifcopenshell.util.element as Element
    import tempfile
    #file = ifcopenshell.open(file_name)

    with tempfile.NamedTemporaryFile(suffix=".ifc", delete=False) as tmp_file:
        tmp_file.write(bytesio.read())
        tmp_file_name = tmp_file.name

    file = ifcopenshell.open(tmp_file_name)
    filename_without_extension, _ = os.path.splitext(bytesio.name)

    def get_objects_data_by_class(file, class_type):
        def add_pset_attributes(psets):
            for pset_name, pset_data in psets.items():
                for property_name in pset_data.keys():
                    append_attribute_data(object, pset_name+'.'+property_name, pset_data[property_name])

        def append_attribute_data(object, attribute, value):
            objects_data.append({
                "Model Name": filename_without_extension,
                "Express ID": object.id(),
                "Field Name": attribute,
                "Field Value": value
            })

        objects_data = []
        objects = file.by_type(class_type)

        for object in objects:
            append_attribute_data(object, "GlobalId", object.GlobalId)
            append_attribute_data(object, "Class", object.is_a())
            append_attribute_data(object, "PredefinedType", Element.get_predefined_type(object))
            append_attribute_data(object, "Name", object.Name)
            if Element.get_container(object):
                append_attribute_data(object, "Level", Element.get_container(object).Name)
            if Element.get_type(object):
                append_attribute_data(object, "Type", Element.get_type(object).Name)
            psets = Element.get_psets(object, psets_only=True)
            add_pset_attributes(psets)
            qtos = Element.get_psets(object, qtos_only=True)
            add_pset_attributes(qtos)
            
        return objects_data

    data = get_objects_data_by_class(file, "IfcBuildingElement")

    import pandas as pd
    dataframe = pd.DataFrame.from_records(data)

    return dataframe