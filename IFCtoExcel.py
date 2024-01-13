import sys

# See if the filename has been provided as a command line argument and set the file variable otherwise set the file variable to the default file
if len(sys.argv) > 1:
    file_name = sys.argv[1]
else:
    file_name = "c:\\Temp\\SMWSASBT-CPG-SWD-CP000-TU-M3D-020751.ifc"

# determine the file name minus the file extension and path from the file_name variable
file_name_no_ext = file_name.split("\\")[-1].split(".")[0]

import ifcopenshell
import ifcopenshell.util.element as Element
file = ifcopenshell.open(file_name)

# determine the file name minus the file extension and path from the file_name variable
file_name_no_ext = file_name.split("\\")[-1].split(".")[0]

def get_objects_data_by_class(file, class_type):
    def add_pset_attributes(psets):
        for pset_name, pset_data in psets.items():
            for property_name in pset_data.keys():
                pset_attributes.add(f'{pset_name}.{property_name}')

    pset_attributes = set()
    objects_data = []
    objects = file.by_type(class_type)

    for object in objects:
        psets = Element.get_psets(object, psets_only=True)
        add_pset_attributes(psets)
        qtos = Element.get_psets(object, qtos_only=True)
        add_pset_attributes(qtos)

        object_id = object.id()
        objects_data.append({
            "Express ID": object.id(),
            "GlobalId": object.GlobalId,
            "Class": object.is_a(),
            "PredefinedType": Element.get_predefined_type(object),
            "Name": object.Name,
            "Level": Element.get_container(object).Name
            if Element.get_container(object)
            else "",
            "Type": Element.get_type(object).Name
            if Element.get_type(object)
            else "",
            "QuantitySets": qtos,
            "PropertySets": psets
        })
    return objects_data, list(pset_attributes)

def get_attribute_value(object_data, attribute):
    if "." not in attribute:
        return object_data[attribute]
    elif "." in attribute:
        pset_name = attribute.split(".",1)[0]
        prop_name = attribute.split(".",-1)[1]
        if pset_name in object_data["PropertySets"].keys():
            if prop_name in object_data["PropertySets"][pset_name].keys():
                return object_data["PropertySets"][pset_name][prop_name]
            else:
                return None
        if pset_name in object_data["QuantitySets"].keys():
            if prop_name in object_data["QuantitySets"][pset_name].keys():
                return object_data["QuantitySets"][pset_name][prop_name]
            else:
                return None
    else:
        return None

data, pset_attributes = get_objects_data_by_class(file, "IfcBuildingElement")

attributes = ["Express ID", "GlobalId", "Class", "PredefinedType", "Name", "Level", "Type"] + pset_attributes

pandas_data = []
for object_data in data:
    row = []
    for attribute in attributes:
        value = get_attribute_value(object_data, attribute)
        row.append(value)
    pandas_data.append(tuple(row))

import pandas as pd
dataframe = pd.DataFrame.from_records(pandas_data, columns=attributes)

#dataframe.to_csv("./out/test.csv")
writer = pd.ExcelWriter("./out/"+file_name_no_ext+".xlsx", engine='xlsxwriter')

for object_class in dataframe["Class"].unique():
    df_class = dataframe[dataframe["Class"] == object_class]
    df_class = df_class.dropna(axis=1, how='all')
    df_class.to_excel(writer, sheet_name=object_class)
writer.save()

print("Executed!")