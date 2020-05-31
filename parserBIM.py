import sys
import ifcopenshell
import xlsxwriter
import datetime

file = ifcopenshell.open("TEST.ifc")
ifc_types = ['IfcProduct']
prop_keys = []
instances = []
mdataobjs = []

for ifc_type in ifc_types:
        print('Exporting ' + ifc_type+'...')
        #Get all elements for current type
        elements = file.by_type(ifc_type)
        #Define a dictionary for storing current element
        instance = {}
        for element in elements:
                mdataobj = {}
                prop_sets = []
                prop_sets = element.IsDefinedBy
                for prop_set in prop_sets:
                        if (prop_set.is_a('IfcRelDefinesByProperties')):
                                properties = prop_set.RelatingPropertyDefinition.HasProperties
                                for prop in properties:
                                        if (prop.Name.startswith("mdata_")):
                                                mdataobj['GlobalID'] = element.GlobalId
                        if prop_set.is_a('IfcRelDefinesByType'):
                                if prop_set.RelatingType.HasPropertySets != None:
                                        type_prop_sets = prop_set.RelatingType.HasPropertySets
                                        for type_prop_set in type_prop_sets:
                                                if (type_prop_set.is_a('IfcPropertySet')):
                                                        properties = type_prop_set.HasProperties
                                                        for prop in properties:
                                                                if (prop.Name.startswith("mdata_")):
                                                                        mdataobj['GlobalID'] = element.GlobalId
                if mdataobj != {}:
                        mdataobjs.append(mdataobj)

        for element in elements:
                instance = {}
                instance_properties = {}
                prop_sets = []
                prop_sets = element.IsDefinedBy
                for mdataobj in mdataobjs:
                        if mdataobj['GlobalID'] == element.GlobalId:
                                instance['Name'] = element.Name
                                instance['Type'] = element.is_a()
                                instance['GlobalID'] = element.GlobalId
                                for prop_set in prop_sets:
                                        if (prop_set.is_a('IfcRelDefinesByProperties')):
                                                properties = prop_set.RelatingPropertyDefinition.HasProperties
                                                for prop in properties:
                                                        instance_properties[prop.Name] = prop.NominalValue.wrappedValue
                                                        if prop.Name not in prop_keys:
                                                                prop_keys.append(prop.Name)
                                        if prop_set.is_a('IfcRelDefinesByType'):
                                                if prop_set.RelatingType.HasPropertySets != None:
                                                        type_prop_sets = prop_set.RelatingType.HasPropertySets
                                                        for type_prop_set in type_prop_sets:
                                                                if (type_prop_set.is_a('IfcPropertySet')):
                                                                        properties = type_prop_set.HasProperties
                                                                        for prop in properties:
                                                                                instance_properties[prop.Name] = prop.NominalValue.wrappedValue
                                                                                if prop.Name not in prop_keys:
                                                                                        prop_keys.append(prop.Name)
                                instance['properties'] = instance_properties
                                instances.append(instance)
                
