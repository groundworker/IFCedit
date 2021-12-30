# -*- coding: utf-8 -*-
"""
Created on Sun Dec 19 18:08:47 2021

@author: Andreas Sorgatz-Wenzel
"""

class IFCELEMENT:
    # special typ of IfcRoot - IfcObjectDefinition - IfcObject - IfcProduct
    def __init__(self, number, ifcproduct, ifcpara):
        # number (location) of product in IFC-file 
        self.Number = number
        # kind of IFC-product
        self.Product = ifcproduct
        # parameter for any subtype of IFCROOT
        self.GlobalId = ifcpara[0]          # 1
        self.OwnerHistory = ifcpara[1]      # 2
        self.Name = ifcpara[2]              # 3
        self.Description = ifcpara[3]       # 4
        # parameter for any subtype of IFCELEMENT
        self.ObjectType = ifcpara[4]        # 5
        self.ObjectPlacement = ifcpara[5]   # 6
        self.Representation = ifcpara[6]    # 7
        self.Tag = ifcpara[7]               # 8

class IFCPROPERTYSET:
    # special typ of IfcRoot -IfcPropertyDefinition - IfcPropertySetDefinition
    def __init__(self, number, ifcproduct, ifcpara):
        # number (location) of product in IFC-file 
        self.Number = number
        # kind of IFC-product
        self.Product = ifcproduct
        # parameter for any subtype of IFCROOT
        self.GlobalId = ifcpara[0]          # 1
        self.OwnerHistory = ifcpara[1]      # 2
        self.Name = ifcpara[2]              # 3
        self.Description = ifcpara[3]       # 4
        # parameter for any subtype of IFCPROPERTYSET
        if len(ifcpara) == 5:
            self.HasProperties = ifcpara[4]     # 5   

class IFCRELDEFINESBYPROPERTIES:
    # special typ of IfcRelationship - IfcRelDefines
    def __init__(self, number, ifcproduct, ifcpara):
        # number (location) of product in IFC-file 
        self.Number = number
        # kind of IFC-product
        self.Product = ifcproduct
        # parameter for any subtype of IFCROOT
        self.GlobalId = ifcpara[0]          # 1
        self.OwnerHistory = ifcpara[1]      # 2
        self.Name = ifcpara[2]              # 3
        self.Description = ifcpara[3]       # 4
        # parameter for IFCRELDEFINESBYPROPERTIES
        self.RelatedObjects = ifcpara[4]    # 5
        self.RelatingPropertyDefinition = ifcpara[5]    # 6
        
class IFCPROPERTYSINGLEVALUE:
    # special typ of IfcPropertyAbstraction - IfcProperty - IfcSimpleProperty
    def __init__(self, number, ifcproduct, ifcpara):
        # number (location) of product in IFC-file 
        self.Number = number
        # kind of IFC-product
        self.Product = ifcproduct
        # parameter for any subtype of IfcProperty
        self.Name = ifcpara[0]          # 1
        self.Description = ifcpara[1]      # 2
        self.NominalValue = ifcpara[2]      # 3
        if len(ifcpara) == 4:
            self.Unit = ifcpara[3]              # 4