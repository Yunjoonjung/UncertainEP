import numpy as np 

from eppy import modeleditor
from eppy.modeleditor import IDF
from eppy.results import readhtml 
import pprint



idd_file = "Energy+.idd"
file_name = "New_TGS.idf"
epwfile = "Atlanta.epw"

IDF.setiddname(idd_file)
##idf1 = IDF(file_name)

##
##lighting1 = idf1.idfobjects['lights'][0]
##lighting2 = idf1.idfobjects['LIGHTS'][1]
##lighting3 = idf1.idfobjects['LIGHTS'][2]
##lighting4 = idf1.idfobjects['LIGHTS'][3]
##lighting5 = idf1.idfobjects['LIGHTS'][4]

##a = lighting1.Watts_per_Zone_Floor_Area
##print(a)
##print(type(a))

##People = idf1.idfobjects['people'][0]
##b = People.Zone_Floor_Area_per_Person
##print(b)

##outdoor = idf1.idfobjects['DesignSpecification:OutdoorAir'][0]
##c = outdoor.Outdoor_Air_Flow_per_Person
##
##windowexample = idf1.idfobjects['WindowMaterial:SimpleGlazingSystem'][0]
##d = windowexample.UFactor
##e= windowexample.Solar_Heat_Gain_Coefficient
####print(d)
##idf1.idfobjects['WindowMaterial:SimpleGlazingSystem'][0].UFactor = 15 # It works


##idf1.save()

idf = IDF(file_name, epwfile)
idf.run()
