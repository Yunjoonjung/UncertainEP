from eppy import modeleditor
from eppy.modeleditor import IDF


iddfile = "Energy+.idd"
fname1= 'ASHRAE90.1_OfficeSmall_STD2004_PortAngeles.idf'
epwfile = "Atlanta.epw"


idf = IDF.setiddname(iddfile)

idf = IDF(fname1,epwfile)
idf.run(readvars=True)
