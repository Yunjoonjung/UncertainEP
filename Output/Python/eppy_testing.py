from eppy import modeleditor
from eppy.modeleditor import IDF
import eppy
import witheppy.runner


iddfile = "Energy+.idd"
fname1= 'instance.idf'
epwfile = "Atlanta2018.epw"


##IDF.setiddname(iddfile)

##idf = IDF(fname1,epwfile)

# Change People
# Zone_Floor_Area = [16,18,29,22,24] 
##for i in [16,18,29,22,24]:
##    idf = eppy.openidf(fname1, epw=epwfile)
##    idf.idfobjects["People"][0].Zone_Floor_Area_per_Person = i##    a = idf.idfobjects["People"][0].Zone_Floor_Area_per_Person
##    print(a)
##
##    idf.save()
##
##    witheppy.runner.eplaunch_run(idf)

##    fname = 'instanceTable.htm' # the html file you want to read
##    filehandle = open(fname, 'r').read() # get a file handle to the html file
##    htables = readhtml.titletable(filehandle) # reads the tables with their titles
##    firstitem = htables[0]
##    seconditem = htables[2]
##
##    pp = pprint.PrettyPrinter()
##    total_site_energy = firstitem[1][1][1]
##    total_conditioned_area = seconditem[1][2][1]
##    print(total_conditioned_area)   

##idf.run(readvars=True)
##for i in range(5):
filename = ["idf_instance_1.idf","idf_instance_2.idf","idf_instance_3.idf","idf_instance_4.idf","idf_instance_5.idf"]
##    a = (filename[i])
##    print(a)

idf = eppy.openidf(filename[4], epw=epwfile)
witheppy.runner.eplaunch_run(idf)
