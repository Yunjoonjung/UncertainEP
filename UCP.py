import numpy as np

from scipy.stats import norm 
from scipy.stats import uniform 
from scipy.stats import triang

from pyDOE import lhs

from eppy import modeleditor
from eppy.modeleditor import IDF
from eppy.results import readhtml 
import pprint

import openpyxl

import seaborn as sns
import matplotlib.pyplot as plt


class Uncertain_EP(object):

    def __init__(self, IDF_FileName, epw_FileName, IDD_FileName, IDD_FileNameDraw_SA_Result=True, Draw_UA_Result=True):
        self.epw_FileName = epw_FileName
        self.IDF_FileName = IDF_FileName
        self.IDD_FileName = IDD_FileName
        self.IDD_FileNameDraw_SA_Result = IDD_FileNameDraw_SA_Result
        self.Draw_UA_Result = Draw_UA_Result
        
        
    def EP_iteration(self, SA_quantified_matrix): # Method for EnergyPlus iteration
       
        # Assign the quantified values into idf and epw
        for i, X in enumerate(SA_quantified_matrix):
            # Open .idf file
            IDF.setiddname(self.IDD_FileName)
            incetance_idf = IDF(self.IDF_FileName)

            # Assign quantified values in the idf instace
            if uncertain_input_sheet.cell(row=i+2, column=2).value != "Climate":
                for j in range(0,uncertain_input_sheet.cell(row=1, column=14).value):
                    replace_to_EnergyPlus_format = str(uncertain_input_sheet.cell(row=i[0]+2, column=2).value.replace("_", ":"))
                    obj_num = int(uncertain_input_sheet.cell(row=i[0]+2, column=4).value[3])
                    name_in_idf_instance = incetance_idf.idfobjects[replace_to_EnergyPlus_format][obj_num-1].uncertain_input_sheet.cell(row=i[0]+2, column=3).value

                    print(name_in_idf_instance)
                                                                                 
                    name_in_idf_instance = X[j]
                    
            else: # Climate case
                


            # Run
            incetance_idf.save()
            idf = IDF(self.IDF_FileName, self.epw_FileName)
            idf.run(readvars=True)
            
            
            # Collect the result
            fname = "eplustbl.htm" # the html file you want to read
            filehandle = open(fname, 'r').read() # get a file handle to the html file
            htables = readhtml.titletable(filehandle) # reads the tables with their titles
            firstitem = htables[0]

            pp = pprint.PrettyPrinter()
            pp.pprint(firstitem)
            print(firstitem[1][1][1])
            total_site_energy = firstitem[1][1][1]
            output[i] = total_site_energy

        a = sns.distplot(output[:])
        plt.show()
        
        incetance_idf.save() # close idf 
        uncertain_input.save('Uncertain_EP_Input.xlsx') # close Excel
        return output

    
    def SA(self, number_of_samples=10):
        self.number_of_samples = number_of_samples

        # Read the "Uncertain_EP_Input" excel file
        uncertain_input = openpyxl.load_workbook('Uncertain_EP_Input.xlsx', data_only=False)
        uncertain_input_sheet = uncertain_input['Input']

        self.number_of_uncertain_parameters = uncertain_input_sheet.cell(row=1, column=14).value
        self.self.distribution_repository = np.zeros((self.number_of_samples,self.number_of_uncertain_parameters))

        print(self.number_of_samples)
        print(type(self.number_of_samples ))

        # Open .epw file
        EPW_data_repository =[]
        with open(self.epw_FileName, 'r') as f:
            for line in f:
                EPW_data_repository.append(line.strip().split(','))
        f.close()

        # Open .idf file
        IDF.setiddname(self.IDD_FileName)
        incetance_idf = IDF(self.IDF_FileName)
        
        # Define the variables for Morris sensitivity analysis
        self.energyPlus_input_setup = {'num_vars': self.number_of_uncertain_parameters}

        name_repository = []
        for i in range(0,self.number_of_uncertain_parameters):
            if uncertain_input_sheet.cell(row=i+2, column=2).value != "Climate":
                replace_to_EnergyPlus_format = str(uncertain_input_sheet.cell(row=i+2, column=2).value.replace("_", ":"))
                obj_num = int(uncertain_input_sheet.cell(row=i+2, column=4).value[3])
                name_in_idf_instance = incetance_idf.idfobjects[replace_to_EnergyPlus_format][obj_num-1].Name
                sample_string = name_in_idf_instance +"__IN__"+ str(uncertain_input_sheet.cell(row=i+2, column=3).value)
                name_repository.append(sample_string)
            else:
                name_for_climate = uncertain_input_sheet.cell(row=i+2, column=2).value
                name_repository.append(name_for_climate)


        self.energyPlus_input_setup.update('names'=name_repository)
        
        # Uncertainty quantification
        design_lhs = lhs(self.number_of_uncertain_parameters, samples=self.number_of_samples)
        
        for i in range(i,self.number_of_uncertain_parameters):
            if uncertain_input_sheet.cell(row=i+2, column=2).value != "Climate":
                if uncertain_input_sheet.cell(row=i+2, column=5).value == "NormalRelative":
                    replace_to_EnergyPlus_format = str(uncertain_input_sheet.cell(row=i+2, column=2).value.replace("_", ":"))
                    obj_num = int(uncertain_input_sheet.cell(row=i+2, column=2).value[3])
                    value_in_idf_instance = incetance_idf.idfobjects[replace_to_EnergyPlus_format][obj_num-1].str(uncertain_input_sheet.cell(row=i+2, column=3).value)
                    
                    self.distribution_repository[:,i] = norm(loc=value_in_idf_instance, scale=uncertain_input_sheet.cell(row=i+2, column=6).value).ppf(design_lhs[:,i])
                    
                elif uncertain_input_sheet.cell(row=i+2, column=5).value == "UniformRelative":                
                    self.distribution_repository[:,i] = uniform(loc=uncertain_input_sheet.cell(row=i+2, column=6).value, scale=uncertain_input_sheet.cell(row=i+2, column=7).value).ppf(design_lhs[:,i])
                    
                elif uncertain_input_sheet.cell(row=i+2, column=5).value == "TriangleRelative":
                    self.distribution_repository[:,i] = triang(c=uncertain_input_sheet.cell(row=i+2, column=6).value, loc=uncertain_input_sheet.cell(row=i+2, column=7).value, scale=uncertain_input_sheet.cell(row=i+2, column=8).value).ppf(design_lhs[:,i])

            else:
                if uncertain_input_sheet.cell(row=i+2, column=3).value == "Temperature":
                    if uncertain_input_sheet.cell(row=i+2, column=5).value == "NormalRelative":
                        self.distribution_repository[:,i] = norm(loc=      , scale=uncertain_input_sheet.cell(row=i+2, column=6).value).ppf(design_lhs[:,i])
                    
                    elif uncertain_input_sheet.cell(row=i+2, column=5).value == "UniformRelative":                
                        self.distribution_repository[:,i] = uniform(loc=uncertain_input_sheet.cell(row=i+2, column=6).value, scale=uncertain_input_sheet.cell(row=i+2, column=7).value).ppf(design_lhs[:,i])
                    
                    elif uncertain_input_sheet.cell(row=i+2, column=5).value == "TriangleRelative":
                        self.distribution_repository[:,i] = triang(c=uncertain_input_sheet.cell(row=i+2, column=6).value, loc=uncertain_input_sheet.cell(row=i+2, column=7).value, scale=uncertain_input_sheet.cell(row=i+2, column=8).value).ppf(design_lhs[:,i])

                elif uncertain_input_sheet.cell(row=i+2, column=3).value == "Relative_Humidity":


                elif uncertain_input_sheet.cell(row=i+2, column=3).value == "Atmospheric_Pressure":


                elif uncertain_input_sheet.cell(row=i+2, column=3).value == "Horizontal_Solar_Radiation":


                elif uncertain_input_sheet.cell(row=i+2, column=3).value == "Direct_Solar_Radiation":


                elif uncertain_input_sheet.cell(row=i+2, column=3).value == "Diffuse_Solar_Radiation":

                elif uncertain_input_sheet.cell(row=i+2, column=3).value == "Wind_Speed":


                

         
        incetance_idf.save() # close idf 
        uncertain_input.save('Uncertain_EP_Input.xlsx') # close Excel

        
        # Conduct Morris Method
        Y = self.EP_iteration(self.distribution_repository)
        Si = morris.analyze(self.energyPlus_input_setup, self.distribution_repository, Y, conf_level=0.95, print_to_console=True, num_levels=4)


        # Visualize the morris method result

        


##    def UA(self):
##
##        # Conduct uncertainty Analysis
##        for i in range(0,self.number_of_samples):
##            
##            # Open IDF file
##            IDF.setiddname(self.IDD_FileName)
##            idf1 = IDF(self.IDF_FileName)
##
##            # Assign quantified values in the idf instace
##            lightins.Watts_per_Zone_Floor_Area = X[0]
##
##
##            # Run
##            idf1.save()
##            idf = IDF(self.IDF_FileName, self.epw_FileName)
##            idf.run(readvars=True)

            
        # Visualize the Uncertainty Analysis Results
        
            
