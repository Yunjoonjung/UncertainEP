import numpy as np

from scipy.stats import norm 
from scipy.stats import uniform 
from scipy.stats import triang
from scipy.stats import lognorm

from pyDOE import lhs

from eppy import modeleditor
from eppy.modeleditor import IDF
from eppy.results import readhtml 
import pprint

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

        lighting1 = idf1.idfobjects['LIGHTS'][0]
        
        # Assign the
        for i, X in enumerate(SA_quantified_matrix):
            # Open IDF file
            IDF.setiddname(self.IDD_FileName)
            idf1 = IDF(self.IDF_FileName)

            # Assign quantified values in the idf instace
            lightins.Watts_per_Zone_Floor_Area = X[0]


            # Run
            idf1.save()
            idf = IDF(file_name, epwfile)
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

        return output

    
    def SA(self, number_of_uncertain_parameters, number_of_samples=1000):
        self.number_of_uncertain_parameters = number_of_uncertain_parameters
        self.number_of_samples = number_of_samples

        self.self.distribution_repository = np.zeros((self.number_of_samples,self.number_of_uncertain_parameters))
        
        # 1. Define the variables for Morris sensitivity analysis
        self.energyPlus_input_setup = {'num_vars': self.number_of_uncertain_parameters,
                                      'names': [ '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
                                                 '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
                                                 '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
                                                 '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
                                                 'Temperature', 'Wind Speed' ]}

        # Open .epw file
        EPW_data_repository =[]
        with open(self.epw_FileName, 'r') as f:
            for line in f:
                EPW_data_repository.append(line.strip().split(','))
        f.close()

        # Open .idf file
        IDF.setiddname(self.IDD_FileName)
        incetance_idf = IDF(self.IDF_FileName)
        
                
        # Uncertainty quantification (both parameter and scenario(epw))
        design_lhs = lhs(self.number_of_uncertain_parameters, samples=self.number_of_samples)
        
        for i in range(0,8760):
            # Parameter Uncertainty in EnergyPlus
            self.distribution_repository[:,0] = norm(loc=, scale=3).ppf(self.design_lhs[:,0])
            self.distribution_repository[:,1] = uniform(loc=, scale=15).ppf(self.design_lhs[:,1])
            self.distribution_repository[:,2] = triang(c=0., loc=5, scale=5).ppf(self.design_lhs[:,2])
            self.distribution_repository[:,3] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,3])
            self.distribution_repository[:,4] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,4])

            self.distribution_repository[:,5] = triang(c=0., loc=5, scale=5)..ppf(self.design_lhs[:,5])
            self.distribution_repository[:,6] = uniform(loc=, scale=15).ppf(self.design_lhs[:,6])
            self.distribution_repository[:,7] = triang(c=0., loc=5, scale=5).ppf(self.design_lhs[:,7])
            self.distribution_repository[:,8] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,8])
            self.distribution_repository[:,9] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,9])

            self.distribution_repository[:,10] = norm(loc=, scale=3).ppf(self.design_lhs[:,10])
            self.distribution_repository[:,11] = uniform(loc=, scale=15).ppf(self.design_lhs[:,11])
            self.distribution_repository[:,12] = triang(c=0., loc=5, scale=5).ppf(self.design_lhs[:,12])
            self.distribution_repository[:,13] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,13])
            self.distribution_repository[:,14] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,14])

            self.distribution_repository[:,15] = norm(loc=, scale=3).ppf(self.design_lhs[:,15])
            self.distribution_repository[:,16] = uniform(loc=, scale=15).ppf(self.design_lhs[:,16])
            self.distribution_repository[:,17] = triang(c=0., loc=5, scale=5).ppf(self.design_lhs[:,17])
            self.distribution_repository[:,18] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,18])
            self.distribution_repository[:,19] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,19])


            # Scenario (weather) in EPW file
            self.distribution_repository[i,0] = norm(loc=float(EPW_data_repository[i+8][]), scale=3).ppf(design_lhs[:,0]) # Temperature
            self.distribution_repository[i,1] = norm(loc=float(EPW_data_repository[i+8][]), scale=3).ppf(design_lhs[:,1]) # Solar
            self.distribution_repository[i,2] = norm(loc=float(EPW_data_repository[i+8][]), scale=3).ppf(design_lhs[:,2]) #
            self.distribution_repository[i,3] = norm(loc=float(EPW_data_repository[i+8][]), scale=3).ppf(design_lhs[:,3])
            self.distribution_repository[i,4] = norm(loc=float(EPW_data_repository[i+8][]), scale=3).ppf(design_lhs[:,4])

           
        incetance_idf.save()
        # Conduct Morris Method
        Y = self.EP_iteration(self.distribution_repository)
        Si = morris.analyze(self.energyPlus_input_setup, self.distribution_repository, Y, conf_level=0.95, print_to_console=True, num_levels=4)

        
        # Visualize the morris method result

        
        return

    def UA(self):

        # Conduct uncertainty Analysis
        for i in range(0,self.number_of_samples):
            
            # Open IDF file
            IDF.setiddname(self.IDD_FileName)
            idf1 = IDF(self.IDF_FileName)

            # Assign quantified values in the idf instace
            lightins.Watts_per_Zone_Floor_Area = X[0]


            # Run
            idf1.save()
            idf = IDF(self.IDF_FileName, self.epw_FileName)
            idf.run(readvars=True)

            
        # Visualize the Uncertainty Analysis Results
        
            
