import numpy as np

from scipy.stats import norm 
from scipy.stats import uniform 
from scipy.stats import triang

from pyDOE import lhs # For LHS

import matplotlib.pyplot as plt 
import seaborn as sns



class epSAUA(object):

    def __init__(self, epw_FileName, IDF_FileName, IDD_FileName, number_of_uncertain_parameters, number_of_samples=1000, Draw_SA_Result=True, Draw_UA_Result=True):
        self.epw_FileName = epw_FileName
        self.IDF_FileName = IDF_FileName
        self.IDD_FileName = IDD_FileName
        self.number_of_uncertain_parameters = number_of_uncertain_parameters
        self.number_of_samples = number_of_samples

        self.self.distribution_repository = np.zeros((12,5))

        self.design_lhs = lhs(self.number_of_uncertain_parameters, samples=self.number_of_samples)

        self.energyPlus_input_setup = {'num_vars': self.number_of_uncertain_parameters,
                                       'names': ['Core_ZN_Lights', 'Perimeter_ZN_1_Lights', 'Perimeter_ZN_2_Lights', 'Perimeter_ZN_3_Lights', 'Perimeter_ZN_4_Lights']}

        
    def EnergyPlus_Iteration(self):
       # 1. Open idf file
        IDF.setiddname(self.IDD_FileName)
        idf1 = IDF(self.IDF_FileName)

        idf.run()


    def SA(self):
        # 1. Open idf file
        IDF.setiddname(self.IDD_FileName)
        idf1 = IDF(self.IDF_FileName)

        # 2. Conduct LHS based on the assigned distributions
        ####----------------------------User Modification required ----------------------------####
        # Open .epw file
        with open(self.epw_FileName, 'r') as f:
            data =[]
            for line in f:
                line = line.strip()
                data.append(line.strip().split(','))

##        Lat = float(data[0][6])
##        Lon = float(data[0][7])
##        time_zone = float(data[0][8])
                
        # 2.1 Climate Scenario Uncertainty
        for i in range(0,8760):
            self.distribution_repository[i,0] = norm(loc=float(data[i+][]), scale=3).ppf(self.design_lhs[:,0]) # Temperature
            self.distribution_repository[i,1] = norm(loc=float(data[i+][]), scale=3).ppf(self.design_lhs[:,1]) # Solar
            self.distribution_repository[i,2] = norm(loc=float(data[i+][]), scale=3).ppf(self.design_lhs[:,2]) #
            self.distribution_repository[i,3] = norm(loc=float(data[i+][]), scale=3).ppf(self.design_lhs[:,3])
            self.distribution_repository[i,4] = norm(loc=float(data[i+][]), scale=3).ppf(self.design_lhs[:,4])
            
        f.close()

        # 2.2 Parameter Uncertainty in EnergyPlus 
        self.distribution_repository[:,0] = norm(loc=8, scale=3).ppf(self.design_lhs[:,0])
        self.distribution_repository[:,1] = uniform(loc=7, scale=15).ppf(self.design_lhs[:,1])
        self.distribution_repository[:,2] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,2])
        self.distribution_repository[:,3] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,3])
        self.distribution_repository[:,4] = triang(c=0.5, loc=5, scale=5).ppf(self.design_lhs[:,4])
        ####-----------------------------------------------------------------------------------####


        # 4. Conduct Morris Method
        Y = morris_method(self.distribution_repository)
        Si = morris.analyze(energyPlus_input_setup, self.distribution_repository, Y, conf_level=0.95, print_to_console=True, num_levels=3)


    def UA(self):
        # 1. Conduct uncertainty Analysis
        for i in range(0,self.number_of_samples):
            
            
