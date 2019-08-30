from Uncertain_EP import Uncertain_EP


# Execution
testing = Uncertain_EP('C:\\Users\\YunJoon Jung\\Dropbox (GaTech)\\Uncertain_EP\\New_TGS.idf', 'C:\\Users\\YunJoon Jung\\Dropbox (GaTech)\\Uncertain_EP\\Atlanta2.epw',
                       'C:\\Users\\YunJoon Jung\\Dropbox (GaTech)\\Uncertain_EP\\Energy+.idd', 'New_TGSTable.htm')

##testing.SA(number_of_SA_samples=22)
testing.UA(number_of_UA_samples=10, only_idf_instances_generation=False)
