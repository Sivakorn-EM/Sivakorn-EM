#This code (RR_to_Fill.py) creates two output .csv files to fill up the MV SQL DB Xref Bulk Uploader directly from RunROMEO XREF_Masters tab.

#Input files: \
##Filled RunROMEO XREF_Masters tab. -- Rows containing "*" or "Select" in the Alias column will not be included.\
##Unit_Dict: containing a dictionary to map units to type of units. \
##PIMS ROM map: contains a dictionary to map ROMEO tags to PIMS tags. 

#Output files: Two csv files containing page 1 and page 2 of MV SQL DB Xref Bulk Uploader. Please copy the .csv files into the template mechanically and double check that all values are right.
