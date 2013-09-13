# Introduction
PatMap is designed to measure distances and travel times between demand and service locations in healthcare.  Demand locations are typically patients’ homes and service locations will typically be hospitals, clinics or other service providers.
There are two main functions of the tool and in both cases, a Google Maps interface embedded within the tool allows for a visual representation of demand and service locations.  Firstly, distances and travel times between demand and service locations can be calculated.  The user has the option to select the mode of travel for these calculations; driving, walking or cycling.  The imagery of the demand locations may be altered in order to reflect density of demand at the location.  Secondly, the user can specify a driving time or distance limit, and calculate which demand locations are within the limit and which are not, for a user-selected service location.  These can also 
This tool was developed by members of the School of Mathematics at Cardiff University.  The project was funded by the Cardiff Undergraduate Research Opportunities Programme (CUROP) scheme.  

# System requirements
PatMap does not require any specialist software.  It is a standalone user interface was developed and tested on a Windows 7 machine.  Compatibility with other Windows operating systems is currently under investigation.
Data input files can be comma-separated value (CSV) files or Microsoft Excel spreadsheets.  Files with the following extensions are supported by PatMap: .csv, .xls and .xlsx.
Output files created by PatMap are CSV and JPEG file types.  Excel spreadsheets are also appended to, but not created.  CSV files are Excel compatible; for more information on using CSV files within Microsoft Excel, see http://office.microsoft.com/en-gb/excel-help/import-or-export-text-txt-or-csv-files-HP010099725.aspx.

# Google Maps API
PatMap utilises a variety of Google Maps applications.  The four application programming interfaces (APIs) used are Distance Matrix, Static Maps, Geocoding and Google Maps Javascript API version 3.   The API gives users a means of embedding Google Maps into web pages and user-built interfaces, allowing for customisable options.
Note that the use of any Google Maps API product is limited by Google’s terms of service.  Please refer to the following webpage for further details: https://developers.google.com/maps/terms.

# Python 
Python 2.7, an object-orientated programming language, was used to implement the Google Maps API into PatMap, and to develop the tool itself.  This is a widely used general purpose programming language, and enabled PatMap to be developed without the requirement for any specialist software.  The user interface was developed using the extension module WxPython, while py2exe was used to create the Windows executable program.

# Installation

## Windows

To install PatMap onto your machine, extract all files from the provided zip file and save to a local drive.  You will only need to do this once.  
Click on the .exe file to open PatMap.  
See the contact information at the end of this document should you require further assistance.
 
## Other 

Coming soon. 
 
# Using PatMap
There are three tasks which must be completed each time PatMap is used.  Tasks 1 and 2 involve uploading and/or entering demand and service location information.  The user selects various output options in Task 3, including the type of output and the location of saved output files.
There are two output functionalities, distance and limit.  

# Data preparation
Before making full use of PatMap, some preparation of input data may be required.  While PatMap is designed to handle erroneous data, such as non-existent postcodes, data cleaning is advised prior to input in order to maximise usability.  

# Postcode formats
British postcodes are structured hierarchically, supporting four levels of geographic segmentation; area, district, sector and unit.  An example is given in Figure ….
It is recommended that unit postcodes are used whenever possible.  Sector postcodes often cause problems within PatMap; using the above postcode as an example, entering AB12 3 into Google Maps returns results with ‘3’ in the address within the AB12 district.  If postcodes are stored at sector level, it is recommended that they are truncated to district level.  In this case, Google Maps will return results based on the postcode centroid (at district level).

