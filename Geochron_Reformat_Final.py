# -*- coding: utf-8 -*-
"""
The purpose of this script is to reformat old data scheme (age_date_100k) to new data scheme (geochron) for WGS
Alex Wernle
Created on Tue Nov 19 09:52:27 2019
@author: awer490
"""
# Import modules
import pandas as pd

# Read in the excel files
df = pd.read_excel('age_date_100k.xlsx')
df2 = pd.read_excel('geochron.xlsx') 

# Define your columns from excel docs
Geo_Age = df['GEOLOGIC_AGE_VALUE'] 
Date_sample = df['DATE_SAMPLE_COMMENT'] 

#Geologic_Age_Data-------------------------------------------------------------

# Create empty arrays to fill with new data 
Numeric_Age= []
Age_Error= []
Age_Units= []

# Loop through age data in specified column
for age in Geo_Age:
    
    #Substitute number for geologic age
    if age.find('Late Jurassic') != -1: 
        age = "154.3 +/- 9.2 my"
    if age.find('Eocene') != -1: 
        age = "45.0 +/- 11.0 my"    
    if age.find('Late Eocene') != -1:  
        age = "36.0 +/- 2.0 my"
    if age.find('Middle Eocene') != -1:  
        age = "42.9 +/- 4.9 my"
    if age.find('Cretaceous') != -1: 
        age = "105.5 +/- 39.5 my"
    if age.find('Early Cretaceous') != -1: 
        age = "122.8 +/- 22.2 my"
    if age.find('Jurassic') != -1:    
        age = "173.2 +/- 28.1 my"
    if age.find('Triassic') != -1:    
        age = "226.2 +/- 25.3 my"
    if age.find('Late Triassic') != -1:   
        age = "219.2 +/- 17.8 my"
    if age.find('Oligocene') != -1:   
        age = "28.5 +/- 5.4 my"
    if age.find('Permian') != -1:   
        age = "275.2 +/- 23.3 my"    
    if age.find('Mesozoic') != -1:   
        age = "159.0 +/- 93.0 my"
    if age.find('Mississippian') != -1:   
        age = "341.0 +/- 17.9 my"
    if age.find('Devonian') != -1:   
        age = "389.0 +/- 30.2 my"
    if age.find('Paleocene') != -1:   
        age = "61.0 +/- 5.0 my"
    if age.find('Pennsylvanian') != -1:   
        age = "311.0 +/- 12.2 my"
        

    # Split data with common delineator first
    split1=age.split("+/-")  

    #split for multiple dates
    if len(split1) > 2:
        split1=age.split(",")
        split1= split1[0].split("+/-") 

    # Split again with first space
    if len(split1) ==2:
        for item in split1[1:]:
            split2 = (item.split(" ",1))
            
            # fix the occasional extra space
            if len(split2) ==3:
                split2=''.join(split2[1:2])
            stritem1=(split2[0]).split(sep=None, maxsplit=-1)
            
            # fix the occasional missing space
            if len(split2) ==2:
                stritem2=(split2[1]).split(sep=None, maxsplit=-1)  
                
            #fix the extra split cases    
            if len(stritem2)>=3: 
                stritem2=' '.join(stritem2[1:])
                stritem2=stritem2.split("*")
            stritem3=(split1[0]).split(sep=None, maxsplit=-1)
            
            #Add all the string items together
            finalstring= stritem1 + stritem2+ stritem3
            print(finalstring)
            
            #Append data to empty arrays
            Numeric_Age.append(finalstring[2])
            Age_Error.append(finalstring[0])
            Age_Units.append(finalstring[1])
    
    # If no common delineator, append as is
    else:
        finalstring= age.split("*")# Split with random delineator to turn str to list
        finalstring=' '.join(finalstring[0:]) #gets rid of brackets
        Numeric_Age.append(finalstring[0:])
        Age_Error.append("NAN")
        Age_Units.append("NAN")

#Date_Sample_Data--------------------------------------------------------------        

# Create empty arrays to fill with new data         
Material_Anlysis= []
Alt_SampleID = []
DT_MET_CD = []

#Loop through data in Date_Sampled column
for data in Date_sample:  
    
    # Split with common delineator
    split1=data.split("Sample No.")

    #For normal split cases, turn to strings
    if len(split1) ==2:
        stritem1=(split1[0]).split(sep=None, maxsplit=-1)
        split2=' '.join(stritem1[0:])
        stritem2=(split1[1]).split(sep=None, maxsplit=-1)
        split3=' '.join(stritem2[0:])
        
        #Append data to empty array
        Material_Anlysis.append(split2)
        Alt_SampleID.append(split3)

    #If no common delineator, append as is
    else:
        split1=' '.join(split1[0:]) #gets rid of brackets
        Material_Anlysis.append(split1)
        Alt_SampleID.append("NAN")

for data in Date_sample:  
     # Write DT_MET_CD with correct codes based on strings in Date Sample
    if data.find('radiometric') != -1: 
        DT_MET_CD.append(1)
    elif data.find('fossil') != -1:  
        DT_MET_CD.append(2)
    elif data.find('zircon') != -1:  
        DT_MET_CD.append(3)
    elif data.find('K-Ar') != -1:  
        DT_MET_CD.append(4)
    elif data.find('U-Pb') != -1:   
        DT_MET_CD.append(5)
    elif data.find('14C') != -1:   
        DT_MET_CD.append(6)
    elif data.find('luminescence') != -1: 
        DT_MET_CD.append(7)
    elif data.find('amino') != -1:  
        DT_MET_CD.append(8)
    elif data.find('Ar-Ar') != -1:  
        DT_MET_CD.append(9)
    elif data.find('Rb-Sr') != -1:  
        DT_MET_CD.append(10)
    elif data.find('optically stimulated luminescence') != -1:   
        DT_MET_CD.append(11)
    elif data.find('tephra') != -1:   
        DT_MET_CD.append(12)
    elif data.find('2.5-minute') != -1:   
        DT_MET_CD.append(13)    
    elif data.find('infrared stimulated luminescence') != -1:   
        DT_MET_CD.append(14)
    else:
        DT_MET_CD.append("NA")

    
# Overwrite to existing excel doc and save as new excel doc
df3 = pd.read_excel('age_date_100k.xlsx') 
df3.to_excel("output.xlsx") 
fn = r'C:/Users/awer490/Desktop/AW_SSSP/Python_ish/Geochron_Table_Reformat/output.xlsx'

# convert arrays to series to match index lengths in excel doc
Numeric_Age= pd.Series(Numeric_Age) 
Age_Error= pd.Series(Age_Error)
Age_Units= pd.Series(Age_Units)
Material_Anlysis= pd.Series(Material_Anlysis)
Alt_SampleID = pd.Series(Alt_SampleID)
DT_MET_CD = pd.Series(DT_MET_CD)

#Convert series to dataframe and insert as columns into excel doc
df3.insert(3, "NumericAge", Numeric_Age, True) 
df3.insert(4, "AgePlusError", Age_Error, True)
df3.insert(5, "AgeMinusError", Age_Error, True)
df3.insert(6, "AgeUnits", Age_Units, True)
df3.insert(8, "MaterialAnalyzed", Material_Anlysis, True)
df3.insert(9, "AlternateSampleID",Alt_SampleID, True)
df3.insert(11, "DT_MET_CD",DT_MET_CD, True)
df3.to_excel("output.xlsx") 

#End---------------------------------------------------------------------------