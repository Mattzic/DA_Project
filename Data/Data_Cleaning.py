import numpy as np
import pandas as pd
import re
from numpy.random import default_rng

# pd.set_option('max_columns',100) 
# pd.set_option('max_rows',500)  
File_Path = 'NZ_Admin_JOBS.xlsx'
Reading_Engine = 'openpyxl'
Columns = ['Job_Title', 'Link', 'Orgnisation', 'Location', 'Time_Posted', 'Classification']
New_Order = ['Job_Title', 'Link', 'Orgnisation', 'Location', 'Area', 'Classification', 'Subclassification', 'Lo_Salary', 'Hi_Salary', 'Time_Posted']
File_Save = 'NZ_Admin_JOBS_finished.xlsx'

def Read_Data():
    data = pd.read_excel(File_Path, engine=Reading_Engine)
    # data.drop('Unnamed: 0',axis=1, inplace=True)
    return data

#=========================================================================================================================================    

# Assign names to columns
def Name_Columns(data):
    data.columns = Columns
    return data
    
#=========================================================================================================================================      

# Fill Nan's in 'Orgnisation' with 'Private Advertiser'
def Set_Orgnisation(data):
    data.Orgnisation.fillna('Private Advertiser', inplace=True)
    return data
    
#=========================================================================================================================================    

# Extract salary info from 'Classification' and make a new column 'Salary'
def Extract_Salary(data):
    def apply_saraly(a):
        if 'classification:' in a:
            return 'None'
        else:
            return a
    data['Salary'] = data.Classification.apply(apply_saraly)
    return data

#=========================================================================================================================================    
  
# Clean salary info in 'Classificaiton'
def Clean_SaInfo(data):
    def clean_salary(a):
        if 'classification:' in a:
            return a
        else:
            return 'None'
    data.Classification = data.Classification.apply(clean_salary)
    return data

#=========================================================================================================================================      

def deduplication(x):
    x = x.strip()
    if x == 'None':
        return 'None'
    return re.match(r'(.*)\1', x).group(1)
    
#=========================================================================================================================================    
    
def Clean_Classification(data):
    data = Extract_Salary(data)
    data = Clean_SaInfo(data)
    
    # Split 'Classification' into 'Subclassification' and 'Classification' 
    data[['Classification', 'Subclassification']] = data.Classification.str.split('subClassification:', expand=True)
    
    # Get rid of 'classification:'
    data.Classification = data.Classification.str.replace('classification:','')
    
    # Set 'Subclassification' dtype to str and deduplicate 'Subclassification'
    data.Subclassification = data.Subclassification.astype('str')
    data.Subclassification = data.Subclassification.apply(deduplication)
    
    # Set 'Classification' dtype to str and deduplicate 'Classification'
    data.Classification = data.Classification.astype('str')
    data.Classification = data.Classification.apply(deduplication)

    data.Classification = data.Classification.astype('str')
    data.Classification = data.Classification.apply(CleanNone) 
    
    data.Subclassification = data.Subclassification.astype('str')
    data.Subclassification = data.Subclassification.apply(CleanNone)

    return data

#=========================================================================================================================================        

def Clean_Location(data):
    
    # Split 'Location' into 'Location' and 'Area' 
    data[['Location', 'Area']] = data.Location.str.split('area:', expand=True)
    
    # Get rid of 'location:' in 'Location' and deduplicate 'Location'
    data.Location = data.Location.str.replace('location:','')
    data.Location = data.Location.apply(deduplication)

    data.Location = data.Location.astype('str')
    data.Location = data.Location.apply(CleanNone)    
    return data

#=========================================================================================================================================        

def Clean_Area(data):
    
    # Get rid of salary info from 'Area'
    data['Area'] = data.Area.str.replace(r'(.*[a-z]),.+', r'\1', regex=True)
    
    # Set 'Area' dtype to str and deduplicate 'Area'
    data.Area = data.Area.astype('str')
    data.Area = data.Area.apply(deduplication)

    data['Area'] = data['Area'].astype('str')
    data['Area'] = data['Area'].apply(CleanNone)
    return data
    
#=========================================================================================================================================      
  
def Clean_Time_Posted(data):
    
    # If not in some form similar to '7d ago', set it to 0
    # otherwise extract '7d' 
    def apply_TP(x):
        if re.match(r'^(\d[a-z]\s)',x):
            return re.match(r'^(\d[a-z])\s',x).group(1)
        else:
            return 0   
    
    data['Time_Posted'] = data['Time_Posted'].astype('str')
    data['Time_Posted'] = data['Time_Posted'].apply(apply_TP) 

    # Time should be measured in days  
    def date_transfer(x):
        if x == '0' or 'h' in x:
            return 0
        y = int(x[:-1])
        if 'd' in x:
            return y
        if 'm' in x:
            return y // 2 * 31 + (y - y // 2) * 30
    
    data['Time_Posted'] = data['Time_Posted'].astype('str')
    data['Time_Posted'] = data['Time_Posted'].apply(date_transfer)
    return data

#=========================================================================================================================================        

def Clean_Salary(data):
    
    # Get rid of ','
    data.Salary = data.Salary.str.replace(',', '')
    
    # Replace 'to' with '-' and split 'Salary' by the first '-' intp 'Lo_Salary' and 'Hi_salary'
    data.Salary = data.Salary.str.replace('to', '-', regex=True)
    data[['Lo_Salary', 'Hi_Salary'] ] = data.Salary.str.split('-', n=1, expand=True)
    data.drop('Salary', axis=1, inplace=True)
    
    return data

#=========================================================================================================================================       

# 1.Check if the str contains digits
# 2. If it does, check if it is a percentage
# 3. If it is, leave it there with no change, "8% holiday pay" -> "8% holiday pay"
# 3. If it is a salary, convert it to annual salary 
# 4. Fomular:  x * 8 * 20 * 10
# 5. Only keep numeric info, "NZD75000 per annum" -> "75000.0"; "NZD25" -> "40000.0"
def to_annual(x):
    matchObj = re.match('[^0-9]*(\d+\.*\d*)\s*[a-z]*',x)
    if matchObj:
        if (re.match('.*\d+\.*\d*\s*\%', x)) and ("$" not in x):
            pass
        else:
            x = float(matchObj.group(1))
            # Assume that under 60 is salary per hour
            if x < 60:
                x = x * 8 * 20 * 10
            # For x in between 60 and 250, assume it to be annual salary but without a 'k' for some reason   
            elif x < 250:
                x = x * 1000
            else:
                pass
    return x

#=========================================================================================================================================       

# Find max and min in some column
def Find_Value(data, column):
    df1 = data[data[column].str.contains(r'\d$', regex=True, na=False)][column]
    df1 = df1[df1.values!='0.0']
    df1 = df1.astype('float')
    return df1.max(),df1.min()

#=========================================================================================================================================        

def Clean_LSalary(data):
    data['Lo_Salary'] = data['Lo_Salary'].astype('str')

    # Get rid of white space within a number
    data['Lo_Salary'] = data['Lo_Salary'].str.replace(r'(\d+)\s(\d+)', r'\1\2', regex=True)

    # Replace 'k' with '000'
    data['Lo_Salary'] = data['Lo_Salary'].str.replace(r'k', r'000',case=False)

    data['Lo_Salary'] = data['Lo_Salary'].apply(to_annual)
    
    # Fill all nan's with some artificially made data
    data['Lo_Salary'] = data['Lo_Salary'].astype('str')
    max_LS, min_LS = Find_Value(data,'Lo_Salary')

    # Generate random values in between max and min
    def set_value(x):
        rng = default_rng()
        if x == 'None':
            x = rng.integers(min_LS, max_LS, endpoint=True)
        return x
    
    data['Lo_Salary'] = data['Lo_Salary'].apply(set_value)

    # There is one special case
    # 'Salary' is "Up to $23 p.h. + + 8% Holiday Pay"
    # The info in 'Lo_Salary' would be 'Up '
    # We need set some value to it and make it less than what in 'Hi_Salary' 
    rng = default_rng()
    for index, value in enumerate(data['Lo_Salary'].values):
        if data['Lo_Salary'][index] == 'Up ':
            num = int(rng.integers(20, int(re.match('[^0-9]*(\d+)\.*\d*\s*',data['Hi_Salary'][index]).group(1))))
            data.loc[index, 'Lo_Salary'] = num * 8 * 20 * 10
    
    data['Lo_Salary'] = data['Lo_Salary'].astype('str')
    data['Lo_Salary'] = data['Lo_Salary'].apply(CleanNone)
    return data

#=========================================================================================================================================        

def Clean_HSalary(data):
    data['Hi_Salary'] = data['Hi_Salary'].astype('str')
    data['Hi_Salary'] = data['Hi_Salary'].str.replace(r'(\d+)\s(\d+)', r'\1\2', regex=True)
    data['Hi_Salary'] = data['Hi_Salary'].str.replace('k', '000', case=False)
    data['Hi_Salary'] = data['Hi_Salary'].apply(to_annual)
    data['Hi_Salary'] = data['Hi_Salary'].astype('str')
    max_LS, min_LS = Find_Value(data,'Hi_Salary')
    
    data['Hi_Salary'] = data['Hi_Salary'].astype('str')
    data['Lo_Salary'] = data['Lo_Salary'].astype('str')
    rng = default_rng()

    # less than max, greater than max(Lo_Salary, min)
    for i, v in enumerate(data['Hi_Salary'].values):
        v = v.strip()
        if v == 'None':
            if re.match('^\d+\.*\d*$', data['Lo_Salary'][i]):
                Lo = float(data['Lo_Salary'][i])
                min_LS = max(min_LS, Lo)
                num = rng.integers(max(min_LS,float(data['Lo_Salary'][i])), max_LS)        
            # if it is some str in 'Lo_Salary' like 'Good salary' or 'bonus plus medical insurance', do noting
            elif re.match('.*[a-z]+\s*$', data['Lo_Salary'][i]):
                num = 'Unknown'
                continue
            else:
                num = rng.integers(min_LS, max_LS)
            data.loc[i, 'Hi_Salary'] = num
    
    data['Hi_Salary'] = data['Hi_Salary'].astype('str')
    data['Hi_Salary'] = data['Hi_Salary'].apply(CleanNone)
    return data

#=========================================================================================================================================        

def CleanNone(x):
    x = x.strip()
    if x == 'None':
        x = 'Unknown'
    return x


def Clean(data):
    data = Name_Columns(data)
    data = Set_Orgnisation(data)
    data = Clean_Classification(data)
    data = Clean_Location(data)
    data = Clean_Area(data)
    data = Clean_Time_Posted(data)
    data = Clean_Salary(data)
    data = Clean_LSalary(data)
    data = Clean_HSalary(data)
    data = data[New_Order]
    return data


if __name__ == '__main__':
    data = Read_Data()
    data = Clean(data)
    data.to_excel(File_Save)