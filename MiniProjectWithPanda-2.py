import openpyxl
import csv
import sys
import os
import pandas as pd
cmd_arguments = sys.argv # 2 arguments expected 1- input director, 2-Output directory
#print(f"**\n\nOs.path = {os.fspath}")
script_dir = os.path.dirname(os.path.abspath(__file__))
print(f"Arguments >>>>>>>>{len(cmd_arguments)}")

if len(cmd_arguments)<2: # if no parameters were supplied - programs directory is assumed for both input and outpur
    xlfile_path=script_dir + (r"\rule_engine.xlsx")
    output_path=script_dir
else:
    if len(cmd_arguments)<3: # if only one directory is supplied ie input directory, same directory assumed for output directory
        xlfile_path=cmd_arguments[1] +(r"\rule_engine.xlsx")
        print(f" input directory = {xlfile_path}")
        output_path = cmd_arguments[1]
        print(f" outpur directory = {output_path}")
    else: # botH directories supplied 
        xlfile_path=cmd_arguments[1] +(r"\rule_engine.xlsx")
        output_path = cmd_arguments[2]


print(f"\n\nxl file = {xlfile_path}")
print(f"\noutput path  = {output_path}")


if not os.path.exists(xlfile_path):
   raise Exception(f"fine name {xlfile_path} does not exist")
if not os.path.isdir(output_path):
    raise Exception(f"Output directory {output_path} does not exist")

csvfile_path_TEMP_CP_1 = output_path + r"\TEMP_CP_1.csv" #r"C:\Users\riyaz\Documents\Projects\Python\NASCOM\CodeFilesBatch2\Class\Projects\Project1\TEMP_CP_1.csv"
csvfile_path_TEMP_AC_1 = output_path + r"\TEMP_AC_1.csv" #r"C:\Users\riyaz\Documents\Projects\Python\NASCOM\CodeFilesBatch2\Class\Projects\Project1\TEMP_AC_1.csv"
csvfile_path_TEMP_CS_1 = output_path + r"\TEMP_CS_1.csv" #r"C:\Users\riyaz\Documents\Projects\Python\NASCOM\CodeFilesBatch2\Class\Projects\Project1\TEMP_CS_1.csv"
csvfile_path_TEMP_CS_2 = output_path + r"\TEMP_CS_2.csv"#r"C:\Users\riyaz\Documents\Projects\Python\NASCOM\CodeFilesBatch2\Class\Projects\Project1\TEMP_CS_2.csv"


df  = pd.read_excel(xlfile_path, sheet_name='Test Data Set')
df["Temperature"] = df['Temperature Condition'].str[:-1].astype(float) # add a column by extracting the numeric value of temperature condition

#Check if material category 'API' OR ' MARKETED' has the temperature condition maintained
filtered_df_TEMP_CP_1 = df[
      (df['Temperature Condition'].notna())  &
    (df['Material Category'].isin(['API', 'MARKETED']))
]
filtered_df_TEMP_CP_1.pop("Temperature")
#If  temperature condition is maintained and  check if the value is in the range of -10 to -2 C for material category 'API', 
# and check if the value is in the range of 2 to 8 C for material category 'MARKETED'
filtered_df_TEMP_AC_1 = df[
      (df['Temperature Condition'].notna()) &
        (
         ( (df['Material Category'].isin(['API'])) & 
           (df['Temperature'] <= -2)  &
           (df['Temperature'] >= -10) 
          )
                  |
         ( 
             (df['Material Category']=='MARKETED') & 
             (df['Temperature'] <= 8)  &
           (df['Temperature'] >= 2)  
          )
        )

  ]
filtered_df_TEMP_AC_1.pop("Temperature")
#If  temperature condition is maintained and check if the value is  < 2 C if shipping condition is 'Truck'
filtered_df_TEMP_CS_1 = df[
      (df['Temperature Condition'].notna()) &
        (
         ( (df["Shipping Condition"] =='Truck') & 
           (df['Temperature'] <= 2)  
          )
                  |
         ( 
             (df["Shipping Condition"] =='Truck')
          )
        )     
     ]
filtered_df_TEMP_CS_1.pop("Temperature")

# If temperature condition is < = 0 C, storage condition should be 'COLD'

filtered_df_TEMP_CS_2 = df[
      (df['Temperature Condition'].notna()) &
        (
         ( (df["Storage Condition"] == "COLD") & 
           (df['Temperature'] <= 0)  
          )
                  |
         ( 
             (df['Temperature'] >0)
          )
        )
    ]
filtered_df_TEMP_CS_2.pop("Temperature")
#filtered_df_TEMP_CP_1.to_excel(csvfile_path_TEMP_CP_1, index=False)
filtered_df_TEMP_CP_1.to_csv(csvfile_path_TEMP_CP_1, index=False)
filtered_df_TEMP_AC_1.to_csv(csvfile_path_TEMP_AC_1, index=False)
filtered_df_TEMP_CS_1.to_csv(csvfile_path_TEMP_CS_1, index=False)
filtered_df_TEMP_CS_2.to_csv(csvfile_path_TEMP_CS_2, index=False)
print("\n****** Completed ******")


