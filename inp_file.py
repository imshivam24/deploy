import os
import streamlit as st
import pandas as pd
import numpy as np
from mkm_parameters import *
def inp_file_gen(uploaded_file):
        if uploaded_file:
            try:
                data = pd.read_excel(uploaded_file)
            except Exception as e:
                st.error(f"Error Loading Data: {str(e)}")
                return
            try:
                data1 = pd.read_excel(uploaded_file, sheet_name="Reactions")
                df1 = data1
            except Exception as e:
                st.error(f"Error reading Reactions sheet: {str(e)}")
                return
            
            try:
                data2 = pd.read_excel(uploaded_file, sheet_name="Local Environment")
                df2 = data2
            except Exception as e:
                st.error(f"Error reading Local Environment sheet: {str(e)}")
                return
            
            try:
                data3 = pd.read_excel(uploaded_file, sheet_name="Input-Output Species")
                df3 = data3
            except Exception as e:
                st.error(f"Error reading Input-output sheet: {str(e)}")
                return
        # Extract necessary data from the dataframes
            try:
                global pH_list, V_list, gases, rxn, concentrations, Ea, Eb, P
                pH_list = df2["pH"].tolist()
                V_list = df2["V"][0]
                gases = df3["Species"].tolist()
                concentrations = df3["Concentration"].tolist()
                rxn = df1["Reactions"]
                Ea = df1["G_f"]
                Eb = df1["G_b"]
                P = df2["Pressure"][0]
                st.write("Parameters extracted successfully!")
            except Exception as e:
                st.error(f"Error extracting parameters: {str(e)}")
                return     

            try:
                global adsorbates, activity, Reactant1, Reactant2, Reactant3, Product1, Product2, Product3
                Reactant1 = []
                Reactant2 = []
                Reactant3 = []
                Product1 = []
                Product2 = []
                Product3 = []
                adsorbates = []

                for i in range(len(rxn)):
                    Reactant1.append("{" + rxn[i].split("→")[0].split("+")[0].strip() + "}")
                    if len(rxn[i].split("→")[0].split("+")) == 3:
                        Reactant2.append("{" + rxn[i].split("→")[0].split("+")[1].strip() + "}")
                        Reactant3.append("{" + rxn[i].split("→")[0].split("+")[2].strip() + "}")
                    elif len(rxn[i].split("→")[0].split("+")) == 2:
                        Reactant2.append("{" + rxn[i].split("→")[0].split("+")[1].strip() + "}")
                        Reactant3.append("")
                    else:
                        Reactant2.append("")
                        Reactant3.append("")

                    Product1.append("{" + rxn[i].split("→")[1].split("+")[0].strip() + "}")
                    if len(rxn[i].split("→")[1].split("+")) == 3:
                        Product2.append("{" + rxn[i].split("→")[1].split("+")[1].strip() + "}")
                        Product3.append("{" + rxn[i].split("→")[1].split("+")[2].strip() + "}")
                    elif len(rxn[i].split("→")[1].split("+")) == 2:
                        Product2.append("{" + rxn[i].split("→")[1].split("+")[1].strip() + "}")
                        Product3.append("")
                    else:
                        Product2.append("")
                        Product3.append("")

                for index in Reactant1:
                    if "*" in index and index.strip("{").strip("}") not in adsorbates:
                        adsorbates.append(index.strip("{").strip("}"))

                for index in Reactant2:
                    if "*" in index and index.strip("{").strip("}") not in adsorbates:
                        adsorbates.append(index.strip("{").strip("}"))

                for index in Product1:
                    if "*" in index and index.strip("{").strip("}") not in adsorbates:
                        adsorbates.append(index.strip("{").strip("}"))

                for index in Product2:
                    if "*" in index and index.strip("{").strip("}") not in adsorbates:
                        adsorbates.append(index.strip("{").strip("}"))

                # Check if '*' exists in the list before removing it
                if "*" in adsorbates:
                    adsorbates.remove("*")

                activity = np.zeros(len(adsorbates))

            except Exception as e:
                st.error(f"Error generating input file: {str(e)}")
    

        import os
# Now use these values directly in the directory creation and file path logic
        parent_folder = os.path.join(os.getcwd(), f"single_run")

        # Create directories if they do not exist
        if not os.path.exists(parent_folder):
            os.makedirs(parent_folder)

        input_file_path = os.path.join(parent_folder, "input_file.mkm")

        # Open the file for writing
        with open(input_file_path, 'w') as inp_file:
            inp_file.write('&compounds\n\n')
            inp_file.write("#gas-phase compounds\n\n#Name; isSite; concentration\n\n")
            for compound,concentration in zip(gases,concentrations):
                inp_file.write("{:<15}; 0; {}\n".format(compound,concentration))

            inp_file.write("\n\n#adsorbates\n\n#Name; isSite; activity\n\n")   
            for compound,concentration in zip(adsorbates,activity):
                inp_file.write("{:<15}; 1; {}\n".format(compound,concentration))

            inp_file.write("\n#free sites on the surface \n\n")
            inp_file.write("#Name; isSite; activity\n\n")   
            inp_file.write("*; 1; {}\n\n".format(1.0))    

            inp_file.write('&reactions\n\n')
            pre_exp=6.21e12
            for j in range(len(rxn)):
                if Reactant3[j]!="":
                    line = "AR; {:<15} + {:<15} + {:<5} => {:<15}{:<15};{:<10.2e} ;  {:<10.2e} ;  {:<10} ;  {:<10} \n".format(Reactant1[j],Reactant2[j],Reactant3[j],Product1[j],Product2[j],pre_exp, pre_exp, Ea[j],Eb[j] )   
                elif Product3[j]!="":
                    line = "AR; {:<15} + {:<14}  => {:<10} + {:<15} + {:<7};{:<10.2e} ;  {:<10.2e} ;  {:<10} ;  {:<10} \n".format(Reactant1[j],Reactant2[j],Product1[j],Product2[j],Product3[j],pre_exp, pre_exp, Ea[j],Eb[j] )     
                elif  Reactant2[j]!="" and Product2[j]!="":
                    line = "AR; {:<15} + {:<15} => {:<15} + {:<20};{:<10.2e} ;  {:<10.2e} ;  {:<10} ;  {:<10} \n".format(Reactant1[j],Reactant2[j],Product1[j],Product2[j],pre_exp, pre_exp, Ea[j],Eb[j] )
                elif  Reactant2[j]=="" and Product2[j]!="":
                    line = "AR; {:<15} {:<17} => {:<15} + {:<20};{:<10.2e} ;  {:<10.2e} ;  {:<10} ;  {:<10} \n".format(Reactant1[j],"",Product1[j],Product2[j],pre_exp, pre_exp, Ea[j],Eb[j] )
                elif Reactant2[j]!="" and Product2[j]=="":
                    line = "AR; {:<15} + {:<15} => {:<15}{:<23};{:<10.2e} ;  {:<10.2e} ;  {:<10} ;  {:<10} \n".format(Reactant1[j],Reactant2[j],Product1[j],"",pre_exp, pre_exp, Ea[j],Eb[j] )
                elif Reactant2[j]=="" and Product2[j]=="":
                    line = "AR; {:<15} {:<17} => {:<15}{:<23};{:<10.2e} ;  {:<10.2e} ;  {:<10} ;  {:<10} \n".format(Reactant1[j],"",Product1[j],"",pre_exp, pre_exp, Ea[j],Eb[j] )
                
                inp_file.write(line)    
            inp_file.write("\n\n&settings\nTYPE = SEQUENCERUN\nPRESSURE = {}".format(P))
            inp_file.write("\nPOTAXIS=1\nDEBUG=0\nNETWORK_RATES=1\nNETWORK_FLUX=1\nUSETIMESTAMP=0")
            inp_file.write('\n\n&runs\n')
            inp_file.write("# Temp; Potential;Time;AbsTol;RelTol\n")
            line2 = "{:<5};{:<5};{:<5.2e};{:<5};{:<5}".format(Temp,V_list,Time,Abstol,Reltol)
            inp_file.write(line2)   
            inp_file.close()  
        st.write(f"Input file successfully created at {os.getcwd()}")

        return input_file_path


