import os
import subprocess
import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from mkm_parameters import *

def recalculate_excel(file_path):
    """
    Recalculates formulas in an Excel file using LibreOffice CLI and returns the path to the recalculated file.
    
    Args:
    file_path (str): Path to the original Excel file.
    
    Returns:
    str: Path to the recalculated Excel file.
    """
    try:
        # Define output file path
        recalculated_file_path = file_path.replace(".xlsx", "_recalculated.xlsx")
        
        # Use LibreOffice CLI to recalculate formulas
        subprocess.run(
            [
                "libreoffice", "--headless", "--convert-to", "xlsx",
                file_path, "--outdir", os.path.dirname(file_path)
            ],
            check=True
        )
        return recalculated_file_path
    except Exception as e:
        st.error(f"Error recalculating Excel formulas: {str(e)}")
        raise e

def read_formulas(file_name, sheet_name, column_name):
    """
    Reads the formulas in an Excel file, saves the file, and returns the DataFrame with the calculated values.
    
    Args:
    file_name (str): Path to the Excel file.
    sheet_name (str): Name of the sheet to process.
    column_name (str): The column name to focus on in the resulting DataFrame.
    
    Returns:
    pd.DataFrame: DataFrame with the computed values from the specified column.
    """
    try:
        # Load workbook with recalculated formulas
        workbook = load_workbook(file_name, data_only=True)
        sheet = workbook[sheet_name]

        # Read sheet data into a list
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(row)

        # Convert list to DataFrame
        df = pd.DataFrame(data[1:], columns=data[0])  # First row as headers

        # Filter the DataFrame for the specified column
        if column_name not in df.columns:
            raise ValueError(f"Column '{column_name}' not found in the sheet.")
        
        return df[[column_name]]  # Return only the requested column
    
    except Exception as e:
        st.write(f"Error: {e}")
        print(f"Error: {e}")
        return None

def inp_file_gen_multiple(uploaded_file, children_folder):
    """
    Generates an input file based on Excel file data.
    """
    if uploaded_file:
        try:
            # Recalculate formulas in the uploaded file
            recalculated_file = recalculate_excel(uploaded_file)

            # Read necessary sheets from the recalculated Excel file
            data1 = pd.read_excel(recalculated_file, sheet_name="Reactions")
            data2 = pd.read_excel(recalculated_file, sheet_name="Local Environment")
            data3 = pd.read_excel(recalculated_file, sheet_name="Input-Output Species")
        except Exception as e:
            st.error(f"Error reading sheets: {str(e)}")
            return
        
        try:
            # Extract required parameters
            pH_list = data2["pH"].tolist()
            V_list = data2["V"][0]
            gases = data3["Species"].tolist()
            concentrations = read_formulas(recalculated_file, 'Input-Output Species', 'Input MKMCXX')['Input MKMCXX']
            rxn = data1["Reactions"]
            Ea = read_formulas(recalculated_file, 'Reactions', 'G_f')['G_f']
            Eb = read_formulas(recalculated_file, 'Reactions', 'G_b')['G_b']
            P = data2["Pressure"][0]
            st.write("Parameters extracted successfully!")
        except Exception as e:
            st.error(f"Error extracting parameters: {str(e)}")
            return

        try:
            # Process reactions
            Reactant1, Reactant2, Reactant3 = [], [], []
            Product1, Product2, Product3 = [], [], []
            adsorbates = []

            for i in range(len(rxn)):
                # Extract Reactants and Products
                reactants = rxn[i].split("→")[0].split("+")
                products = rxn[i].split("→")[1].split("+")
                
                Reactant1.append(f"{{{reactants[0].strip()}}}")
                Reactant2.append(f"{{{reactants[1].strip()}}}" if len(reactants) > 1 else "")
                Reactant3.append(f"{{{reactants[2].strip()}}}" if len(reactants) > 2 else "")

                Product1.append(f"{{{products[0].strip()}}}")
                Product2.append(f"{{{products[1].strip()}}}" if len(products) > 1 else "")
                Product3.append(f"{{{products[2].strip()}}}" if len(products) > 2 else "")

                # Collect adsorbates
                for r in (Reactant1[-1], Reactant2[-1], Product1[-1], Product2[-1]):
                    if "*" in r and r.strip("{").strip("}") not in adsorbates:
                        adsorbates.append(r.strip("{").strip("}"))
            
            # Remove "*" from adsorbates
            adsorbates = [a for a in adsorbates if a != "*"]
            activity = np.zeros(len(adsorbates))
        except Exception as e:
            st.error(f"Error processing reactions: {str(e)}")
            return

        # Write input file
        inp_file_path = os.path.join(children_folder, 'input_file.mkm')
        try:
            with open(inp_file_path, 'w') as inp_file:
                inp_file.write('&compounds\n\n#gas-phase compounds\n\n#Name; isSite; concentration\n\n')
                for compound, concentration in zip(gases, concentrations):
                    inp_file.write(f"{compound:<15}; 0; {concentration}\n")

                inp_file.write("\n\n#adsorbates\n\n#Name; isSite; activity\n\n")
                for compound, activity_value in zip(adsorbates, activity):
                    inp_file.write(f"{compound:<15}; 1; {activity_value}\n")

                inp_file.write("\n#free sites on the surface \n\n#Name; isSite; activity\n\n")
                inp_file.write("*; 1; 1.0\n\n")

                inp_file.write('&reactions\n\n')
                pre_exp = 6.21e12
                for j in range(len(rxn)):
                    line = f"AR; {Reactant1[j]} + {Reactant2[j]} => {Product1[j]} + {Product2[j]}; {pre_exp:.2e}; {pre_exp:.2e}; {Ea[j]}; {Eb[j]}\n"
                    inp_file.write(line)

                inp_file.write("\n\n&settings\nTYPE = SEQUENCERUN\nPRESSURE = {}\nPOTAXIS=1\nDEBUG=0\n".format(P))
                inp_file.write("NETWORK_RATES=1\nNETWORK_FLUX=1\nUSETIMESTAMP=0\n\n&runs\n# Temp; Potential;Time;AbsTol;RelTol\n")
                inp_file.write("{:<5};{:<5};{:<5.2e};{:<5};{:<5}\n".format(Temp, V_list, Time, Abstol, Reltol))
            st.write(f"Input file successfully created at {inp_file_path}")
        except Exception as e:
            st.error(f"Error writing input file: {str(e)}")
