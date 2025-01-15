import os
import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from sympy import sympify, symbols
from mkm_parameters import *

def evaluate_excel_formula(formula, context, sheet_data):
    """
    Evaluates an Excel-style formula by replacing references with actual values.
    
    Args:
    formula (str): The formula to evaluate.
    context (dict): A dictionary for variable substitution.
    sheet_data (dict): Dictionary of sheet names to DataFrame mappings.
    
    Returns:
    float: The evaluated result.
    """
    try:
        import re

        def replace_reference(match):
            # Extract reference like 'SheetName'!Cell
            ref = match.group(1)
            sheet, cell = ref.split('!')
            sheet = sheet.strip("'")  # Remove quotes
            col, row = re.match(r"([A-Z]+)(\d+)", cell).groups()  # Extract column and row
            row_idx = int(row) - 1  # Convert to 0-based index
            col_idx = ord(col.upper()) - ord('A')  # Convert column letter to index
            return str(sheet_data[sheet].iloc[row_idx, col_idx])  # Lookup value

        # Replace Excel-style references with actual values
        pattern = r"'([^']+)'!([A-Z]+\d+)"
        formula = re.sub(pattern, replace_reference, formula)

        # Evaluate the modified formula using sympy
        expr = sympify(formula)
        result = expr.evalf(subs=context)
        return result

    except Exception as e:
        st.error(f"Error evaluating formula '{formula}': {e}")
        return np.nan

def read_and_compute(file_name, sheet_name, column_name, sheet_data):
    """
    Reads formulas from an Excel sheet, computes their values, and returns a Series.

    Args:
    file_name (str): Path to the Excel file.
    sheet_name (str): Name of the sheet to process.
    column_name (str): Column containing formulas.
    sheet_data (dict): Mapping of sheet names to DataFrames.

    Returns:
    pd.Series: Series with computed values for the specified column.
    """
    try:
        # Load the workbook and extract the desired sheet as a DataFrame
        data = pd.read_excel(file_name, sheet_name=sheet_name)

        # Compute values for the specified column
        if column_name not in data.columns:
            raise ValueError(f"Column '{column_name}' not found in the sheet.")

        computed_values = []
        for formula in data[column_name]:
            if isinstance(formula, str) and '=' in formula:  # Formula detected
                computed_values.append(evaluate_excel_formula(formula.strip('='), {}, sheet_data))
            else:  # Not a formula, use value as is
                computed_values.append(formula)

        return pd.Series(computed_values, name=column_name)

    except Exception as e:
        st.error(f"Error reading and computing formulas: {e}")
        return pd.Series(dtype=float)

def inp_file_gen_multiple(uploaded_file, children_folder):
    """
    Generates an input file based on Excel file data, evaluating formulas manually.
    """
    if uploaded_file:
        try:
            # Read all required sheets into a dictionary for formula evaluation
            sheet_data = {
                'Local Environment': pd.read_excel(uploaded_file, sheet_name='Local Environment'),
                'Input-Output Species': pd.read_excel(uploaded_file, sheet_name='Input-Output Species'),
                'Reactions': pd.read_excel(uploaded_file, sheet_name='Reactions')
            }

            # Extract parameters
            data2 = sheet_data['Local Environment']
            data3 = sheet_data['Input-Output Species']

            dependencies = {
                'pH': data2['pH'].iloc[0],
                'V': data2['V'].iloc[0],
                'Pressure': data2['Pressure'].iloc[0],
            }

            # Compute formula columns
            concentrations = read_and_compute(uploaded_file, 'Input-Output Species', 'Input MKMCXX', sheet_data)
            Ea = read_and_compute(uploaded_file, 'Reactions', 'G_f', sheet_data)
            Eb = read_and_compute(uploaded_file, 'Reactions', 'G_b', sheet_data)
            gases = data3["Species"].tolist()
            rxn = sheet_data['Reactions']["Reactions"]

            st.write("Parameters extracted and formulas computed successfully!")
        except Exception as e:
            st.error(f"Error extracting parameters or computing formulas: {str(e)}")
            return

        try:
            # Process reactions
            Reactant1, Reactant2, Reactant3 = [], [], []
            Product1, Product2, Product3 = [], [], []
            adsorbates = []

            for reaction in rxn:
                reactants, products = reaction.split("→")
                reactants = [r.strip() for r in reactants.split("+")]
                products = [p.strip() for p in products.split("+")]

                Reactant1.append(f"{{{reactants[0]}}}")
                Reactant2.append(f"{{{reactants[1]}}}" if len(reactants) > 1 else "")
                Reactant3.append(f"{{{reactants[2]}}}" if len(reactants) > 2 else "")

                Product1.append(f"{{{products[0]}}}")
                Product2.append(f"{{{products[1]}}}" if len(products) > 1 else "")
                Product3.append(f"{{{products[2]}}}" if len(products) > 2 else "")

                for item in reactants + products:
                    if "*" in item and item not in adsorbates:
                        adsorbates.append(item)

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

                inp_file.write("\n\n&settings\nTYPE = SEQUENCERUN\nPRESSURE = {}\nPOTAXIS=1\nDEBUG=0\n".format(dependencies['Pressure']))
                inp_file.write("NETWORK_RATES=1\nNETWORK_FLUX=1\nUSETIMESTAMP=0\n\n&runs\n# Temp; Potential;Time;AbsTol;RelTol\n")
                inp_file.write("{:<5};{:<5};{:<5.2e};{:<5};{:<5}\n".format(Temp, dependencies['V'], Time, Abstol, Reltol))
            st.write(f"Input file successfully created at {inp_file_path}")
        except Exception as e:
            st.error(f"Error writing input file: {str(e)}")
