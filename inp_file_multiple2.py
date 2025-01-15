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
        # Replace Excel-style sheet references with actual values
        import re

        def replace_reference(match):
            ref = match.group(1)  # Extract reference inside quotes
            sheet, cell = ref.split('!')
            sheet = sheet.strip("'")  # Remove surrounding quotes
            row, col = cell[1:], cell[:1]  # Extract row and column
            row = int(row) - 1  # Convert to 0-based index
            col_index = ord(col.upper()) - ord('A')  # Convert column to index
            return str(sheet_data[sheet].iloc[row, col_index])  # Lookup value

        # Find all references like 'SheetName'!Cell
        pattern = r"'([^']+)'!([A-Z]+\d+)"
        formula = re.sub(pattern, replace_reference, formula)

        # Evaluate the modified formula using sympy
        expr = sympify(formula)
        result = expr.evalf(subs=context)
        return result

    except Exception as e:
        st.error(f"Error evaluating formula '{formula}': {e}")
        return np.nan


def read_and_compute(file_name, sheet_name, column_name, dependencies):
    """
    Reads formulas from an Excel sheet, computes their values, and returns a DataFrame.

    Args:
    file_name (str): Path to the Excel file.
    sheet_name (str): Name of the sheet to process.
    column_name (str): Column containing formulas.
    dependencies (dict): Mapping of column names to their computed values.

    Returns:
    pd.Series: Series with computed values for the specified column.
    """
    try:
        # Load the workbook and sheet
        workbook = load_workbook(file_name)
        sheet = workbook[sheet_name]

        # Extract column names and values
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append(row)

        df = pd.DataFrame(data[1:], columns=data[0])  # First row as headers

        # Compute values for the formula column
        if column_name not in df.columns:
            raise ValueError(f"Column '{column_name}' not found in the sheet.")

        computed_values = []
        for formula in df[column_name]:
            if isinstance(formula, str) and '=' in formula:  # Formula detected
                # Compute formula using dependencies
                computed_values.append(evaluate_excel_formula(formula.strip('='), dependencies))
            else:  # Not a formula
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
            # Read required data
            data2 = pd.read_excel(uploaded_file, sheet_name="Local Environment")
            data3 = pd.read_excel(uploaded_file, sheet_name="Input-Output Species")

            # Define dependencies for formula computation
            dependencies = {
                'pH': data2['pH'][0],
                'V': data2['V'][0],
                'Pressure': data2['Pressure'][0],
            }

            # Compute formula columns
            concentrations = read_and_compute(uploaded_file, 'Input-Output Species', 'Input MKMCXX', dependencies)
            Ea = read_and_compute(uploaded_file, 'Reactions', 'G_f', dependencies)
            Eb = read_and_compute(uploaded_file, 'Reactions', 'G_b', dependencies)
            gases = data3["Species"].tolist()
            rxn = pd.read_excel(uploaded_file, sheet_name="Reactions")["Reactions"]

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
