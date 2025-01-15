import streamlit as st
import pandas as pd
import os
import subprocess

from inp_file import *
from mkm_parameters import *
import shutil
from io import StringIO

# Function to run the executable and generate the required outputs
def run_executable(input_file):
    # Path to the Windows executable
    executable_path = r"/mount/src/deploy/bin/mkmcxx.exe"  # Adjust the path as needed

    # Display debugging information in Streamlit
    st.write("Executable Path:", executable_path)
    st.write("Current Working Directory:", os.getcwd())
    st.write("Contents of Directory:", os.listdir(r"/mount/src/deploy/bin"))

    # Check if the executable exists at the given path
    if os.path.exists(executable_path):
        st.write(f"Executable found at: {executable_path}")
        try:
            # Use Wine to run the Windows executable on Linux
            result = subprocess.run(
                ['wine', executable_path, '-i', input_file],
                capture_output=True,
                text=True
            )
            
            # Display the output and error in Streamlit
            st.write("Solver Output (stdout):")
            st.text(result.stdout)

            if result.stderr:
                st.write("Solver Error Output (stderr):")
                st.text(result.stderr)
            
            # Return success or failure
            if result.returncode == 0:
                return "Solver ran successfully!", True
            else:
                return f"Error running solver: {result.stderr}", False

        except Exception as e:
            return f"Error executing command: {str(e)}", False
    else:
        return "Executable not found at the given path.", False
def get_val (cov_path):   
    cov_file = open(cov_path)
    cov_val = []
    cov_lines = cov_file.readlines()
    adsorbate_keys = cov_lines[0].strip().split()
    cov_dat_dict = {}
    for key in adsorbate_keys:
        cov_dat_dict[key] = []
    for line in cov_lines[1:]:
        vals = line.strip().split()
        vals = list(map(lambda x: float(x), vals))
        c=0
        for key in adsorbate_keys:
            cov_dat_dict[key].append(vals[c])
            c += 1
    cov_file.close()
    return cov_dat_dict

def coverage():
    coverage_file_path = "run/range/coverage.dat"
    
    if os.path.exists(coverage_file_path):
        covs = get_val(coverage_file_path)
        covs_relevant ={}

        for key in covs.keys():
            if '*' in key:
                covs_relevant[key] = covs[key]
        #Read the data from the coverage.dat file
        # list coverrage dictionary
        covs_relevant_df = pd.DataFrame(covs_relevant)
        covs_relevant_df = covs_relevant_df.T
        covs_relevant_df = covs_relevant_df.reset_index()
        covs_relevant_df.columns = ['Adsorbates', 'Coverage']
        st.write("Coverage Data:")
        st.write(covs_relevant_df)
    else:
        st.error("coverage.dat file not found in the expected directory.")  
def coverage_V(pH, V):
    coverage_file_path = f"run/range/pH_{pH}/V_{V}/coverage.dat"
    if os.path.exists(coverage_file_path):
        try:
            covs = get_val(coverage_file_path)
            covs_relevant = {key: covs[key] for key in covs.keys() if '*' in key}
            
            # Collect data in DataFrame for easier manipulation
            covs_relevant_df = pd.DataFrame(covs_relevant).T.reset_index()
            covs_relevant_df.columns = ['Adsorbates', 'Coverage']
            return covs_relevant_df
        except Exception as e:
            st.error(f"Error processing coverage data for pH={pH}, V={V}: {e}")
            return None
    else:
        st.error(f"coverage.dat file not found for pH={pH}, V={V}.")
        return None

def plot_coverage_data(pH_list, V_list):
    all_data = []  # Store coverage data for all pH, V combinations

    # Loop over all pH and V combinations and gather the coverage data
    for pH in pH_list:
        for V in V_list:
            coverage_data = coverage_V(pH, V)
            if coverage_data is not None:
                # Adding the pH and V values to the DataFrame
                coverage_data['pH'] = pH
                coverage_data['V'] = V
                all_data.append(coverage_data)

    # Concatenate all data into a single DataFrame
    if all_data:
        df_all_coverage = pd.concat(all_data, ignore_index=True)

        # Plot coverage for each pH and V combination
        fig, ax = plt.subplots(figsize=(10, 6))

        # Plot each adsorbate's coverage as a function of pH and V
        for adsorbate in df_all_coverage['Adsorbates'].unique():
            subset = df_all_coverage[df_all_coverage['Adsorbates'] == adsorbate]
            ax.plot(subset['V'], subset['Coverage'], label=adsorbate)

        ax.set_xlabel('Potential (V)')
        ax.set_ylabel('Coverage')
        ax.set_title('Coverage vs. Potential (V) for each Adsorbate')
        ax.legend(title='Adsorbates')
        st.pyplot(fig)
    else:
        st.error("No coverage data available for the given pH and V combinations.")