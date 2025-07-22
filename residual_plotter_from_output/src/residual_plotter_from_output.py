# 
# This Python script analyzes and visualizes convergence 
# data from an ANSYS solver output file (solve.out). 
# It extracts time-dependent convergence values for 
# forces (FORCE), displacements (DISP), and moments (MOMENT) 
# and plots them in logarithmic-scale diagrams. 
# Additionally, it displays the corresponding convergence 
# criteria to evaluate solution quality.
#

import re
import matplotlib.pyplot as plt
import numpy as np

# Path to the ANSYS output file
file_path = 'solve.out'

# Reading the file content
with open(file_path, 'r') as file:
    content = file.readlines()

# Patterns to extract time and convergence values
time_pattern = re.compile(r'\*\*\* TIME\s*=\s*([\d.E+-]+)')
force_pattern = re.compile(r'FORCE CONVERGENCE VALUE\s*=\s*([\d.E+-]+)\s*CRITERION\s*=\s*([\d.E+-]+)', re.IGNORECASE)
disp_pattern = re.compile(r'DISP CONVERGENCE VALUE\s*=\s*([\d.E+-]+)\s*CRITERION\s*=\s*([\d.E+-]+)', re.IGNORECASE)
moment_pattern = re.compile(r'MOMENT CONVERGENCE VALUE\s*=\s*([\d.E+-]+)\s*CRITERION\s*=\s*([\d.E+-]+)', re.IGNORECASE)

# Data structures to hold extracted data
time_values = []
force_values = []
force_criteria = []
disp_values = []
disp_criteria = []
moment_values = []
moment_criteria = []

curent_time = None

# Parsing the file to extract time and convergence values
for line in content:
    time_match = time_pattern.search(line)
    if time_match:
        current_time = float(time_match.group(1))
        time_values.append(current_time)
        
    if current_time is not None:
        force_match = force_pattern.search(line)
        if force_match:
            force_values.append((current_time, float(force_match.group(1))))
            force_criteria.append((current_time, float(force_match.group(2))))
        
        disp_match = disp_pattern.search(line)
        if disp_match:
            disp_values.append((current_time, float(disp_match.group(1))))
            disp_criteria.append((current_time, float(disp_match.group(2))))
        
        moment_match = moment_pattern.search(line)
        if moment_match:
            moment_values.append((current_time, float(moment_match.group(1))))
            moment_criteria.append((current_time, float(moment_match.group(2))))


# Extracting separate lists for time, values, and criteria
force_times, force_vals = zip(*force_values) if force_values else ([], [])
_, force_criteria_vals = zip(*force_criteria) if force_criteria else ([], [])
disp_times, disp_vals = zip(*disp_values) if disp_values else ([], [])
_, disp_criteria_vals = zip(*disp_criteria) if disp_criteria else ([], [])
moment_times, moment_vals = zip(*moment_values) if moment_values else ([], [])
_, moment_criteria_vals = zip(*moment_criteria) if moment_criteria else ([], [])

# Creating subplots for force, displacement, and moment convergence values and criteria
fig, axs = plt.subplots(3, 1, figsize=(12, 15), sharex=False)

# Plotting force convergence values and criteria
axs[0].plot(force_times, force_vals, label='Force Convergence Value')
axs[0].plot(force_times, force_criteria_vals, label='Criterion')
axs[0].set_ylabel('Force [N]')
axs[0].set_title('Force Convergence Values and Criteria')


# Plotting displacement convergence values and criteria
axs[1].plot(disp_times, disp_vals, label='Displacement Convergence Value')
axs[1].plot(disp_times, disp_criteria_vals, label='Criterion')
axs[1].set_ylabel('Displacement [mm]')
axs[1].set_title('Displacement Convergence Values and Criteria')


# Plotting moment convergence values and criteria
axs[2].plot(moment_times, moment_vals, label='Moment Convergence Value')
axs[2].plot(moment_times, moment_criteria_vals, label='Criterion')
axs[2].set_ylabel('Moment [Nmm]')
axs[2].set_title('Moment Convergence Values and Criteria')




xticks = np.linspace(min(time_values), max(time_values), num=16)
rounded_xticks = [round(tick, 3) for tick in xticks]

for ax in axs:
    ax.set_xlabel('Time [s]')
    ax.set_yscale('log')
    ax.legend(loc = 'best')
    ax.grid(True, which="both", ls="--")
    ax.set_xticks(xticks)
    ax.set_xticklabels(rounded_xticks)

# Displaying the plots
plt.tight_layout()
plt.show()
