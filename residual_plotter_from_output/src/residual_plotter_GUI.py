import re
import matplotlib.pyplot as plt
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class ConvergenceGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ANSYS Solve Out Viewer")
        self.file_path = None

        # Buttons
        self.select_button = tk.Button(root, text="Select solve.out", command=self.select_file)
        self.select_button.pack(pady=5)

        self.update_button = tk.Button(root, text="Update Plot", command=self.plot_data)
        self.update_button.pack(pady=5)

        # Placeholder for matplotlib figure
        self.figure = None
        self.canvas = None

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select solve.out file",
            filetypes=[("Text files", "*.out"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.plot_data()

    def parse_file(self):
        if not self.file_path:
            messagebox.showwarning("No file selected", "Please select a solve.out file first.")
            return None

        with open(self.file_path, 'r') as file:
            content = file.readlines()

        # Patterns
        time_pattern = re.compile(r'\*\*\* TIME\s*=\s*([\d.E+-]+)')
        force_pattern = re.compile(r'FORCE CONVERGENCE VALUE\s*=\s*([\d.E+-]+)\s*CRITERION\s*=\s*([\d.E+-]+)', re.IGNORECASE)
        disp_pattern = re.compile(r'DISP CONVERGENCE VALUE\s*=\s*([\d.E+-]+)\s*CRITERION\s*=\s*([\d.E+-]+)', re.IGNORECASE)
        moment_pattern = re.compile(r'MOMENT CONVERGENCE VALUE\s*=\s*([\d.E+-]+)\s*CRITERION\s*=\s*([\d.E+-]+)', re.IGNORECASE)

        # Data
        time_values = []
        force_values, force_criteria = [], []
        disp_values, disp_criteria = [], []
        moment_values, moment_criteria = [], []

        current_time = None
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

        return time_values, force_values, force_criteria, disp_values, disp_criteria, moment_values, moment_criteria

    def plot_data(self):
        parsed = self.parse_file()
        if not parsed:
            return
        time_values, force_values, force_criteria, disp_values, disp_criteria, moment_values, moment_criteria = parsed

        # Extract lists
        force_times, force_vals = zip(*force_values) if force_values else ([], [])
        _, force_criteria_vals = zip(*force_criteria) if force_criteria else ([], [])
        disp_times, disp_vals = zip(*disp_values) if disp_values else ([], [])
        _, disp_criteria_vals = zip(*disp_criteria) if disp_criteria else ([], [])
        moment_times, moment_vals = zip(*moment_values) if moment_values else ([], [])
        _, moment_criteria_vals = zip(*moment_criteria) if moment_criteria else ([], [])

        # Clear previous figure if exists
        if self.figure:
            self.figure.clf()

        self.figure, axs = plt.subplots(3, 1, figsize=(10, 12), sharex=False)

        # Force
        if force_vals:
            axs[0].plot(force_times, force_vals, label='Force Convergence Value')
            axs[0].plot(force_times, force_criteria_vals, label='Criterion')
        else:
            axs[0].text(0.5, 0.5, 'No data', ha='center', va='center')
        axs[0].set_ylabel('Force [N]')
        axs[0].set_title('Force Convergence Values and Criteria')

        # Displacement
        if disp_vals:
            axs[1].plot(disp_times, disp_vals, label='Displacement Convergence Value')
            axs[1].plot(disp_times, disp_criteria_vals, label='Criterion')
        else:
            axs[1].text(0.5, 0.5, 'No data', ha='center', va='center')
        axs[1].set_ylabel('Displacement [mm]')
        axs[1].set_title('Displacement Convergence Values and Criteria')

        # Moment
        if moment_vals:
            axs[2].plot(moment_times, moment_vals, label='Moment Convergence Value')
            axs[2].plot(moment_times, moment_criteria_vals, label='Criterion')
        else:
            axs[2].text(0.5, 0.5, 'No data', ha='center', va='center')
        axs[2].set_ylabel('Moment [Nmm]')
        axs[2].set_title('Moment Convergence Values and Criteria')

        # X-axis
        if time_values:
            xticks = np.linspace(min(time_values), max(time_values), num=16)
            rounded_xticks = [round(tick, 3) for tick in xticks]
            for ax in axs:
                ax.set_xticks(xticks)
                ax.set_xticklabels(rounded_xticks)

        # Log scale only for positive values
        for ax, vals in zip(axs, [force_vals, disp_vals, moment_vals]):
            if vals and any(v > 0 for v in vals):
                ax.set_yscale('log')
            ax.set_xlabel('Time [s]')
            ax.legend(loc='best')
            ax.grid(True, which="both", ls="--")

        plt.tight_layout()

        if self.canvas:
            self.canvas.get_tk_widget().pack_forget()
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.root)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)


if __name__ == "__main__":
    root = tk.Tk()
    gui = ConvergenceGUI(root)
    root.mainloop()
