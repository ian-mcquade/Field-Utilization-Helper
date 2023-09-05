import os
import pandas as pd
from tkinter import Tk, filedialog, Button, Label
from tkinter import ttk 


def calculate_field_utilization(filenames, output_dir):
    total_files = len(filenames)
    for i, filename in enumerate(filenames):
        # Determine file type and load the data accordingly
        if filename.endswith('.xlsx'):
            df = pd.read_excel(filename)
        elif filename.endswith('.csv'):
            df = pd.read_csv(filename)
        else:
            print(f"Unsupported file type: {filename}")
            continue

        # Save the original headers
        headers = df.columns.tolist()

        # Clean the data - remove leading/trailing spaces and non-breaking spaces
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        df = df.replace('\xa0', '', regex=True) #\xa0 is a non-breaking space in Unicode

        # Calculate field utilization
        total_records = len(df)
        utilization_percentages = [(df[column].count() / total_records) * 100 for column in df.columns]
        utilization_percentages = [round(x, 2) for x in utilization_percentages] # Round to two decimal places

        # Insert utilization percentages as a new row
        df.loc[-1] = utilization_percentages
        df.index = df.index + 1  # Shift index
        df.sort_index(inplace=True) 

        # Insert the original headers as the column names directly
        df.columns = headers

        # Get the original file name (without extension)
        base_filename = os.path.basename(filename)
        base_filename_without_ext = os.path.splitext(base_filename)[0]

        # Write the resulting dataframe to a new file in the selected output directory
        if filename.endswith('.xlsx'):
            df.to_excel(os.path.join(output_dir, f"{base_filename_without_ext}_with_utilization.xlsx"), index=False)
        else: # .csv
            df.to_csv(os.path.join(output_dir, f"{base_filename_without_ext}_with_utilization.csv"), index=False)

        # Update the progress bar
        progress['value'] = (i+1) / total_files * 100
        root.update_idletasks()  # Force redraw of the progress bar

from tkinter import messagebox

def select_files():
    root = Tk()
    root.withdraw()  # Hide the main window
    root.call('wm', 'attributes', '.', '-topmost', True)  # Bring the file selection dialog on top

    # Select files
    filenames = filedialog.askopenfilenames()  # This line will open file explorer to select one or multiple files

    if filenames:
        # Show a message
        messagebox.showinfo("Select Output Directory", "Please select the output directory for the new file(s).")
        
        # Select output directory
        output_dir = filedialog.askdirectory()

        if output_dir:
            calculate_field_utilization(filenames, output_dir)
            result_label.config(text=f"Processed {len(filenames)} file(s) successfully! New files are saved in {output_dir}")
        else:
            result_label.config(text="No output directory selected.")
        
        progress['value'] = 0  # Reset progress bar
    else:
        result_label.config(text="No files selected.")



root = Tk()
root.title("Field Utilization Helper")

# Set window size
root.geometry("600x200")

label = Label(root, text="Click the button below to select Excel or CSV files:")
label.pack()

browse_button = Button(root, text="Select Files", command=select_files)
browse_button.pack()

# Create a progress bar
progress = ttk.Progressbar(root, length=100, mode='determinate')
progress.pack()

result_label = Label(root, text="")
result_label.pack()

root.mainloop()

