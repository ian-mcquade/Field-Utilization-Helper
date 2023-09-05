# Field Utilization Helper

## Introduction
Field Utilization Helper is a Python-based tool that allows you to quickly and easily calculate the field utilization percentages for each column in an Excel or CSV file. This information is inserted into a new row in the data and the resulting file is saved in the specified directory.

## Features
- Supports both Excel and CSV files.
- Simple and intuitive GUI for selecting multiple files and output directory.
- Automatically strips leading/trailing spaces and removes non-breaking spaces.
- Efficiently calculates the field utilization for each column and adds it to a new row.

## Installation
To get started, clone this repository to your local machine:
```
git clone https://github.com/ian-mcquade/Field-Utilization-Helper.git
```

## Usage
1. Open the program and click on the "Select Files" button to choose the Excel or CSV files you want to process.
2. You'll then be prompted to select the output directory where the new files will be saved.
3. Click "OK" and the program will process the files, calculating the field utilization for each column.

## Dependencies
This project is built with Python and makes use of the following libraries:
- pandas
- tkinter

You can install these using pip:
```
pip install pandas
```
Note: Tkinter comes pre-installed with Python.

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.
