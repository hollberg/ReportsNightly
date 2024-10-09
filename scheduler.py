r"""
scheduler.py

This script does the following:

Reads the CSV file to get the list of scripts and whether they should be executed.
Iterates over each row in the DataFrame. If the "Execute" column is 'Y', it executes
the script using subprocess.run().
Prints messages to indicate the status of each script execution.
To have this master script run nightly, you can use the Task Scheduler on Windows since
you mentioned using Windows 10 or Server 2019. Here's a brief overview of how to set it up:

Open Task Scheduler.
Create a new task and give it a name.
Under the "Triggers" tab, set it to run daily at the desired time.
Under the "Actions" tab, set the action to "Start a program". For
the program/script field, browse and select your Python executable
(usually found in your Python installation directory,
e.g., C:Python39\python.exe). In the "Add arguments" field,
input the path to your master script.

Configure any other settings as required and save the task.
This setup will ensure your master script runs every night, executing any scripts
you have specified in your CSV file.

"""
import os
import sys

import pandas as pd
import subprocess

# Path to the CSV file that lists the scripts to run
# csv_file = 'scripts_to_run.csv'
# Trying the full absolute path to see if script will then
# run when executing via a *.bat file (to avoid this error:
#    FileNotFoundError: [Errno 2] No such file or directory: 'scripts_to_run.csv'
csv_file = r'C:\Users\mitchell.hollberg\OneDrive - Community Foundation for Greater Atlanta\Documents\py\ReportsNightly\scripts_to_run.csv'


# Capture the current environment variables
env = os.environ.copy()
python_executable_path = sys.executable


def run_scripts_from_csv(csv_path):
    # Load the CSV file
    df = pd.read_csv(csv_path)

    # Loop through the rows in the DataFrame
    for index, row in df.iterrows():
        # Check if the script is marked for execution
        if row['Execute'] == 'Y':
            script_path = row['FilePath']
            print(f"Executing {script_path}...")
            try:
                # Execute the script using subprocess
                subprocess.run([python_executable_path, script_path], env=env, check=True)
                print(f"Successfully executed {script_path}")
            except subprocess.CalledProcessError as e:
                print(f"Failed to execute {script_path}: {e}")


# Call the function to start executing scripts
run_scripts_from_csv(csv_file)
