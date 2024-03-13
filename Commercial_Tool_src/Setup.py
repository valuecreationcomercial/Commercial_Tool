#!/usr/bin/env python
# coding: utf-8
# Downloading and Importing Python Libraries
print("Downloading and Importing Python Libraries")
import subprocess
subprocess.run("python -m pip install pandas", shell=True)
subprocess.run("python -m pip install numpy", shell=True)
subprocess.run("python -m pip install openpyxl", shell=True)
subprocess.run("python -m pip install tqdm", shell=True)
subprocess.run("python -m pip install pysimplegui", shell=True)