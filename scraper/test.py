import os

filedir = os.path.dirname(os.path.abspath(__file__))
print(os.path.join(filedir, "..", "generator", "fill.ps1"))