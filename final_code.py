import pandas as pd
from pathlib import Path
from docxtpl import DocxTemplate
import numpy as np

word_template_path = Path.cwd() / "template.docx"
excel_path = Path.cwd() / "YOBE.xlsx"
output_dir = Path.cwd() / "OUTPUT"
output_dir.mkdir(exist_ok=True)


# Define a function that takes a dictionary and a prefix as parameters
def modify_dict(d, prefix):
    # Loop through the  keys in the dictionary
    for key in d.keys():
        # Check if the key starts with the prefix
        if key.startswith(prefix):
            # Check if the value starts with the prefix after trimming the string
            if d[key].strip().startswith(prefix):
                # Leave the value as it is
                pass
            else:
                # Get the next two keys after the current key
                next_keys = list(d.keys())[
                    list(d.keys()).index(key) + 1 : list(d.keys()).index(key) + 3
                ]
                # Check if both of them have empty values
                if d[next_keys[0]] == "" and d[next_keys[1]] == "":
                    # Change the current value to an empty string
                    d[key] = ""
                # Check if one of them has a value and the current value is empty
                elif (d[next_keys[0]] != "" or d[next_keys[1]] != "") and d[key] == "":
                    # Change the current value to the prefix
                    d[key] = prefix


df = pd.read_excel(excel_path, sheet_name="Sheet1")
df = df.replace(np.nan, "", regex=True)

for record in df.to_dict(orient="records"):
    modify_dict(record, "PB")
    doc = DocxTemplate(word_template_path)
    # print(record)
    doc.render(record)
    output_path = output_dir / f"{record['NAME']}.docx"
    doc.save(output_path)
