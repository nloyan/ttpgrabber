import docx
import requests
import json
import re
from stix2 import MemoryStore, Filter

# specify the input and output file paths
input_file = (r"C:/users/xxxx/Downloads/Python TTP Replacement.docx")
output_file = (r"C:/users/xxxxx/Downloads/Python TTP Replacement.docx")

# download the STIX data feed
response = requests.get("https://raw.githubusercontent.com/mitre/cti/master/enterprise-attack/enterprise-attack.json")
stix_data = json.loads(response.text)
store = MemoryStore(stix_data)

# open the input Word document
doc = docx.Document(input_file)

# loop through each paragraph in the document
for para in doc.paragraphs:
    # loop through each run in the paragraph
    for run in para.runs:
        # check if the run text matches a TTP ID pattern (e.g., T1234)
        if re.match(r"^T\d+(\.\d+)?$", run.text):
            # look up the TTP name using the STIX data feed
            technique_id = run.text
            techniques = store.query([Filter("type", "=", "attack-pattern"), Filter("external_references.external_id", "=", technique_id)])
            if techniques:
                technique = techniques[0]
            else:
                print(f"No TTP found for ID {technique_id}")

            technique_name = technique.name

            # replace the TTP ID with "TTP ID - TTP Name"
            new_text = f"{technique_id} - {technique_name}"
            run.text = new_text

# save the updated document
doc.save(output_file)
print(f"Updated document saved to {output_file}.")
