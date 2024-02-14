import openai
import os
import openpyxl
import yaml

from pathlib import Path
import sys

# Add the directory containing the file to the sys.path list
file_path = Path(r"C:\Users\ahj28\Desktop\Stem Fellowship\automatic_prompt_engineer-main\automatic_prompt_engineer")
sys.path.append(str(file_path))

# Now you can import the modules
import ape
import config


from dotenv import load_dotenv, find_dotenv
_ = load_dotenv(find_dotenv())

openai.api_key=os.environ.get("OPENAI_API_KEY")


age = []
samples = []

eval_template = \
"""Instruction: [PROMPT]
Input: [INPUT]
Output: [OUTPUT]"""


textdata = openpyxl.load_workbook('DepressionText_parsed.xlsx')
# Get all the sheets in the workbook
sheets = textdata.sheetnames

ws = textdata[sheets[0]]
for row in ws.rows:
    age.append(str(row[1].value))
    samples.append(str(row[0].value))

with open(os.path.join(os.path.dirname(__file__), "new_config.yaml")) as f:
    new_config = yaml.safe_load(f)
    
config.update_config(new_config)

result, demo_fn = ape.simple_ape(
    dataset=(age, samples),
    eval_template=eval_template,
)

print(result)