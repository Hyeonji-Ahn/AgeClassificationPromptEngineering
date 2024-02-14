import openai
import os
import openpyxl
from automatic_prompt_engineer import ape

from dotenv import load_dotenv, find_dotenv
_ = load_dotenv(find_dotenv())

openai.api_key=os.environ.get("OPENAI_API_KEY")


age = []
samples = []

eval_template = \
"""Instruction: [PROMPT]

Input: [INPUT]
Output: [OUTPUT]"""

prompt_gen_template = \
"""
Based on the instruction, a program produced the following input-output pairs:

    [full_DEMO]
    
    The instruction was to [APE]
"""


textdata = openpyxl.load_workbook('DepressionText_parsed.xlsx')
# Get all the sheets in the workbook
sheets = textdata.sheetnames


# for i in range(5):
#     # Get the active worksheet
#     ws = textdata[sheets[i]]
#     # Iterate through the rows in the worksheet
#     for row in ws.rows:
#         #print(row[1].value , " " ,row[0].value)
#         samples.append([row[1].value,row[0].value])

for sheet in sheets:
    ws = textdata[sheet]
    for row in ws.rows:
        if(row[1].value == "age"):
            continue
        age.append(str(row[1].value))
        samples.append(str(row[0].value))

result, demo_fn = ape.simple_ape(
    dataset=(samples, age),
    eval_template=eval_template,
    eval_model = "davinci-002",
    prompt_gen_model = "davinci-002",
    num_prompts=35
)

print(result)