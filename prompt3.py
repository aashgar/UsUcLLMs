import pandas as pd
import numpy as np
from rapidfuzz import fuzz
import re


file_path = "LLM/usecases.xlsx" 

#sheet_name = "g04-recycling-p3"
#sheet_name = "g11-nsf-p3"
#sheet_name = "g12-camperplus-p3"
#sheet_name = "g13-planningpoker-p3"
#sheet_name = "g23-archivesspace-p3"
sheet_name =  "g28-zooniverse-p3"


def normalize_plantuml(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def script_similarity(script1, script2):
    s1 = normalize_plantuml(script1)
    s2 = normalize_plantuml(script2)
    return fuzz.token_sort_ratio(s1, s2)


df = pd.read_excel(file_path, sheet_name=sheet_name)

llm_columns = sorted([col for col in df.columns if "llm" in col.lower()])

scripts = []
for col in llm_columns:
    full_script = "\n".join(df[col].dropna().astype(str))
    scripts.append(full_script)


n = len(scripts)

sim_matrix = pd.DataFrame(
    index=[f"LLM{i+1}" for i in range(n)],
    columns=[f"LLM{i+1}" for i in range(n)],
    dtype=float
)

pairwise_scores = []

for i in range(n):
    for j in range(n):
        if i == j:
            sim_matrix.iloc[i, j] = 100.0
        elif pd.isna(sim_matrix.iloc[i, j]):
            sim = script_similarity(scripts[i], scripts[j])
            sim = round(sim, 2)
            sim_matrix.iloc[i, j] = sim
            sim_matrix.iloc[j, i] = sim
            pairwise_scores.append(sim)

overall_agreement = round(np.mean(pairwise_scores), 2)


summary_df = pd.DataFrame({
    "Metric": ["Overall Agreement Across All LLMs"],
    "Value (%)": [overall_agreement]
})

blank_row = pd.DataFrame({col: [""] for col in sim_matrix.columns})
combined_df = pd.concat(
    [sim_matrix.reset_index(), blank_row, summary_df],
    ignore_index=True
)

# Write to Excel
output_sheet = f"{sheet_name}_similarity"
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    combined_df.to_excel(writer, sheet_name=output_sheet, index=False)

print(f"\nSimilarity matrix and overall agreement written to sheet '{output_sheet}'")
print("weighted Similarity Matrix:")
print(sim_matrix)
print(f"Overall Agreement: {overall_agreement}%")
