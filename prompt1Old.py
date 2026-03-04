import pandas as pd
from rapidfuzz import fuzz
from itertools import combinations

# -----------------------------
# Configuration
# -----------------------------
file_path= "LLM/usecases.xlsx"

#sheet_name = "g04-recycling-p1"
#sheet_name = "g11-nsf-p1"
#sheet_name = "g12-camperplus-p1"
#sheet_name = "g13-planningpoker-p1"
#sheet_name = "g23-archivesspace-p1"
sheet_name = "g28-zooniverse-p1"

output_sheet = sheet_name + "_Similarity"

llm_columns_actor= ["Actor_LLM1", "Actor_LLM2", "Actor_LLM3", "Actor_LLM4"]
llm_columns_action = ["Action_LLM1", "Action_LLM2", "Action_LLM3", "Action_LLM4"]
llm_columns_object = ["Object_LLM1", "Object_LLM2", "Object_LLM3", "Object_LLM4"]

weight_actor= 0.3
weight_action = 0.3
weight_object = 0.4

llm_names = ["LLM1", "LLM2", "LLM3", "LLM4"]

# -----------------------------
# Helper Functions
# -----------------------------
def similarity(a, b):
    return fuzz.token_sort_ratio(
        str(a).lower().strip(),
        str(b).lower().strip()
    )

def average_pairwise_similarity(values):
    clean_values = [str(v).lower().strip() for v in values
                    if pd.notna(v) and str(v).strip() != ""]
    if len(clean_values) < 2:
        return 0.0
    pairs = list(combinations(clean_values, 2))
    scores = [fuzz.token_sort_ratio(a, b) for a, b in pairs]
    return round(sum(scores) / len(scores), 2)

# -----------------------------
# Read Data
# -----------------------------
df = pd.read_excel(file_path, sheet_name=sheet_name)

# -----------------------------
# Row-wise Similarity
# -----------------------------
df['Actor_Similarity_Avg'] = df.apply(
    lambda row: average_pairwise_similarity(row[llm_columns_actor]), axis=1)

df['Action_Similarity_Avg'] = df.apply(
    lambda row: average_pairwise_similarity(row[llm_columns_action]), axis=1)

df['Object_Similarity_Avg'] = df.apply(
    lambda row: average_pairwise_similarity(row[llm_columns_object]), axis=1)

overall_actor_avg = round(df['Actor_Similarity_Avg'].mean(), 2)
overall_action_avg = round(df['Action_Similarity_Avg'].mean(), 2)
overall_object_avg = round(df['Object_Similarity_Avg'].mean(), 2)

overall_weighted_avg = round(
    (overall_actor_avg * weight_actor) +
    (overall_action_avg * weight_action) +
    (overall_object_avg * weight_object),
    2
)

# -----------------------------
# weighted Similarity sim_matrix
# -----------------------------
sim_matrix = pd.DataFrame(index=llm_names, columns=llm_names)

for i in range(4):
    for j in range(4):
        if i == j:
            sim_matrix.iloc[i, j] = 100.0
        else:
            weighted_scores = []
            for _, row in df.iterrows():
                actor_sim = similarity(row[llm_columns_actor[i]], row[llm_columns_actor[j]])
                action_sim = similarity(row[llm_columns_action[i]], row[llm_columns_action[j]])
                object_sim = similarity(row[llm_columns_object[i]], row[llm_columns_object[j]])

                weighted = (
                    actor_sim * weight_actor+
                    action_sim * weight_action +
                    object_sim * weight_object
                )
                weighted_scores.append(weighted)

            sim_matrix.iloc[i, j] = round(sum(weighted_scores) / len(weighted_scores), 2)

# -----------------------------
# Append Summary Rows
# -----------------------------
summary_row = {col: "" for col in df.columns}
summary_row[df.columns[0]] = "OVERALL AVERAGE"
summary_row['Actor_Similarity_Avg'] = overall_actor_avg
summary_row['Action_Similarity_Avg'] = overall_action_avg
summary_row['Object_Similarity_Avg'] = overall_object_avg

weighted_row = {col: "" for col in df.columns}
weighted_row[df.columns[0]] = "weightED OVERALL AVERAGE"
weighted_row['Object_Similarity_Avg'] = overall_weighted_avg

df_final = pd.concat([
    df,
    pd.DataFrame([summary_row]),
    pd.DataFrame([weighted_row])
], ignore_index=True)

# -----------------------------
# Prepare sim_matrix for Right-Side Placement
# -----------------------------
sim_matrix_display = sim_matrix.reset_index()
sim_matrix_display.columns = ["LLM"] + llm_names

# Extend sim_matrix rows to match df_final height
if len(sim_matrix_display) < len(df_final):
    extra_rows = pd.DataFrame(
        [[""] * len(sim_matrix_display.columns)] *
        (len(df_final) - len(sim_matrix_display)),
        columns=sim_matrix_display.columns
    )
    sim_matrix_display = pd.concat([sim_matrix_display, extra_rows], ignore_index=True)

# Rename to avoid column duplication
sim_matrix_display.columns = [col for col in sim_matrix_display.columns]

# -----------------------------
# Add Empty Separator Column
# -----------------------------
separator_column = pd.DataFrame({"": [""] * len(df_final)})

# -----------------------------
# Merge Everything
# -----------------------------
df_output = pd.concat([
    df_final.reset_index(drop=True),
    separator_column,
    sim_matrix_display.reset_index(drop=True)
], axis=1)

# -----------------------------
# Write to Excel
# -----------------------------
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a',
                    if_sheet_exists='replace') as writer:
    df_output.to_excel(writer, sheet_name=output_sheet, index=False)

print(f"Sheet '{output_sheet}' updated successfully.")
print("weighted Similarity sim_matrix:")
print(sim_matrix)
print(f"weighted Overall Average: {overall_weighted_avg}")
