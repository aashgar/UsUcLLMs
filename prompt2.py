import pandas as pd
import numpy as np
from rapidfuzz import fuzz

file_path = "LLM/usecases.xlsx"

#sheet_name = "g04-recycling-p2"
#sheet_name = "g11-nsf-p2"
#sheet_name = "g12-camperplus-p2"
#sheet_name = "g13-planningpoker-p2"
#sheet_name = "g23-archivesspace-p2"
sheet_name =  "g28-zooniverse-p2"

weight_actor = 0.3
weight_action = 0.2
weight_object = 0.5


def safe_str(x):
    return "" if pd.isna(x) else str(x)

def weighted_row_similarity(r1, r2):
    actor_sim = fuzz.token_sort_ratio(safe_str(r1["Actor"]), safe_str(r2["Actor"]))
    action_sim = fuzz.token_sort_ratio(safe_str(r1["Action"]), safe_str(r2["Action"]))
    object_sim = fuzz.token_sort_ratio(safe_str(r1["Object"]), safe_str(r2["Object"]))
    return weight_actor * actor_sim + weight_action * action_sim + weight_object * object_sim


def directional_similarity(df1, df2):
    scores = []
    for _, row1 in df1.iterrows():
        best_score = 0
        for _, row2 in df2.iterrows():
            sim = weighted_row_similarity(row1, row2)
            best_score = max(best_score, sim)
        scores.append(best_score)
    return np.mean(scores) if scores else 0

def bidirectional_similarity(df1, df2):
    forward = directional_similarity(df1, df2)
    backward = directional_similarity(df2, df1)
    return (forward + backward) / 2


df = pd.read_excel(file_path, sheet_name=sheet_name)

llm_numbers = sorted({int(col.split("_LLM")[-1]) for col in df.columns if "_LLM" in col})

llm_tables = {}
for i in llm_numbers:
    temp = pd.DataFrame({
        "Actor": df[f"Actor_LLM{i}"],
        "Action": df[f"Action_LLM{i}"],
        "Object": df[f"Object_LLM{i}"]
    }).dropna(how="all").reset_index(drop=True)
    llm_tables[i] = temp


all_rows = []

for llm_id, table in llm_tables.items():
    for idx, row in table.iterrows():

        similarities = []

        for other_id, other_table in llm_tables.items():
            if other_id == llm_id:
                continue

            # Forward (this row → other table)
            forward_best = 0
            for _, other_row in other_table.iterrows():
                sim = weighted_row_similarity(row, other_row)
                forward_best = max(forward_best, sim)

            # Backward (best row in other → this row)
            backward_best = 0
            for _, other_row in other_table.iterrows():
                sim = weighted_row_similarity(other_row, row)
                backward_best = max(backward_best, sim)

            similarities.append((forward_best + backward_best) / 2)

        row_agreement = round(np.mean(similarities), 2) if similarities else 100.0

        all_rows.append({
            "Source LLM": f"LLM{llm_id}",
            "Actor": row["Actor"],
            "Action": row["Action"],
            "Object": row["Object"],
            "Row Agreement (%)": row_agreement
        })

main_table = pd.DataFrame(all_rows)


n = len(llm_numbers)

sim_matrix = pd.DataFrame(
    index=[f"LLM{i}" for i in llm_numbers],
    columns=[f"LLM{i}" for i in llm_numbers],
    dtype=float
)

pairwise_scores = []

for i_idx, i in enumerate(llm_numbers):
    for j_idx, j in enumerate(llm_numbers):
        if i == j:
            sim_matrix.iloc[i_idx, j_idx] = 100.0
        elif pd.isna(sim_matrix.iloc[i_idx, j_idx]):
            sim = bidirectional_similarity(llm_tables[i], llm_tables[j])
            sim = round(sim, 2)
            sim_matrix.iloc[i_idx, j_idx] = sim
            sim_matrix.iloc[j_idx, i_idx] = sim
            pairwise_scores.append(sim)

overall_agreement = round(np.mean(pairwise_scores), 2)

# Add overall agreement row
sim_matrix.loc["Overall Agreement"] = ""
sim_matrix.loc["Overall Agreement", sim_matrix.columns[0]] = overall_agreement

sim_matrix_reset = sim_matrix.reset_index()
sim_matrix_reset.rename(columns={"index": "LLM"}, inplace=True)


# Combine
max_rows = max(len(main_table), len(sim_matrix_reset))

main_table_extended = main_table.reindex(range(max_rows))
matrix_extended = sim_matrix_reset.reindex(range(max_rows))

separator = pd.DataFrame({"": [""] * max_rows})

final_output = pd.concat(
    [main_table_extended, separator, matrix_extended],
    axis=1
)

# Write to Excel
output_sheet = f"{sheet_name}_similarity"

with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    final_output.to_excel(writer, sheet_name=output_sheet, index=False)

print(f"\nResults written to sheet '{output_sheet}' in {file_path}")
print("weighted Similarity Matrix:")
print(sim_matrix)
print(f"Overall Bidirectional Agreement: {overall_agreement}%")
