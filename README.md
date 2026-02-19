Generating Use Cases from User Stories using LLMs

This repository contains the datasets, experimental outputs, and Python scripts used in our study on generating use cases from user stories using four LLMs: ChatGPT, Gemini, Claude, and DeepSeek.

The study evaluates inter-LLM agreement across three sequential prompting steps.

Datasets

Nine user story datasets were used:

g03-loudoun

g04-recycling

g11-nsf

g12-camperplus

g13-planningpoker

g14-datahub

g23-archivesspace

g24-unibath

g28-zooniverse

Each dataset consists of structured user stories used as input for the prompts.

Processing Steps
Prompt 1 — Use Case Extraction (prompt1.py)

Input: Raw user stories

Output: Markdown table (| # | Actor | Action | Object |)

One use case per user story

Prompt 2 — Filtering & Normalization (prompt2.py)

Input: Prompt 1 output

Retains objects appearing multiple times

Normalizes terms and standardizes action to “Manage”

Produces a reduced use case table

Prompt 3 — UML Script Generation (prompt3.py)

Input: Prompt 2 output

Output: PlantUML script representing consolidated use cases

Evaluation

Pairwise inter-LLM similarity is computed for each prompt.

Best-match bidirectional row-level similarity is applied where needed.

Overall agreement is calculated as the average of pairwise similarities.

Similarity results are stored in Excel sheets with the suffix _similarity.

How to Run
python prompt1.py
python prompt2.py
python prompt3.py

Repository Contents

Datasets

LLM-generated outputs

Prompt 3 UML scripts

Similarity matrices and overall agreement scores

Source code for similarity computation