import ollama
import pandas as pd

model_name = "llama3.2:latest"
data = pd.read_csv("questions.csv")

answers = []
for i, row in data.iterrows():
    question = row['question']

    prompt = f"""Answer the following question factually correct and in-depth.

Question:
{question}
"""

    # Run the model
    response = ollama.chat(model=model_name, messages=[{"role": "user", "content": prompt}])
    answers.append(response['message']['content'])


data['answers'] = answers
data.to_csv("dataset.csv", index=False)