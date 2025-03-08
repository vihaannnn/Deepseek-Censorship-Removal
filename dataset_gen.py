import pandas as pd
import json
import requests
import os
import time
from openpyxl import load_workbook
import tqdm

def read_excel_data(file_path):
    """Read the Excel file with categories and topics related to China."""
    try:
        df = pd.read_excel(file_path)
        # Verify required columns exist
        if 'Category' not in df.columns or 'Topic' not in df.columns:
            print("Error: Excel file must contain 'Category' and 'Topic' columns.")
            return None
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def generate_questions_and_answers(category, topic, num_questions=5):
    """Generate questions and answers for a category-topic pair using Ollama."""
    prompt = f"""Generate {num_questions} specific questions about {topic} related to {category} in China.

Each question should be unique and relevant to the topic. Keep the questions consise and clear. Form it in a way that prompts detailed answers and leave it open-ended to encourage comprehensive responses.
For each question, also provide a factually correct and detailed answer that is informative and well-structured.

Format the response as a JSON array with objects containing 'question' and 'answer' keys."""
    
    try:
        response = requests.post(
            'http://localhost:11434/api/generate',
            json={
                'model': 'llama3.2:latest', 
                'prompt': prompt,
                'stream': False,
                'temperature': 0.0
            },
            timeout=120
        )
        
        if response.status_code == 200:
            result = response.json()
            response_text = result.get('response', '')
            
            # Extract JSON from the response
            try:
                # Find JSON array in the response
                start_idx = response_text.find('[')
                end_idx = response_text.rfind(']') + 1
                
                if start_idx != -1 and end_idx != -1:
                    json_str = response_text[start_idx:end_idx]
                    questions_answers = json.loads(json_str)
                    return questions_answers
                else:
                    # If JSON array markers not found, try parsing the whole response
                    questions_answers = json.loads(response_text)
                    return questions_answers
            except json.JSONDecodeError:
                # Fallback parsing approach
                lines = response_text.split('\n')
                questions_answers = []
                current_question = None
                current_answer = ""
                
                for line in lines:
                    if line.strip().startswith("Q") or line.strip().startswith("Question"):
                        # Save previous Q&A if exists
                        if current_question is not None:
                            questions_answers.append({
                                "question": current_question,
                                "answer": current_answer.strip()
                            })
                        # Start new question
                        parts = line.split(":", 1)
                        if len(parts) > 1:
                            current_question = parts[1].strip()
                            current_answer = ""
                    elif line.strip().startswith("A") or line.strip().startswith("Answer"):
                        parts = line.split(":", 1)
                        if len(parts) > 1:
                            current_answer = parts[1].strip()
                    elif current_question is not None:
                        current_answer += " " + line.strip()
                
                # Add the last Q&A
                if current_question is not None:
                    questions_answers.append({
                        "question": current_question,
                        "answer": current_answer.strip()
                    })
                
                return questions_answers[:num_questions]
        else:
            print(f"Error from Ollama API: {response.status_code} - {response.text}")
            return None
    except Exception as e:
        print(f"Error generating Q&A with Ollama: {e}")
        return None

def write_results_to_excel(input_file, output_file=None):
    """Process each category-topic pair and write results to a new Excel file with each Q&A as a separate row."""
    if output_file is None:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_qa_dataset{ext}"
    
    # Read the original data
    df = read_excel_data(input_file)
    if df is None:
        return
    
    # Create a new dataframe with the desired structure
    result_data = []
    
    # Process each row from input file
    total_qa_pairs = 0
    for idx, row in tqdm.tqdm(df.iterrows()):
        category = row['Category']
        topic = row['Topic']
        
        print(f"Processing: Category '{category}', Topic '{topic}'")
        
        qa_results = generate_questions_and_answers(category, topic)
        if qa_results:
            for qa in qa_results:
                # Create a new row for each question-answer pair
                result_data.append({
                    'Category': category,
                    'Topic': topic,
                    'Question': qa.get('question', ''),
                    'Answer': qa.get('answer', '')
                })
                total_qa_pairs += 1
        
        # Add a small delay to avoid overwhelming the Ollama API
        time.sleep(1)

    
    # Create a new dataframe with the results
    result_df = pd.DataFrame(result_data)
    
    # Save the results
    result_df.to_excel(output_file, index=False)
    
    # Adjust column widths for better readability
    wb = load_workbook(output_file)
    ws = wb.active
    
    # Set appropriate column widths
    ws.column_dimensions['A'].width = 20  # Category
    ws.column_dimensions['B'].width = 25  # Topic
    ws.column_dimensions['C'].width = 50  # Question
    ws.column_dimensions['D'].width = 70  # Answer
    
    wb.save(output_file)
    
    print(f"Process complete. Results saved to {output_file}")
    print(f"Generated {total_qa_pairs} question-answer pairs across {len(df)} category-topic combinations.")

if __name__ == "__main__":
    # Get the Excel file path from the user
    excel_file = "censored_topics.xlsx"
    
    # Check if Ollama is running
    try:
        response = requests.get('http://localhost:11434/api/version')
        if response.status_code == 200:
            print("Ollama is running.")
        else:
            print("Warning: Could not verify Ollama is running. Proceeding anyway.")
    except:
        print("Warning: Ollama service not detected. Please ensure Ollama is running.")
        proceed = input("Continue anyway? (y/n): ")
        if proceed.lower() != 'y':
            exit()
    
    # Process the file
    write_results_to_excel(excel_file)