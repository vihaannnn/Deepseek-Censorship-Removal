from transformers import AutoTokenizer, AutoModelForCausalLM
import torch

# Load the tokenizer and model
model_name = "deepseek-ai/deepseek-llm-7b-base"
tokenizer = AutoTokenizer.from_pretrained(model_name)
model = AutoModelForCausalLM.from_pretrained(model_name, torch_dtype=torch.float16, device_map="auto")

# Function to generate text
def generate_text(prompt, max_length=200):
    inputs = tokenizer(prompt, return_tensors="pt").to(model.device)
    
    with torch.no_grad():
        output = model.generate(**inputs, max_length=max_length, temperature=0.7, top_p=0.9)
    
    return tokenizer.decode(output[0], skip_special_tokens=True)

# Example usage
prompt = "Explain the significance of artificial intelligence in modern society."
response = generate_text(prompt)

print("Generated Response:\n", response)