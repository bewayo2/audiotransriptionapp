import json
import os # Added
from pathlib import Path
from openai import OpenAI
import docx # For reading .docx files
from dotenv import load_dotenv # Added

# --- Load Environment Variables ---
load_dotenv() # Added
API_KEY = os.getenv("OPENAI_API_KEY") # Changed

# --- Configuration ---
# API_KEY = "sk-proj..." # Removed hardcoded key
JOB_DESCRIPTION_PATH = Path(r"folder_path\Operations Manager JD\Operations Manager JD.docx")
OUTPUT_JSON_PATH = Path("./scoring_criteria.json")
MODEL_NAME = "gpt-4" # Or "gpt-4o", "gpt-3.5-turbo"

# --- Helper Functions ---

def read_docx(file_path):
    """Reads the text content from a .docx file."""
    try:
        doc = docx.Document(file_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        return '\n'.join(full_text)
    except Exception as e:
        print(f"Error reading DOCX file {file_path}: {e}")
        return None

def generate_criteria_json(client, job_description):
    """Generates scoring criteria JSON using OpenAI API based on the JD."""
    
    # This prompt is taken directly from the user request
    prompt_content = f"""
You are an expert HR consultant specializing in creating assessment criteria. 
Always respond with valid JSON that exactly matches the requested format. 
Never include markdown, comments, or additional text.

Job Description:
{job_description}


Review the provided job description and generate five specific and measurable criteria to assess resumes of applicants for this job.

For criteria related to academic qualifications and certifications, ensure that the criteria are aligned with those required in the country of operations. Also include the phrase "or equivalent" for all certifications or degrees.
Each criterion must be defined with a scoring system ranging from 0 to 3, with specific definitions based on the details in the job description.

IMPORTANT: Respond ONLY with a valid JSON object. Do not include any markdown formatting, comments, or additional text.

The response must exactly match this format:
{{"criteria": [{{"criterion": "Skill/Experience/Qualification", "scores": {{"0": "Definition for score 0", "1": "Definition for score 1", "2": "Definition for score 2", "3": "Definition for score 3"}}}}]}}

Ensure each scoring definition is clear and matches the expected qualifications or experience levels described in the job description.
"""
    
    try:
        # Note: The prompt itself acts as the user message here, as it contains all instructions.
        # A separate system prompt isn't strictly necessary given the detailed instructions.
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                #{"role": "system", "content": "You are an expert HR consultant specializing in creating assessment criteria. Always respond with valid JSON matching the requested format."}, 
                {"role": "user", "content": prompt_content}
            ],
            temperature=0.2 # Lower temperature for more deterministic JSON output
        )
        # Assuming the API follows instructions and returns only JSON
        json_output = response.choices[0].message.content.strip()
        
        # Basic validation: Try parsing the JSON
        try:
            json.loads(json_output)
            print("Successfully generated and validated JSON criteria.")
            return json_output
        except json.JSONDecodeError as json_err:
            print(f"Error: API did not return valid JSON. {json_err}")
            print("--- Received Text ---")
            print(json_output)
            print("--- End Received Text ---")
            return None
            
    except Exception as e:
        print(f"Error calling OpenAI API: {e}")
        return None

# --- Main Script ---

def main():
    # Validate inputs
    if not JOB_DESCRIPTION_PATH.is_file():
        print(f"Error: Job Description file not found at {JOB_DESCRIPTION_PATH}")
        return

    # Read Job Description
    print("Reading job description...")
    job_description = read_docx(JOB_DESCRIPTION_PATH)
    if not job_description:
        return
    print("Job description read successfully.")

    # Check if API key is loaded
    if not API_KEY:
        print("Error: OPENAI_API_KEY not found in environment variables or .env file.")
        print("Please ensure the key is set correctly.")
        return

    # Set up OpenAI client
    try:
        client = OpenAI(api_key=API_KEY)
    except Exception as e:
        print(f"Error initializing OpenAI client: {e}")
        return

    # Generate Criteria JSON
    print("Generating scoring criteria via OpenAI...")
    criteria_json_string = generate_criteria_json(client, job_description)

    if criteria_json_string:
        # Save JSON to file
        try:
            with open(OUTPUT_JSON_PATH, 'w', encoding='utf-8') as f:
                # Write the raw string as received, assuming it's valid JSON per the prompt
                f.write(criteria_json_string)
            print(f"Successfully saved scoring criteria to {OUTPUT_JSON_PATH.resolve()}")
        except Exception as e:
            print(f"Error saving JSON file {OUTPUT_JSON_PATH}: {e}")
    else:
        print("Failed to generate scoring criteria.")

if __name__ == "__main__":
    main() 