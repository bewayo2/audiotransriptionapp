import json
import os
import re # Added for name extraction
from pathlib import Path
from openai import OpenAI
import docx # For reading .docx files
import time
import pandas as pd # Added for DataFrame and Excel export
from dotenv import load_dotenv # Added

# --- Load Environment Variables ---
load_dotenv() # Added
API_KEY = os.getenv("OPENAI_API_KEY") # Changed

# --- Configuration ---
JOB_DESCRIPTION_PATH = Path(r"C:\Users\timsi\Downloads\AI Recruitment Demo\Operations Manager JD\Operations Manager JD.docx")
CRITERIA_JSON_PATH = Path("./scoring_criteria.json")
CV_FOLDER_PATH = Path("./generated_cvs")
OUTPUT_EXCEL_FILE = Path(r"C:\Users\timsi\Downloads\AI Recruitment Demo\cv_score_summary.xlsx") # Changed output path
MODEL_NAME = "gpt-4" # Or "gpt-4o", "gpt-3.5-turbo"
REQUEST_DELAY_SECONDS = 1 # Delay between API calls to avoid rate limits

# --- Helper Functions ---

def read_docx(file_path):
    """Reads the text content from a .docx file."""
    try:
        doc = docx.Document(file_path)
        full_text = [para.text for para in doc.paragraphs]
        return '\n'.join(full_text)
    except Exception as e:
        print(f"Error reading DOCX file {file_path}: {e}")
        return None

def read_text_file(file_path):
    """Reads the text content from a generic text file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"Error reading text file {file_path}: {e}")
        return None

def load_json_file(file_path):
    """Loads JSON data from a file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Error: JSON file not found at {file_path}")
        return None
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in file {file_path}: {e}")
        return None
    except Exception as e:
        print(f"Error reading JSON file {file_path}: {e}")
        return None

def extract_name_from_cv(cv_text):
    """Attempts to extract a name from the beginning of the CV text."""
    if not cv_text:
        return "Name Not Found"
    
    lines = cv_text.split('\n')
    name_pattern = re.compile(r"^[A-Z][a-z']+ [A-Z][a-z']+'?$")
    
    for i, line in enumerate(lines):
        cleaned_line = line.strip()
        if len(cleaned_line) > 3 and ' ' in cleaned_line and cleaned_line.istitle():
             if name_pattern.match(cleaned_line):
                 return cleaned_line
        if i > 10:
            break
            
    for line in lines:
        if line.strip():
            return line.strip()
            
    return "Name Not Found"

def calculate_total_score(score_data):
    """Calculates the total score from the parsed score data."""
    total_score = 0
    if not score_data or 'scores' not in score_data:
        print("Warning: Invalid score data format, cannot calculate total score.")
        return 0
        
    for item in score_data['scores']:
        try:
            score_value = int(item.get('score', 0))
            total_score += score_value
        except (ValueError, TypeError):
            print(f"Warning: Could not parse score '{item.get('score')}' for criterion '{item.get('criterion_id')}'. Treating as 0.")
            continue
    return total_score

def score_resume_and_get_json_string(client, job_description, criteria_json_with_ids, resume_text):
    """Scores a resume against criteria using OpenAI API and returns JSON string."""
    
    prompt_content = f"""
You are an expert HR consultant evaluating a resume against specific job criteria.

Job Description:
{job_description}

Assessment Criteria:
{json.dumps(criteria_json_with_ids, indent=2)} 

Resume Text:
{resume_text}

Evaluate the resume against each criterion. Provide scores (0-3) based solely on the resume content. Justify your scores with verbatim quotes drawn from the resume. Scores must align with the criteria definition, e.g., if a criterion's score of 1 corresponds to having an associate degree, then assign this score only if the resume indicates an associate degree or higher.

If the resume_text is an empty string, never fabricate content. Score 0 for all criteria and use "No text extracted. Recommend manual scoring." as the justification for the scores.

Format your response as a JSON object with no additional commentary:

{{"scores": [{{"criterion_id": "ID from criteria_json", "score": "0-3", "justification": "Direct quote from resume explaining score", "excerpts": ["Direct quote from resume supporting score", "Additional supporting quote if available"]}}]}}

Notes:
- Use only information present in the resume
- Never make up information or make assumptions
- Include direct quotes from the resume in excerpts
- Ensure justification clearly explains the score based on resume content and aligns with the criteria definitions.
"""

    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "user", "content": prompt_content}
            ],
            temperature=0.1
        )
        json_output_string = response.choices[0].message.content.strip()
        
        # Validate the output is JSON before returning
        try:
            json.loads(json_output_string)
            return json_output_string # Return the validated JSON string
        except json.JSONDecodeError as json_err:
            print(f"Error: API response is not valid JSON. {json_err}")
            print("--- Received Text ---")
            print(json_output_string)
            print("--- End Received Text ---")
            return None
            
    except Exception as e:
        print(f"Error calling OpenAI API: {e}")
        return None

# --- Main Script ---

def main():
    # Validate input paths
    if not JOB_DESCRIPTION_PATH.is_file():
        print(f"Error: Job Description file not found: {JOB_DESCRIPTION_PATH}")
        return
    if not CRITERIA_JSON_PATH.is_file():
        print(f"Error: Criteria JSON file not found: {CRITERIA_JSON_PATH}")
        return
    if not CV_FOLDER_PATH.is_dir():
        print(f"Error: CV folder not found: {CV_FOLDER_PATH}")
        return

    # Load Job Description and Criteria
    print("Loading Job Description...")
    job_description = read_docx(JOB_DESCRIPTION_PATH)
    if not job_description: return
    
    print("Loading Scoring Criteria...")
    criteria_data = load_json_file(CRITERIA_JSON_PATH)
    if not criteria_data or 'criteria' not in criteria_data: return
    
    # Create a mapping from criterion ID to criterion text and store original criterion texts
    criteria_id_to_text = {}
    criterion_texts_ordered = []
    criteria_json_with_ids = {'criteria': []}
    for i, criterion in enumerate(criteria_data['criteria']):
        criterion_id = f"criterion_{i+1}"
        criterion_text = criterion.get('criterion', criterion_id) # Use ID as fallback text
        criterion['id'] = criterion_id
        criteria_json_with_ids['criteria'].append(criterion)
        criteria_id_to_text[criterion_id] = criterion_text
        criterion_texts_ordered.append(criterion_text)
    print(f"Loaded {len(criteria_json_with_ids['criteria'])} criteria.")

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

    # Process each CV and collect results
    cv_files = list(CV_FOLDER_PATH.glob("*.txt"))
    print(f"Found {len(cv_files)} CV files in {CV_FOLDER_PATH}.")
    
    results_list = [] # List to store results for DataFrame
    processed_count = 0

    for cv_path in cv_files:
        print(f"\nProcessing CV: {cv_path.name}")

        cv_text = read_text_file(cv_path)
        if cv_text is None: continue

        # Initialize result dictionary for this CV
        cv_result = {"CV Filename": cv_path.name}
        # Initialize excerpt columns with a default value
        for text in criterion_texts_ordered:
            cv_result[text] = "N/A"
        total_score = 0 # Default score

        if not cv_text.strip():
            print("  Resume file is empty. Assigning score 0.")
            # Keep total_score = 0 and excerpts as "N/A"
        else:
            score_json_string = score_resume_and_get_json_string(client, job_description, criteria_json_with_ids, cv_text)
            if score_json_string:
                try:
                    score_data = json.loads(score_json_string)
                    total_score = calculate_total_score(score_data)
                    print(f"  Successfully scored. Total Score: {total_score}")

                    # Extract excerpts for each criterion
                    if 'scores' in score_data:
                        excerpts_found = {}
                        for score_item in score_data['scores']:
                            criterion_id = score_item.get('criterion_id')
                            criterion_text = criteria_id_to_text.get(criterion_id)
                            if criterion_text:
                                excerpts = score_item.get('excerpts', [])
                                # Join list of excerpts into a single string
                                excerpts_found[criterion_text] = "; ".join(excerpts) if excerpts else "No specific excerpt found"
                            else:
                                print(f"Warning: Criterion ID '{criterion_id}' from score response not found in loaded criteria.")
                        # Update the cv_result with found excerpts
                        cv_result.update(excerpts_found)
                except json.JSONDecodeError:
                    print("  Failed to parse score JSON. Assigning score 0.")
                    # Score remains 0, excerpts remain N/A
            else:
                print("  Failed to generate score from API. Assigning score 0.")
                # Score remains 0, excerpts remain N/A

        cv_result["Total Score"] = total_score # Add/Update total score
        results_list.append(cv_result) # Append the complete result for this CV
        processed_count += 1

        print(f"  Waiting {REQUEST_DELAY_SECONDS} second(s)...")
        time.sleep(REQUEST_DELAY_SECONDS)

    print(f"\n--- Processing Complete --- ")
    print(f"Attempted to process {len(cv_files)} CVs.")
    print(f"Successfully added results for {processed_count} CVs.")

    # --- Compile and Save Excel --- 
    if not results_list:
        print("No results collected. Cannot create Excel summary.")
        return

    print("\nCreating Excel summary...")
    # Define column order: Filename, Total Score, then criteria excerpts
    column_order = ["CV Filename", "Total Score"] + criterion_texts_ordered
    df_summary = pd.DataFrame(results_list)
    df_summary = df_summary[column_order] # Reorder columns
    
    # Sort by Total Score descending
    df_summary = df_summary.sort_values(by="Total Score", ascending=False)
    
    # Save to Excel
    try:
        df_summary.to_excel(OUTPUT_EXCEL_FILE, index=False, engine='openpyxl')
        print(f"Successfully created summary file: {OUTPUT_EXCEL_FILE.resolve()}")
    except Exception as e:
        print(f"Error saving Excel summary file: {e}")

if __name__ == "__main__":
    main() 