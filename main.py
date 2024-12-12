import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from tkinter import font as tkfont
from PIL import Image
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
import json
import threading
import queue
import subprocess
import pytesseract
import textract
import docx
import PyPDF2
from win32com.client import Dispatch
import pythoncom
import spacy
from spacy.language import Language
from spacy.pipeline import EntityRuler
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.decomposition import LatentDirichletAllocation
import nltk
import os

# Initialize NLTK resources
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('vader_lexicon')


from nltk.sentiment import SentimentIntensityAnalyzer


try:
    nlp = spacy.load('en_core_web_sm')
except OSError:
    messagebox.showerror("Model Error", "The spaCy model 'en_core_web_sm' is not installed. Please run 'python -m spacy download en_core_web_sm' to install it.")
    exit()


PREDEFINED_SKILLS = {
    'python', 'java', 'c++', 'c#', 'javascript', 'typescript', 'sql', 'docker', 'kubernetes',
    'spring boot', 'flask', 'git', 'machine learning', 'data analysis', 'communication',
    'problem-solving', 'teamwork', 'project management', 'time management', 'linux',
    'aws', 'azure', 'nosql', 'jira', 'scrum', 'tdd', 'ci/cd', 'api development',
    'chatbot', 'continuous integration', 'continuous deployment', 'agile', 'react',
    'node.js', 'html', 'css', 'selenium', 'tensorflow', 'pytorch', 'sql server',
    'postgresql', 'mongodb', 'redis', 'graphql', 'restful services', 'data visualization',
    'd3.js', 'business intelligence', 'excel', 'power bi', 'tableau', 'project planning',
    'risk management', 'stakeholder management', 'user experience', 'ux design',
    'user interface', 'ui design', 'business analysis', 'strategic planning', 'leadership',
    'presentation skills', 'public speaking', 'technical writing', 'devops',
    'microservices', 'virtualization', 'blockchain', 'iot', 'big data', 'hadoop',
    'spark', 'php', 'ruby', 'perl', 'go', 'rust', 'swift', 'objective-c',
    'matlab', 'sas', 'stata', 'powershell', 'bash scripting', 'unit testing',
    'integration testing', 'quality assurance', 'mobile development', 'android',
    'ios development', 'full stack development', 'backend development', 'frontend development',
    'software architecture', 'cloud computing', 'serverless', 'ci/cd pipelines',
    'automation', 'version control', 'linux administration', 'windows administration',
    'networking', 'security', 'ethical hacking', 'cryptography', 'database design',
    'data modeling', 'data mining', 'natural language processing', 'computer vision',
    'robotics', 'reinforcement learning', 'deep learning', 'predictive analytics',
    'prescriptive analytics', 'descriptive analytics', 'data engineering',
    'data warehousing', 'data governance', 'data quality', 'data privacy',
    'compliance', 'gdpr', 'hipaa', 'iso standards', 'itil', 'it governance',
    'service management', 'supply chain management', 'inventory management',
    'salesforce', 'sap', 'oracle', 'hubspot', 'zendesk', 'atlassian', 'confluence',
    'trello', 'asana', 'monday.com', 'slack', 'zoom', 'microsoft teams', 'google workspace',
    'powerpoint', 'word', 'access', 'outlook', 'visio', 'sharepoint',
    'powerapps', 'azure devops', 'terraform', 'ansible', 'chef', 'puppet',
    'spring', 'react.js', 'chat bot'
}


DEFAULT_WEIGHT = 1
keyword_weights = {skill: DEFAULT_WEIGHT for skill in PREDEFINED_SKILLS}


sia = SentimentIntensityAnalyzer()


@Language.component("correct_entity_labels")
def correct_entity_labels(doc):
    corrected_ents = []
    for ent in doc.ents:
        ent_text = ent.text.lower()
        if ent_text in PREDEFINED_SKILLS:
            corrected_ents.append((ent.start_char, ent.end_char, "SKILL"))
        else:
            corrected_ents.append((ent.start_char, ent.end_char, ent.label_))

    new_ents = []
    seen = set()
    for start, end, label in corrected_ents:
        if (start, end) not in seen:
            span = doc.char_span(start, end, label=label, alignment_mode='contract')
            if span:
                new_ents.append(span)
                seen.add((start, end))
    doc.ents = new_ents
    return doc


ruler = nlp.add_pipe("entity_ruler", before="ner", config={"overwrite_ents": True})
patterns = [
    {"label": "SKILL", "pattern": skill} for skill in PREDEFINED_SKILLS
]
ruler.add_patterns(patterns)


nlp.add_pipe("correct_entity_labels", after="ner")


q = queue.Queue()


data_lock = threading.Lock()
result_text = ""
resume_skills = set()
matched_skills = set()
missing_skills = set()


match_rate = 0.0
resume_category = "N/A"
score = 0


def enqueue_task(task):
    with data_lock:
        q.put(task)


def check_queue():
    while not q.empty():
        task = q.get()
        if task[0] == "showwarning":
            messagebox.showwarning(*task[1])
        elif task[0] == "showinfo":
            messagebox.showinfo(*task[1])
        elif task[0] == "update_plot":
            update_visualization(*task[1])
        elif task[0] == "update_insights":
            update_insights(*task[1])
    root.after(100, check_queue)


def preprocess_text(text):
    doc = nlp(text.lower())
    words = [token.text for token in doc if not token.is_stop and not token.is_punct]
    return words


def extract_text(file_path):
    try:
        if file_path.endswith('.docx'):
            return extract_text_from_docx(file_path)
        elif file_path.endswith('.pdf'):
            return extract_text_from_pdf(file_path)
        elif file_path.endswith('.txt'):
            return extract_text_from_txt(file_path)
        elif file_path.endswith('.rtf'):
            return extract_text_from_rtf(file_path)
        elif file_path.endswith('.doc'):
            return extract_text_from_doc(file_path)
        elif file_path.endswith('.odt'):
            return extract_text_from_odt(file_path)
        elif file_path.lower().endswith(('.jpg', '.jpeg', '.png')):
            return extract_text_from_image(file_path)
        else:
            messagebox.showerror("Unsupported File", "Please select a supported file.")
            return ""
    except Exception as e:
        messagebox.showerror("Extraction Error", f"An error occurred while extracting text:\n{str(e)}")
        return ""


def extract_text_from_docx(docx_file):
    doc = docx.Document(docx_file)
    return '\n'.join([para.text for para in doc.paragraphs])

def extract_text_from_pdf(pdf_file):
    text = ""
    with open(pdf_file, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted
    return text

def extract_text_from_txt(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()

def extract_text_from_rtf(file_path):
    try:
        output = subprocess.check_output(['unrtf', '--text', file_path], stderr=subprocess.STDOUT)
        return output.decode('utf-8')
    except Exception as e:
        print(f"RTF Extraction Error: {e}")
        return ""

def extract_text_from_doc(file_path):
    try:
        pythoncom.CoInitialize()
        word = Dispatch("Word.Application")
        word.visible = False
        doc = word.Documents.Open(file_path)
        text = doc.Content.Text
        doc.Close()
        word.Quit()
        return text
    except Exception as e:
        print(f"DOC Extraction Error: {e}")
        return ""

def extract_text_from_odt(file_path):
    try:
        return textract.process(file_path).decode('utf-8')
    except Exception as e:
        print(f"ODT Extraction Error: {e}")
        return ""

def extract_text_from_image(file_path):
    try:
        img = Image.open(file_path)
        text = pytesseract.image_to_string(img)
        return text
    except Exception as e:
        print(f"Image Extraction Error: {e}")
        return ""


def upload_document():
    file_path = filedialog.askopenfilename(filetypes=[
        ("Document Files", "*.docx;*.pdf;*.txt;*.rtf;*.doc;*.odt"),
        ("Image Files", "*.jpg;*.jpeg;*.png")
    ])
    if not file_path:
        return
    extracted_text = extract_text(file_path)
    if extracted_text:
        resume_textarea.delete("1.0", "end")
        resume_textarea.insert("1.0", extracted_text)

def save_analysis_results():
    if not result_text:
        messagebox.showerror("Save Error", "No analysis results to save. Please analyze a resume first.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                             filetypes=[("Text Files", "*.txt"), ("JSON Files", "*.json")])
    if file_path:
        try:
            with open(file_path, "w", encoding='utf-8') as file:
                file.write(result_text)
            messagebox.showinfo("Save Successful", "Analysis results saved successfully.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Error occurred while saving the file:\n{str(e)}")

def calculate_resume_score(resume_skills, job_skills):
    score = 0
    for skill in job_skills:
        if skill in resume_skills:
            score += keyword_weights.get(skill, DEFAULT_WEIGHT)
    return score

def categorize_job_posting(job_posting_text):
    categories_keywords = {
        'Engineering': ['engineer', 'engineering', 'mechanical', 'civil', 'electrical', 'chemical'],
        'Marketing': ['marketing', 'digital', 'seo', 'brand', 'content'],
        'Sales': ['sales', 'account executive', 'business development'],
        'Finance': ['finance', 'financial analyst', 'accountant'],
        'Human Resources': ['human resources', 'hr', 'recruiter', 'talent acquisition'],
        'Information Technology': ['it', 'information technology', 'system administrator', 'network'],
        'Healthcare': ['healthcare', 'nurse', 'medical', 'pharmacist'],
        'Education': ['teacher', 'education', 'academic advisor'],
        'Operations': ['operations', 'logistics', 'supply chain'],
        'Hospitality': ['hospitality', 'serving', 'guests', 'catering', 'restaurant', 'bar']
    }
    
    job_posting_lower = job_posting_text.lower()
    category_scores = {category: 0 for category in categories_keywords}
    
    for category, keywords in categories_keywords.items():
        for keyword in keywords:
            if keyword in job_posting_lower:
                category_scores[category] += 1
    

    best_category = max(category_scores, key=category_scores.get)
    

    if category_scores[best_category] == 0:
        best_category = 'Operations'
    
    return best_category


def generate_insights(matched_skills, missing_skills, entities, match_rate, recommended_jobs):
    insights = ""
    
    # 1. Top Matched Skills
    insights += "1. **Top Matched Skills:**\n"
    if matched_skills:
        for skill in matched_skills:
            insights += f"   - {skill.capitalize()}\n"
    else:
        insights += "   - No matching skills found.\n"
    
    # 2. Skills Missing
    insights += "\n2. **Skills Missing:**\n"
    if missing_skills:
        for skill in missing_skills:
            insights += f"   - {skill.capitalize()}\n"
    else:
        insights += "   - No missing skills. Excellent match!\n"
    
    # 3. Education and Certifications
    education = [ent[0] for ent in entities if ent[1] in ['ORG', 'EDUCATION']]
    certifications = [ent[0] for ent in entities if ent[1] in ['CERT', 'WORK_OF_ART']]
    insights += "\n3. **Education and Certifications:**\n"
    if education or certifications:
        if education:
            insights += "   - Educational Institutions: " + ", ".join(education[:5]) + "\n"
        if certifications:
            insights += "   - Certifications: " + ", ".join(certifications[:5]) + "\n"
    else:
        insights += "   - No educational institutions or certifications identified.\n"
    
    # 4. Additional Skills and Technologies
    additional_skills = matched_skills.union(missing_skills)
    insights += "\n4. **Additional Skills and Technologies:**\n"
    if additional_skills:
        # Convert set to list before slicing
        insights += "   - " + ", ".join(list(additional_skills)[:5]) + "\n"
    else:
        insights += "   - No additional skills or technologies identified.\n"
    
    # 5. Recommended Jobs
    insights += "\n5. **Recommended Jobs:**\n"
    if recommended_jobs:
        for job in recommended_jobs:
            insights += f"   - {job}\n"
    else:
        insights += "   - No recommended jobs available.\n"
    
    insights += "\n6. **Overall Recommendation:**\n"
    if match_rate > 70:
        insights += "   - Your resume is well-aligned with the job posting. Consider applying!\n"
    elif match_rate > 40:
        insights += "   - Your resume matches some key aspects of the job posting. Consider tailoring it further to highlight missing skills.\n"
    else:
        insights += "   - Your resume has limited alignment with the job posting. Consider acquiring the missing skills and revising your resume to better match the job requirements.\n"
    
    return insights


def update_visualization(matched_skills, missing_skills):
    vis_window = tk.Toplevel(root)
    vis_window.title("Visualization")
    vis_window.geometry("1000x600")
    vis_window.grab_set()


    notebook = ttk.Notebook(vis_window)
    notebook.pack(expand=True, fill='both')


    pie_frame = ttk.Frame(notebook)
    notebook.add(pie_frame, text='Skills Distribution')


    bar_frame = ttk.Frame(notebook)
    notebook.add(bar_frame, text='Top Skills')

    labels = ['Matched Skills', 'Missing Skills']
    sizes = [len(matched_skills), len(missing_skills)]
    colors = ['#2ecc71', '#e74c3c']
    explode = (0.1, 0)

    fig1, ax1 = plt.subplots(figsize=(6, 6))
    ax1.pie(sizes, explode=explode, labels=labels, colors=colors,
            autopct='%1.1f%%', shadow=True, startangle=140)
    ax1.axis('equal')
    ax1.set_title('Skills Distribution', fontsize=16)

    canvas1 = FigureCanvasTkAgg(fig1, master=pie_frame)
    canvas1.draw()
    canvas1.get_tk_widget().pack(expand=True, fill='both')

    if matched_skills:
        skill_freq = {skill: 1 for skill in matched_skills}


        word_freq = pd.Series(skill_freq).sort_values(ascending=False).head(10)

        fig2, ax2 = plt.subplots(figsize=(8, 6))
        sns.barplot(x=word_freq.values, y=word_freq.index, ax=ax2, palette='viridis')
        ax2.set_title('Top Matched Skills', fontsize=16)
        ax2.set_xlabel('Frequency', fontsize=14)
        ax2.set_ylabel('Skills', fontsize=14)
        plt.tight_layout()

        canvas2 = FigureCanvasTkAgg(fig2, master=bar_frame)
        canvas2.draw()
        canvas2.get_tk_widget().pack(expand=True, fill='both')
    else:
        no_data_label = ttk.Label(bar_frame, text="No matched skills to display.", font=("Arial", 14))
        no_data_label.pack(pady=20)


    close_button = ttk.Button(vis_window, text="Close", command=vis_window.destroy)
    close_button.pack(pady=10)

def update_insights(insights):
    insights_textarea.config(state='normal')
    insights_textarea.delete("1.0", "end")
    insights_textarea.insert("1.0", insights)
    insights_textarea.config(state='disabled')

def analyze_sentiment(text):
    sentiment = sia.polarity_scores(text)
    return sentiment['compound']


def generate_report_with_gemini(name, contact, insights, match_rate, resume_category, score):
    GENIUS_API_KEY = os.getenv("GENIUS_API_KEY")
    if not GENIUS_API_KEY:
        messagebox.showerror("API Key Error", "Gemini API key not found. Please set the GENIUS_API_KEY environment variable.")
        return None


    try:
        import google.generativeai as genai
    except ImportError:
        messagebox.showerror("Import Error", "google.generativeai package is not installed.")
        return None

    genai.configure(api_key=GENIUS_API_KEY)

    model = genai.GenerativeModel("gemini-1.5-flash")
    prompt = f"""
    Generate a professional improvement report for a CV based on the following details:

    Name: {name}
    Contact Information: {contact}

    Job Description Insights:
    {insights}

    Match Rate: {match_rate:.2f}%
    Resume Category: {resume_category}
    Resume Score: {score}

    Provide detailed suggestions to improve the CV, focusing on areas with low match rates, missing skills, and enhancing the presentation of existing skills and experiences.
    """

    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        print(f"Gemini API Error: {e}")
        return None


def generate_text():
    if not result_text:
        messagebox.showerror("Generate Error", "No analysis results available. Please analyze a resume first.")
        return


    global match_rate, resume_category, score, insights


    if not insights:
        messagebox.showerror("Generate Error", "No insights available to generate a report.")
        return

    def submit_improvement_report():
        name = name_entry.get().strip()
        contact = contact_entry.get().strip()
        if not name or not contact:
            messagebox.showerror("Input Error", "Please enter both your name and contact information.")
            return
        

        report = generate_report_with_gemini(name, contact, insights, match_rate, resume_category, score)
        if report:
            report_window = tk.Toplevel(root)
            report_window.title("Improvement Report")
            report_window.geometry("800x600")
            report_window.grab_set()

            report_textarea = scrolledtext.ScrolledText(report_window, width=100, height=35, bg="white", relief="flat",
                                                       font=("Arial", 12))
            report_textarea.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
            report_textarea.insert("1.0", report)
            report_textarea.config(state='disabled')

 
            close_button = ttk.Button(report_window, text="Close", command=report_window.destroy)
            close_button.pack(pady=10)
        else:
            messagebox.showerror("Generation Error", "Failed to generate report.")


    report_input_window = tk.Toplevel(root)
    report_input_window.title("Generate Improvement Report")
    report_input_window.geometry("400x300")
    report_input_window.grab_set()

    tk.Label(report_input_window, text="Name:", font=("Arial", 12)).pack(pady=10)
    name_entry = ttk.Entry(report_input_window, width=50, font=("Arial", 12))
    name_entry.pack(pady=5)

    tk.Label(report_input_window, text="Contact Information:", font=("Arial", 12)).pack(pady=10)
    contact_entry = ttk.Entry(report_input_window, width=50, font=("Arial", 12))
    contact_entry.pack(pady=5)

    submit_button = ttk.Button(report_input_window, text="Generate Report", command=submit_improvement_report)
    submit_button.pack(pady=20)


def save_config_on_exit():

    config = {
        "keyword_weights": keyword_weights
    }
    with open('config.json', 'w') as f:
        json.dump(config, f)
    root.destroy()

def show_visualization():
    if not matched_skills and not missing_skills:
        messagebox.showerror("Visualization Error", "No analysis results available. Please analyze a resume first.")
        return

    vis_window = tk.Toplevel(root)
    vis_window.title("Visualization")
    vis_window.geometry("1000x600")
    vis_window.grab_set()


    notebook = ttk.Notebook(vis_window)
    notebook.pack(expand=True, fill='both')

    pie_frame = ttk.Frame(notebook)
    notebook.add(pie_frame, text='Skills Distribution')

    bar_frame = ttk.Frame(notebook)
    notebook.add(bar_frame, text='Top Skills')


    labels = ['Matched Skills', 'Missing Skills']
    sizes = [len(matched_skills), len(missing_skills)]
    colors = ['#2ecc71', '#e74c3c']
    explode = (0.1, 0)

    fig1, ax1 = plt.subplots(figsize=(6, 6))
    ax1.pie(sizes, explode=explode, labels=labels, colors=colors,
            autopct='%1.1f%%', shadow=True, startangle=140)
    ax1.axis('equal')
    ax1.set_title('Skills Distribution', fontsize=16)

    canvas1 = FigureCanvasTkAgg(fig1, master=pie_frame)
    canvas1.draw()
    canvas1.get_tk_widget().pack(expand=True, fill='both')


    if matched_skills:
        skill_freq = {skill: 1 for skill in matched_skills}


        word_freq = pd.Series(skill_freq).sort_values(ascending=False).head(10)

        fig2, ax2 = plt.subplots(figsize=(8, 6))
        sns.barplot(x=word_freq.values, y=word_freq.index, ax=ax2, palette='viridis')
        ax2.set_title('Top Matched Skills', fontsize=16)
        ax2.set_xlabel('Frequency', fontsize=14)
        ax2.set_ylabel('Skills', fontsize=14)
        plt.tight_layout()

        canvas2 = FigureCanvasTkAgg(fig2, master=bar_frame)
        canvas2.draw()
        canvas2.get_tk_widget().pack(expand=True, fill='both')
    else:
        no_data_label = ttk.Label(bar_frame, text="No matched skills to display.", font=("Arial", 14))
        no_data_label.pack(pady=20)

    close_button = ttk.Button(vis_window, text="Close", command=vis_window.destroy)
    close_button.pack(pady=10)


def analyze_sentiment(text):
    sentiment = sia.polarity_scores(text)
    return sentiment['compound']


def generate_report_with_gemini(name, contact, insights, match_rate, resume_category, score):
    GENIUS_API_KEY = os.getenv("GENIUS_API_KEY")
    if not GENIUS_API_KEY:
        messagebox.showerror("API Key Error", "Gemini API key not found. Please set the GENIUS_API_KEY environment variable.")
        return None


    try:
        import google.generativeai as genai
    except ImportError:
        messagebox.showerror("Import Error", "google.generativeai package is not installed.")
        return None

    genai.configure(api_key=GENIUS_API_KEY)

    model = genai.GenerativeModel("gemini-1.5-flash")
    prompt = f"""
    Generate a professional improvement report for a CV based on the following details:

    Name: {name}
    Contact Information: {contact}

    Job Description Insights:
    {insights}

    Match Rate: {match_rate:.2f}%
    Resume Category: {resume_category}
    Resume Score: {score}

    Provide detailed suggestions to improve the CV, focusing on areas with low match rates, missing skills, and enhancing the presentation of existing skills and experiences.
    """

    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        print(f"Gemini API Error: {e}")
        return None

def generate_text():
    if not result_text:
        messagebox.showerror("Generate Error", "No analysis results available. Please analyze a resume first.")
        return

    global match_rate, resume_category, score, insights


    if not insights:
        messagebox.showerror("Generate Error", "No insights available to generate a report.")
        return

    def submit_improvement_report():
        name = name_entry.get().strip()
        contact = contact_entry.get().strip()
        if not name or not contact:
            messagebox.showerror("Input Error", "Please enter both your name and contact information.")
            return
        
        report = generate_report_with_gemini(name, contact, insights, match_rate, resume_category, score)
        if report:
            report_window = tk.Toplevel(root)
            report_window.title("Improvement Report")
            report_window.geometry("800x600")
            report_window.grab_set()

            report_textarea = scrolledtext.ScrolledText(report_window, width=100, height=35, bg="white", relief="flat",
                                                       font=("Arial", 12))
            report_textarea.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
            report_textarea.insert("1.0", report)
            report_textarea.config(state='disabled')

            close_button = ttk.Button(report_window, text="Close", command=report_window.destroy)
            close_button.pack(pady=10)
        else:
            messagebox.showerror("Generation Error", "Failed to generate report.")

    report_input_window = tk.Toplevel(root)
    report_input_window.title("Generate Improvement Report")
    report_input_window.geometry("400x300")
    report_input_window.grab_set()

    tk.Label(report_input_window, text="Name:", font=("Arial", 12)).pack(pady=10)
    name_entry = ttk.Entry(report_input_window, width=50, font=("Arial", 12))
    name_entry.pack(pady=5)

    tk.Label(report_input_window, text="Contact Information:", font=("Arial", 12)).pack(pady=10)
    contact_entry = ttk.Entry(report_input_window, width=50, font=("Arial", 12))
    contact_entry.pack(pady=5)

    submit_button = ttk.Button(report_input_window, text="Generate Report", command=submit_improvement_report)
    submit_button.pack(pady=20)


def analyze_resume():
    job_posting_text = job_posting_textarea.get("1.0", "end-1c").strip()
    resume_text = resume_textarea.get("1.0", "end-1c").strip()

    if not job_posting_text or not resume_text:
        messagebox.showerror("Error", "Please fill in both the job posting and resume text areas.")
        return


    progress_bar = ttk.Progressbar(root, mode='determinate', maximum=100)
    progress_bar.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
    progress_bar.start()

    threading.Thread(target=analyze_resume_thread, args=(progress_bar,)).start()


    check_queue()

def analyze_resume_thread(progress_bar):
    global result_text, resume_skills, matched_skills, missing_skills, match_rate, resume_category, score, insights  # Declare globals
    pythoncom.CoInitialize()
    try:
        job_posting_text = job_posting_textarea.get("1.0", "end-1c")
        resume_text = resume_textarea.get("1.0", "end-1c")
        
        job_category = categorize_job_posting(job_posting_text)
        
        job_doc = nlp(job_posting_text)
        resume_doc = nlp(resume_text)

        job_skills = set([ent.text.lower() for ent in job_doc.ents if ent.label_ == 'SKILL'])
        resume_skills = set([ent.text.lower() for ent in resume_doc.ents if ent.label_ == 'SKILL'])


        job_skills.update([skill.lower() for skill in PREDEFINED_SKILLS if skill.lower() in job_posting_text.lower()])
        resume_skills.update([skill.lower() for skill in PREDEFINED_SKILLS if skill.lower() in resume_text.lower()])


        job_skills = set(filter(lambda x: len(x) > 1, job_skills))
        resume_skills = set(filter(lambda x: len(x) > 1, resume_skills))

        matched_skills = job_skills.intersection(resume_skills)
        missing_skills = job_skills.difference(resume_skills)

        match_rate = (len(matched_skills) / len(job_skills)) * 100 if job_skills else 0
        score = calculate_resume_score(resume_skills, job_skills)
        progress_bar['value'] += 15

        sentiment_score_raw = analyze_sentiment(resume_text)
        sentiment_score = ((sentiment_score_raw + 1) / 2) * 4 + 1
        sentiment_score = round(sentiment_score)
        progress_bar['value'] += 10

        entities = [(ent.text, ent.label_) for ent in resume_doc.ents]
        progress_bar['value'] += 10
        vectorizer_lda = CountVectorizer(max_df=1.0, min_df=1, stop_words='english')  
        dtm = vectorizer_lda.fit_transform([resume_text])
        
        if dtm.shape[1] == 0:
            topics = []
            topic_words = []
            progress_bar['value'] += 10
        else:
            lda = LatentDirichletAllocation(n_components=2, random_state=42)
            lda.fit(dtm)
            topics = lda.components_
            topic_words = vectorizer_lda.get_feature_names_out()
            progress_bar['value'] += 10


        resume_category = job_category
        progress_bar['value'] += 10


        recommended_jobs = get_recommended_jobs(resume_category, matched_skills)
        progress_bar['value'] += 10


        result_text = f"Match Rate: {match_rate:.2f}%\n"
        result_text += f"Sentiment Score: {sentiment_score} (1 = Very Negative, 5 = Very Positive)\n"
        result_text += f"Resume Category: {resume_category}\n"
        result_text += f"Resume Score: {score} (Based on matched skills)\n\n"
        result_text += "Extracted Entities:\n"
        for ent, label in entities:
            result_text += f" - {ent} ({label})\n"
        result_text += "\nRecommended Jobs:\n"
        for job in recommended_jobs:
            result_text += f" - {job}\n"

        if topics.size > 0:
            result_text += "\nIdentified Topics:\n"
            for idx, topic in enumerate(topics):
                top_features_ind = topic.argsort()[:-6:-1]
                top_features = [topic_words[i] for i in top_features_ind]
                result_text += f"Topic {idx + 1}: {', '.join(top_features)}\n"


        insights = generate_insights(matched_skills, missing_skills, entities, match_rate, recommended_jobs)
        result_text += "\nInsights and Recommendations:\n"
        result_text += insights

        globals()['insights'] = insights

        enqueue_task(("showinfo", ("Analysis Complete", "Resume analysis completed successfully.")))
        enqueue_task(("update_insights", (insights,)))
        enqueue_task(("update_plot", (matched_skills, missing_skills)))

        progress_bar.stop()
        progress_bar['value'] = 100
    except Exception as e:
        enqueue_task(("showwarning", ("Analysis Error", f"An error occurred during analysis:\n{str(e)}")))
        progress_bar.stop()

def get_recommended_jobs(category, matched_skills):
    job_bank = {
        'Engineering': ['Software Engineer', 'Mechanical Engineer', 'Civil Engineer', 'Electrical Engineer', 'Chemical Engineer'],
        'Marketing': ['Digital Marketing Manager', 'Content Strategist', 'SEO Specialist', 'Brand Manager', 'Marketing Analyst'],
        'Sales': ['Sales Manager', 'Account Executive', 'Business Development Manager', 'Sales Representative', 'Territory Manager'],
        'Finance': ['Financial Analyst', 'Accountant', 'Investment Banker', 'Financial Planner', 'Controller'],
        'Human Resources': ['HR Manager', 'Recruiter', 'Talent Acquisition Specialist', 'HR Coordinator', 'Compensation and Benefits Manager'],
        'Information Technology': ['IT Manager', 'System Administrator', 'Network Engineer', 'Cybersecurity Analyst', 'IT Support Specialist'],
        'Healthcare': ['Registered Nurse', 'Medical Assistant', 'Healthcare Administrator', 'Physician Assistant', 'Pharmacist'],
        'Education': ['Teacher', 'Academic Advisor', 'Curriculum Developer', 'School Administrator', 'Instructional Coordinator'],
        'Operations': ['Operations Manager', 'Logistics Coordinator', 'Supply Chain Analyst', 'Procurement Specialist', 'Operations Analyst'],
        'Hospitality': ['Restaurant Manager', 'Chef', 'Barista', 'Event Coordinator', 'Catering Manager']
    }

    recommended = job_bank.get(category, [])
    return recommended


def generate_insights(matched_skills, missing_skills, entities, match_rate, recommended_jobs):
    insights = ""
    

    insights += "1. **Top Matched Skills:**\n"
    if matched_skills:
        for skill in matched_skills:
            insights += f"   - {skill.capitalize()}\n"
    else:
        insights += "   - No matching skills found.\n"

    insights += "\n2. **Skills Missing:**\n"
    if missing_skills:
        for skill in missing_skills:
            insights += f"   - {skill.capitalize()}\n"
    else:
        insights += "   - No missing skills. Excellent match!\n"

    education = [ent[0] for ent in entities if ent[1] in ['ORG', 'EDUCATION']]
    certifications = [ent[0] for ent in entities if ent[1] in ['CERT', 'WORK_OF_ART']]
    insights += "\n3. **Education and Certifications:**\n"
    if education or certifications:
        if education:
            insights += "   - Educational Institutions: " + ", ".join(education[:5]) + "\n"
        if certifications:
            insights += "   - Certifications: " + ", ".join(certifications[:5]) + "\n"
    else:
        insights += "   - No educational institutions or certifications identified.\n"
    

    additional_skills = matched_skills.union(missing_skills)
    insights += "\n4. **Additional Skills and Technologies:**\n"
    if additional_skills:
        insights += "   - " + ", ".join(list(additional_skills)[:5]) + "\n"
    else:
        insights += "   - No additional skills or technologies identified.\n"
    

    insights += "\n5. **Recommended Jobs:**\n"
    if recommended_jobs:
        for job in recommended_jobs:
            insights += f"   - {job}\n"
    else:
        insights += "   - No recommended jobs available.\n"
    

    insights += "\n6. **Overall Recommendation:**\n"
    if match_rate > 70:
        insights += "   - Your resume is well-aligned with the job posting. Consider applying!\n"
    elif match_rate > 40:
        insights += "   - Your resume matches some key aspects of the job posting. Consider tailoring it further to highlight missing skills.\n"
    else:
        insights += "   - Your resume has limited alignment with the job posting. Consider acquiring the missing skills and revising your resume to better match the job requirements.\n"
    
    return insights


def update_visualization(matched_skills, missing_skills):

    vis_window = tk.Toplevel(root)
    vis_window.title("Visualization")
    vis_window.geometry("1000x600")
    vis_window.grab_set()


    notebook = ttk.Notebook(vis_window)
    notebook.pack(expand=True, fill='both')


    pie_frame = ttk.Frame(notebook)
    notebook.add(pie_frame, text='Skills Distribution')


    bar_frame = ttk.Frame(notebook)
    notebook.add(bar_frame, text='Top Skills')


    labels = ['Matched Skills', 'Missing Skills']
    sizes = [len(matched_skills), len(missing_skills)]
    colors = ['#2ecc71', '#e74c3c']
    explode = (0.1, 0)

    fig1, ax1 = plt.subplots(figsize=(6, 6))
    ax1.pie(sizes, explode=explode, labels=labels, colors=colors,
            autopct='%1.1f%%', shadow=True, startangle=140)
    ax1.axis('equal')
    ax1.set_title('Skills Distribution', fontsize=16)

    canvas1 = FigureCanvasTkAgg(fig1, master=pie_frame)
    canvas1.draw()
    canvas1.get_tk_widget().pack(expand=True, fill='both')

    if matched_skills:
        skill_freq = {skill: 1 for skill in matched_skills}

        word_freq = pd.Series(skill_freq).sort_values(ascending=False).head(10)

        fig2, ax2 = plt.subplots(figsize=(8, 6))
        sns.barplot(x=word_freq.values, y=word_freq.index, ax=ax2, palette='viridis')
        ax2.set_title('Top Matched Skills', fontsize=16)
        ax2.set_xlabel('Frequency', fontsize=14)
        ax2.set_ylabel('Skills', fontsize=14)
        plt.tight_layout()

        canvas2 = FigureCanvasTkAgg(fig2, master=bar_frame)
        canvas2.draw()
        canvas2.get_tk_widget().pack(expand=True, fill='both')
    else:
        no_data_label = ttk.Label(bar_frame, text="No matched skills to display.", font=("Arial", 14))
        no_data_label.pack(pady=20)


    close_button = ttk.Button(vis_window, text="Close", command=vis_window.destroy)
    close_button.pack(pady=10)


def update_insights(insights):
    insights_textarea.config(state='normal')
    insights_textarea.delete("1.0", "end")
    insights_textarea.insert("1.0", insights)
    insights_textarea.config(state='disabled')


def analyze_sentiment(text):
    sentiment = sia.polarity_scores(text)
    return sentiment['compound']

def generate_report_with_gemini(name, contact, insights, match_rate, resume_category, score):
    GENIUS_API_KEY = os.getenv("GENIUS_API_KEY")
    if not GENIUS_API_KEY:
        messagebox.showerror("API Key Error", "Gemini API key not found. Please set the GENIUS_API_KEY environment variable.")
        return None


    try:
        import google.generativeai as genai
    except ImportError:
        messagebox.showerror("Import Error", "google.generativeai package is not installed.")
        return None

    genai.configure(api_key=GENIUS_API_KEY)

    model = genai.GenerativeModel("gemini-1.5-flash")
    prompt = f"""
    Generate a professional improvement report for a CV based on the following details:

    Name: {name}
    Contact Information: {contact}

    Job Description Insights:
    {insights}

    Match Rate: {match_rate:.2f}%
    Resume Category: {resume_category}
    Resume Score: {score}

    Provide detailed suggestions to improve the CV, focusing on areas with low match rates, missing skills, and enhancing the presentation of existing skills and experiences.
    """

    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        print(f"Gemini API Error: {e}")
        return None


def generate_text():
    if not result_text:
        messagebox.showerror("Generate Error", "No analysis results available. Please analyze a resume first.")
        return


    global match_rate, resume_category, score, insights


    if not insights:
        messagebox.showerror("Generate Error", "No insights available to generate a report.")
        return


    def submit_improvement_report():
        name = name_entry.get().strip()
        contact = contact_entry.get().strip()
        if not name or not contact:
            messagebox.showerror("Input Error", "Please enter both your name and contact information.")
            return

        report = generate_report_with_gemini(name, contact, insights, match_rate, resume_category, score)
        if report:

            report_window = tk.Toplevel(root)
            report_window.title("Improvement Report")
            report_window.geometry("800x600")
            report_window.grab_set()

            report_textarea = scrolledtext.ScrolledText(report_window, width=100, height=35, bg="white", relief="flat",
                                                       font=("Arial", 12))
            report_textarea.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
            report_textarea.insert("1.0", report)
            report_textarea.config(state='disabled')


            close_button = ttk.Button(report_window, text="Close", command=report_window.destroy)
            close_button.pack(pady=10)
        else:
            messagebox.showerror("Generation Error", "Failed to generate report.")


    report_input_window = tk.Toplevel(root)
    report_input_window.title("Generate Improvement Report")
    report_input_window.geometry("400x300")
    report_input_window.grab_set()

    tk.Label(report_input_window, text="Name:", font=("Arial", 12)).pack(pady=10)
    name_entry = ttk.Entry(report_input_window, width=50, font=("Arial", 12))
    name_entry.pack(pady=5)

    tk.Label(report_input_window, text="Contact Information:", font=("Arial", 12)).pack(pady=10)
    contact_entry = ttk.Entry(report_input_window, width=50, font=("Arial", 12))
    contact_entry.pack(pady=5)

    submit_button = ttk.Button(report_input_window, text="Generate Report", command=submit_improvement_report)
    submit_button.pack(pady=20)


def save_config_on_exit():

    config = {
        "keyword_weights": keyword_weights
    }
    with open('config.json', 'w') as f:
        json.dump(config, f)
    root.destroy()


root = tk.Tk()
root.title("Advanced ATS Scanner")
root.geometry("1200x800")
root.configure(bg="#f1f2f6")


header_font = tkfont.Font(family="Helvetica", size=16, weight="bold")
subheader_font = tkfont.Font(family="Helvetica", size=12, weight="bold")


container = ttk.Frame(root)
container.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")


root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
container.columnconfigure(0, weight=3)
container.columnconfigure(1, weight=2)
container.rowconfigure(0, weight=1)


main_canvas = tk.Canvas(container, bg="#f1f2f6")
main_canvas.grid(row=0, column=0, sticky="nsew")


scrollbar_y = ttk.Scrollbar(container, orient=tk.VERTICAL, command=main_canvas.yview)
scrollbar_y.grid(row=0, column=1, sticky="ns")

scrollbar_x = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=main_canvas.xview)
scrollbar_x.grid(row=1, column=0, sticky="ew")

main_canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)


def on_configure(event):
    main_canvas.configure(scrollregion=main_canvas.bbox("all"))

main_canvas.bind('<Configure>', on_configure)


frame = ttk.Frame(main_canvas)
main_canvas.create_window((0, 0), window=frame, anchor='nw')


frame.columnconfigure(0, weight=3)
frame.columnconfigure(1, weight=2)
frame.rowconfigure(5, weight=1)


job_posting_label = ttk.Label(frame, text="Job Posting:", font=header_font, foreground="#34495e")
job_posting_label.grid(row=0, column=0, padx=10, pady=(0, 10), sticky="w")

job_posting_textarea = scrolledtext.ScrolledText(frame, width=60, height=15, bg="white", relief="flat",
                                               font=("Arial", 12))
job_posting_textarea.grid(row=1, column=0, padx=10, pady=(0, 20), sticky="w")


resume_label = ttk.Label(frame, text="Resume:", font=header_font, foreground="#34495e")
resume_label.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="w")

resume_textarea = scrolledtext.ScrolledText(frame, width=60, height=15, bg="white", relief="flat",
                                           font=("Arial", 12))
resume_textarea.grid(row=3, column=0, padx=10, pady=(0, 20), sticky="w")


button_frame = ttk.Frame(frame)
button_frame.grid(row=4, column=0, padx=10, pady=10, sticky="w")

analyze_button = ttk.Button(button_frame, text="Analyze Resume", command=analyze_resume)
analyze_button.grid(row=0, column=0, padx=5)

upload_button = ttk.Button(button_frame, text="Upload Document", command=upload_document)
upload_button.grid(row=0, column=1, padx=5)

save_button = ttk.Button(button_frame, text="Save Results", command=save_analysis_results)
save_button.grid(row=0, column=2, padx=5)

generate_report_button = ttk.Button(button_frame, text="Generate Improvement Report", command=generate_text)
generate_report_button.grid(row=0, column=3, padx=5)

show_visual_button = ttk.Button(button_frame, text="Show Visualization", command=show_visualization)
show_visual_button.grid(row=0, column=4, padx=5)

insights_label = ttk.Label(frame, text="Insights and Recommendations:", font=header_font, foreground="#34495e")
insights_label.grid(row=0, column=1, padx=10, pady=(0, 10), sticky="w")

insights_textarea = scrolledtext.ScrolledText(frame, width=60, height=25, bg="white", relief="flat",
                                            font=("Arial", 12), state='disabled')
insights_textarea.grid(row=1, column=1, rowspan=5, padx=10, pady=(0, 20), sticky="n")


resume_textarea.tag_configure("highlight", background="yellow")


style = ttk.Style()
style.theme_use('clam')
style.configure("TButton",
                foreground="white",
                background="#2ecc71",
                font=("Helvetica", 10, "bold"))
style.map("TButton",
          background=[('active', '#27ae60')])


root.protocol("WM_DELETE_WINDOW", save_config_on_exit)
root.mainloop()
