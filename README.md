### **ATS Scanner - Quick Start Guide**

Welcome to the **ATS Scanner**! This tool helps you analyze and optimize your resume to match specific job postings effectively.

---

#### **1. Installation Requirements**

Ensure you have the following installed:

- **Python 3.x**: [Download Python](https://www.python.org/downloads/)
- **Python Packages**: Install required packages using:

  ```bash
  pip install tkinter pillow pandas numpy matplotlib seaborn spacy scikit-learn nltk pytesseract textract PyPDF2 pythoncom pywin32 google-generativeai
  ```

- **spaCy Model**:

  ```bash
  python -m spacy download en_core_web_sm
  ```

- **Tesseract OCR**:
  - **Windows**: [Download Tesseract](https://github.com/tesseract-ocr/tesseract/wiki)
  - **macOS**: `brew install tesseract`
  - **Linux**: `sudo apt-get install tesseract-ocr`

- **Gemini API Key**: Set your Gemini API key as an environment variable `GENIUS_API_KEY`.

---

#### **2. Running the Application**

1. **Download the Script**: Save `main.py` to your computer.
2. **Navigate to Directory**:

   ```bash
   cd path_to_your_directory
   ```

3. **Launch the App**:

   ```bash
   python main.py
   ```

---

#### **3. Using the Application**

- **Upload Your Resume**:
  - Click **"Upload Document"**.
  - Select your resume file (`.docx`, `.pdf`, `.txt`, `.rtf`, `.doc`, `.odt`, `.jpg`, `.jpeg`, `.png`).
  - Extracted text appears in the **"Resume"** section.

- **Enter Job Posting**:
  - Paste or type the job description into the **"Job Posting"** section.

- **Analyze Resume**:
  - Click **"Analyze Resume"**.
  - Wait for the analysis to complete.
  - View insights in **"Insights and Recommendations"**.

- **View Visualizations**:
  - Click **"Show Visualization"** to see charts on skills distribution and top matched skills.

- **Save Results**:
  - Click **"Save Results"** to export your analysis as `.txt` or `.json`.

- **Generate Improvement Report**:
  - Click **"Generate Improvement Report"**.
  - Enter your **Name** and **Contact Information**.
  - Receive a personalized report with enhancement suggestions.

---

#### **4. Exiting the Application**

- Close the window. Your settings are saved automatically.

---

#### **5. Tips**

- **Clear Text**: Ensure your resume text is clear, especially for image uploads, to improve OCR accuracy.
- **Update Regularly**: Keep Python packages and spaCy models up-to-date for optimal performance.
- **Secure API Keys**: Keep your `GENIUS_API_KEY` confidential.

---

#### **6. Troubleshooting**

- **Missing spaCy Model**: Run `python -m spacy download en_core_web_sm`.
- **Tesseract Errors**: Ensure Tesseract is installed and added to your system PATH.
- **API Issues**: Verify `GENIUS_API_KEY` is correctly set and you have an active internet connection.
- **Unsupported Files**: Use supported file formats listed above.

---

Thank you for using the **ATS Scanner**! Optimize your resume and take a step closer to your desired job.