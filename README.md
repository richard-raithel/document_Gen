# Automated Resume Customization Tool

## Overview

This Python script provides an automated workflow to customize a resume based on a specific job description using OpenAI's GPT models. The tool updates the first table in a `.docx` resume (typically containing "Experience" and "Projects") to align better with a given job description while preserving the structure, formatting, and header content.

---

## Features
- **Resume Customization**:
  - Automatically updates the "Experience" and "Projects" sections in the resume to match a job description.
  - Retains the original header and document structure.
  
- **OpenAI-Powered Updates**:
  - Uses OpenAI's GPT model to generate tailored updates for each row and cell in the table.

- **Preserves Formatting**:
  - Maintains the structure and style of the original resume, including table rows and cells.

- **Flexible Output**:
  - Saves the updated resume as a new `.docx` file with the employer name and job title in the filename.

---

## Requirements

### Prerequisites
- Python 3.8 or higher
- OpenAI API key
- Microsoft Word `.docx` file for the resume
- Job description in `.txt` format

### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/resume-customization.git
   cd resume-customization
