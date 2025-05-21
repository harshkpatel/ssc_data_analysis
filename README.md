# ssc_data_analysis

## 📚 Project Overview

This project analyzes student email data using text analysis and machine learning techniques. It is designed for users with no prior coding experience. You will learn how to set up your computer, install the necessary tools, and run the analysis to gain insights from student emails.

---

## 🖥️ Step-by-Step Instructions for Beginners

### 1. **Install Anaconda (Recommended for Beginners)**
Anaconda is a free, easy-to-use Python distribution that works on Windows, Mac, and Linux.

- Go to [Anaconda Downloads](https://www.anaconda.com/products/distribution#download-section)
- Download the installer for your operating system (Windows, Mac, or Linux)
- Run the installer and follow the on-screen instructions (choose default options)

### 2. **Open Anaconda Prompt or Terminal**
- **Windows:** Open "Anaconda Prompt" from the Start Menu
- **Mac/Linux:** Open the Terminal app

### 3. **Download the Project Files**
- Download or clone the project folder `ssc_data_analysis` to a location on your computer.
- If you received a ZIP file, right-click and select "Extract All".

### 4. **Navigate to the Project Folder**
In your Anaconda Prompt or Terminal, type:
```sh
cd path/to/ssc_data_analysis
```
Replace `path/to/ssc_data_analysis` with the actual path to the folder on your computer.

### 5. **Create a Python Environment**
This keeps your project dependencies separate and avoids conflicts.
```sh
conda create -n ssc_env python=3.9 -y
```

### 6. **Activate the Environment**
```sh
conda activate ssc_env
```

### 7. **Install Required Packages**
Install the main packages using conda:
```sh
conda install pandas numpy nltk textblob gensim scikit-learn -y
```
Then install the remaining package with pip:
```sh
pip install rake-nltk
```

### 8. **Download NLTK Data**
Run the following command to download necessary language data for text analysis:
```sh
python download_nltk_data.py
```
You should see messages indicating that NLTK data is being downloaded.

### 9. **Run the Text Analysis Script**
To analyze the sample student emails, run:
```sh
python analysis/text_analysis.py
```
This will print:
- Sentiment analysis results
- Top keywords
- Topic modeling
- Email clustering
- Concordance analysis
- Answers to sample queries

### 10. **Understanding the Output**
- **Sentiment Analysis:** Shows the emotional tone of emails (positive, neutral, negative)
- **Keywords:** Important phrases found in the emails
- **Topic Modeling:** Main topics discussed in the emails
- **Clustering:** Groups of similar emails
- **Concordance:** How a word (like "course") is used in context
- **Query Answering:** Finds the most relevant email for a given question

---

## 🛠️ Troubleshooting
- If you see errors about missing packages, repeat steps 7 and 8.
- If you see a message about "command not found: conda", make sure Anaconda is installed and you are using the Anaconda Prompt or Terminal.
- If you get a permissions error, try running your terminal as an administrator (Windows) or use `sudo` (Mac/Linux).

---

## 📂 Project Structure
```
ssc_data_analysis/
├── analysis/
│   └── text_analysis.py         # Main analysis script
├── csv_files/
│   └── sample_student_emails.csv # Sample data
├── download_nltk_data.py        # Script to download NLTK data
├── requirements.txt             # List of required packages
└── README.md                    # This file
```

---

## ❓ Getting Help
- If you get stuck, search for the error message online or ask a friend for help.
- You can also visit [Stack Overflow](https://stackoverflow.com/) and describe your problem.

---

## 🎉 Congratulations!
You have run your first text analysis project! Feel free to modify the sample data or scripts to explore more.