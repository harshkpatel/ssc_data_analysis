"""
Email Analytics Dashboard - Flask Application
A comprehensive dashboard for analyzing student email data with sentiment analysis,
word clouds, and email classification.
"""

from flask import Flask, jsonify, request, render_template
import pandas as pd
import numpy as np
import json
from datetime import datetime, timedelta
import plotly.graph_objs as go
import plotly.utils
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Use non-interactive backend
import io
import base64
from collections import Counter
import re
from textblob import TextBlob
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
import warnings
warnings.filterwarnings('ignore')

# Download required NLTK data
try:
    nltk.data.find('tokenizers/punkt')
except LookupError:
    nltk.download('punkt')
try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords')

from analysis.email_classifier import EmailClassifier
from analysis.sentiment_analyzer import SentimentAnalyzer
from analysis.word_cloud_generator import WordCloudGenerator
from analysis.dashboard_utils import DashboardUtils
from utils.sqlite_storage import get_all_emails

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'

# Initialize analyzers
email_classifier = EmailClassifier()
sentiment_analyzer = SentimentAnalyzer()
word_cloud_generator = WordCloudGenerator()
dashboard_utils = DashboardUtils()

def load_email_data(stream_filter=None):
    """Load and preprocess email data from CSV files and SQLite database."""
    try:
        # Load from SQLite database first
        db_emails = get_all_emails()
        if db_emails:
            df_db = pd.DataFrame(db_emails, columns=['subject', 'content', 'received', 'stream'])
            df_db['received'] = pd.to_datetime(df_db['received'])
            df_db['content'] = df_db['content'].fillna('')
            df_db['subject'] = df_db['subject'].fillna('')
            
            # Filter by stream if specified
            if stream_filter:
                df_db = df_db[df_db['stream'] == stream_filter]
        else:
            df_db = pd.DataFrame()
        
        # Load from fake CSV file as primary data source
        try:
            df_csv = pd.read_csv('csv_files/fake_uoft_emails.csv')
            df_csv['received'] = pd.to_datetime(df_csv['received'])
            df_csv['content'] = df_csv['content'].fillna('')
            df_csv['subject'] = df_csv['subject'].fillna('')
            # Add stream column if not present
            if 'stream' not in df_csv.columns:
                df_csv['stream'] = 'Unknown'
        except Exception as e:
            print(f"Error loading CSV data: {e}")
            df_csv = pd.DataFrame()
        
        # Combine dataframes, prioritizing database data
        if not df_db.empty and not df_csv.empty:
            # Combine and remove duplicates
            df_combined = pd.concat([df_db, df_csv], ignore_index=True)
            df_combined = df_combined.drop_duplicates(subset=['subject', 'content', 'received'])
        elif not df_db.empty:
            df_combined = df_db
        elif not df_csv.empty:
            df_combined = df_csv
        else:
            return pd.DataFrame()
        
        return df_combined
        
    except Exception as e:
        print(f"Error loading email data: {e}")
        return pd.DataFrame()

@app.route('/')
def index():
    """Main dashboard page - return basic info about available endpoints."""
    return jsonify({
        "message": "Email Analytics Dashboard API",
        "available_endpoints": [
            "/api/overall_stats",
            "/api/sentiment_over_time",
            "/api/word_cloud", 
            "/api/classified_word_cloud",
            "/api/email_categories",
            "/api/sentiment_distribution",
            "/api/email_volume_timeline",
            "/api/top_keywords",
            "/api/available_streams"
        ],
        "usage": "Add ?stream=STREAM_NAME to filter by stream (MPS, LS, RC, SS, HUM, CS)"
    })

@app.route('/dashboard')
def dashboard_page():
    return render_template('dashboard.html')

@app.route('/api/overall_stats')
def overall_stats():
    """Get overall statistics for the dashboard."""
    stream_filter = request.args.get('stream')
    df = load_email_data(stream_filter)
    if df.empty:
        return jsonify({'error': 'No data available'})
    
    stats = dashboard_utils.get_overall_statistics(df)
    return jsonify(stats)

@app.route('/api/sentiment_over_time')
def sentiment_over_time():
    """Get sentiment analysis over time."""
    stream_filter = request.args.get('stream')
    df = load_email_data(stream_filter)
    if df.empty:
        return jsonify({'error': 'No data available'})
    
    time_period = request.args.get('period', 'days')  # days, weeks, months
    sentiment_data = sentiment_analyzer.get_sentiment_over_time(df, time_period)
    return jsonify(sentiment_data)

@app.route('/api/word_cloud')
def word_cloud():
    """Generate word cloud from email content."""
    stream_filter = request.args.get('stream')
    df = load_email_data(stream_filter)
    if df.empty:
        return jsonify({'error': 'No data available'})
    
    # Generate word cloud from content
    content_text = ' '.join(df['content'].astype(str))
    word_cloud_data = word_cloud_generator.generate_word_cloud(content_text)
    return jsonify(word_cloud_data)

@app.route('/api/classified_word_cloud')
def classified_word_cloud():
    """Generate word cloud from classified email categories."""
    stream_filter = request.args.get('stream')
    df = load_email_data(stream_filter)
    if df.empty:
        return jsonify({'error': 'No data available'})
    
    # Classify emails and generate word cloud
    classified_data = email_classifier.classify_emails(df)
    word_cloud_data = word_cloud_generator.generate_classified_word_cloud(classified_data)
    return jsonify(word_cloud_data)

@app.route('/api/email_categories')
def email_categories():
    """Get email classification statistics."""
    stream_filter = request.args.get('stream')
    df = load_email_data(stream_filter)
    if df.empty:
        return jsonify({'error': 'No data available'})
    
    categories = email_classifier.get_category_statistics(df)
    return jsonify(categories)

@app.route('/api/sentiment_distribution')
def sentiment_distribution():
    """Get sentiment distribution statistics."""
    stream_filter = request.args.get('stream')
    df = load_email_data(stream_filter)
    if df.empty:
        return jsonify({'error': 'No data available'})
    
    sentiment_dist = sentiment_analyzer.get_sentiment_distribution(df)
    return jsonify(sentiment_dist)

@app.route('/api/email_volume_timeline')
def email_volume_timeline():
    """Get email volume over time."""
    stream_filter = request.args.get('stream')
    df = load_email_data(stream_filter)
    if df.empty:
        return jsonify({'error': 'No data available'})
    
    timeline_data = dashboard_utils.get_email_volume_timeline(df)
    return jsonify(timeline_data)

@app.route('/api/top_keywords')
def top_keywords():
    """Get top keywords from email content."""
    stream_filter = request.args.get('stream')
    df = load_email_data(stream_filter)
    if df.empty:
        return jsonify({'error': 'No data available'})
    
    keywords = word_cloud_generator.get_top_keywords(df['content'])
    return jsonify(keywords)

@app.route('/api/available_streams')
def available_streams():
    """Get list of available streams in the database."""
    df = load_email_data()
    if df.empty:
        return jsonify({'streams': []})
    
    streams = df['stream'].unique().tolist() if 'stream' in df.columns else []
    return jsonify({'streams': streams})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001) 