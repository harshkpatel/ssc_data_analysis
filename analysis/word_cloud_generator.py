"""
Word Cloud Generator Module
Generates word clouds and extracts top keywords from student emails for the dashboard.
"""

import matplotlib.pyplot as plt
from wordcloud import WordCloud
import base64
import io
from collections import Counter
import pandas as pd
import re

class WordCloudGenerator:
    def __init__(self):
        pass

    def generate_word_cloud(self, text):
        """
        Generate a word cloud from the given text and return as base64 PNG and top keywords.
        """
        wordcloud = WordCloud(width=800, height=400, background_color='white', collocations=False).generate(text)
        img = io.BytesIO()
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.tight_layout(pad=0)
        plt.savefig(img, format='png')
        plt.close()
        img.seek(0)
        img_b64 = base64.b64encode(img.read()).decode('utf-8')
        # Get top keywords
        words = re.findall(r'\w+', text.lower())
        top_keywords = Counter(words).most_common(20)
        return {'wordcloud': img_b64, 'top_keywords': top_keywords}

    def generate_classified_word_cloud(self, classified_df):
        """
        Generate a word cloud from classified email categories (category names weighted by count).
        """
        cat_counts = classified_df['category'].value_counts()
        text = ' '.join([cat + ' ' for cat, count in cat_counts.items() for _ in range(count)])
        return self.generate_word_cloud(text)

    def get_top_keywords(self, content_series):
        """
        Get top keywords from a pandas Series of email content.
        """
        all_text = ' '.join(content_series.astype(str))
        words = re.findall(r'\w+', all_text.lower())
        return dict(Counter(words).most_common(20)) 