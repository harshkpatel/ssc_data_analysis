"""
Sentiment Analyzer Module (VADER version)
Performs sentiment analysis on student emails and provides statistics for the dashboard.
"""

import pandas as pd
import nltk
from nltk.sentiment import SentimentIntensityAnalyzer

def ensure_vader_lexicon():
    """Ensure the VADER lexicon is downloaded."""
    try:
        nltk.data.find('sentiment/vader_lexicon.zip')
    except LookupError:
        nltk.download('vader_lexicon')

class SentimentAnalyzer:
    def __init__(self):
        self.analyzer = SentimentIntensityAnalyzer()

    def analyze_sentiment(self, text):
        """Return compound polarity for a given text (VADER)."""
        score = self.analyzer.polarity_scores(str(text))
        return score['compound']

    def get_sentiment_over_time(self, df, period='days'):
        """
        Get average sentiment polarity over time (days, weeks, or months).
        Returns a dict with time labels and average polarity.
        """
        df = df.copy()
        df['polarity'] = df['content'].apply(self.analyze_sentiment)
        if period == 'days':
            df['time_group'] = df['received'].dt.date
        elif period == 'weeks':
            df['time_group'] = df['received'].dt.to_period('W').astype(str)
        elif period == 'months':
            df['time_group'] = df['received'].dt.to_period('M').astype(str)
        else:
            df['time_group'] = df['received'].dt.date
        grouped = df.groupby('time_group')['polarity'].mean().reset_index()
        # Fill missing dates with 0
        all_times = pd.date_range(df['received'].min(), df['received'].max(), freq={'days':'D','weeks':'W','months':'M'}[period])
        grouped = grouped.set_index('time_group').reindex(all_times, fill_value=0).reset_index()
        grouped.columns = ['time', 'avg_polarity']
        return grouped.to_dict(orient='list')

    def get_sentiment_distribution(self, df):
        """
        Get sentiment distribution (positive/neutral/negative) for the dataset.
        Returns a dict with counts for each category.
        """
        df = df.copy()
        df['polarity'] = df['content'].apply(self.analyze_sentiment)
        # VADER: compound > 0.05 = positive, < -0.05 = negative, else neutral
        def label(p):
            if p > 0.05:
                return 'Positive'
            elif p < -0.05:
                return 'Negative'
            else:
                return 'Neutral'
        df['sentiment'] = df['polarity'].apply(label)
        counts = df['sentiment'].value_counts().to_dict()
        return counts 