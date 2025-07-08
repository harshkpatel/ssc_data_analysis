"""
Dashboard Utilities Module
Provides utility functions for the analytics dashboard (overall stats, timeline, etc).
"""

import pandas as pd

class DashboardUtils:
    def __init__(self):
        pass

    def get_overall_statistics(self, df):
        """
        Get overall statistics for the dashboard.
        Returns a dict with total emails, date range, and other summary stats.
        """
        stats = {
            'total_emails': len(df),
            'date_range': [str(df['received'].min().date()), str(df['received'].max().date())],
        }
        if 'sender' in df.columns:
            stats['unique_senders'] = df['sender'].nunique()
        return stats

    def get_email_volume_timeline(self, df, period='days'):
        """
        Get email volume over time (days, weeks, or months).
        Returns a dict with time labels and email counts.
        """
        if period == 'days':
            df['time_group'] = df['received'].dt.date
        elif period == 'weeks':
            df['time_group'] = df['received'].dt.to_period('W').astype(str)
        elif period == 'months':
            df['time_group'] = df['received'].dt.to_period('M').astype(str)
        else:
            df['time_group'] = df['received'].dt.date
        grouped = df.groupby('time_group').size().reset_index(name='count')
        return {'time': grouped['time_group'].astype(str).tolist(), 'count': grouped['count'].tolist()} 