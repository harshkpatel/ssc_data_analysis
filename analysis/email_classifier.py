"""
Email Classifier Module
Classifies student emails into different categories based on content analysis.
"""

import pandas as pd
import re
from collections import Counter
import numpy as np

class EmailClassifier:
    def __init__(self):
        """Initialize the email classifier with predefined categories and keywords."""
        self.categories = {
            'course_selection': {
                'keywords': [
                    'course', 'courses', 'registration', 'enroll', 'enrollment', 'mat137', 'mat157', 
                    'csc108', 'csc110', 'csc111', 'mat240', 'phy151', 'bio120', 'pump', 'program',
                    'requirement', 'prerequisite', 'credit', 'credits', 'specialist', 'major', 'minor',
                    'breadth', 'distribution', 'acorn', 'timetable', 'schedule', 'semester', 'fall', 'winter'
                ],
                'description': 'Course selection, registration, and academic planning'
            },
            'housing_residence': {
                'keywords': [
                    'housing', 'residence', 'dorm', 'dormitory', 'room', 'apartment', 'meal plan',
                    'dining', 'food', 'cafeteria', 'new college', 'university college', 'trinity',
                    'victoria college', 'st michael', 'innis', 'woodsworth', 'rent', 'lease',
                    'accommodation', 'living', 'roommate', 'suite', 'single room', 'double room'
                ],
                'description': 'Housing, residence, and accommodation inquiries'
            },
            'international_student': {
                'keywords': [
                    'international', 'visa', 'study permit', 'passport', 'immigration', 'sin',
                    'social insurance number', 'work permit', 'co-op', 'coop', 'work study',
                    'uhip', 'health insurance', 'overseas', 'foreign', 'country', 'home',
                    'travel', 'arrival', 'orientation', 'english', 'language', 'esl'
                ],
                'description': 'International student specific inquiries'
            },
            'academic_support': {
                'keywords': [
                    'tutoring', 'help', 'support', 'academic', 'study', 'learning', 'difficult',
                    'struggle', 'challenge', 'writing center', 'math aid', 'office hours',
                    'professor', 'instructor', 'ta', 'teaching assistant', 'mentor', 'advisor',
                    'advising', 'guidance', 'resource', 'library', 'study space', 'quiet'
                ],
                'description': 'Academic support and learning resources'
            },
            'financial_aid': {
                'keywords': [
                    'financial', 'aid', 'scholarship', 'bursary', 'loan', 'grant', 'money',
                    'cost', 'tuition', 'fee', 'payment', 'budget', 'expensive', 'afford',
                    'funding', 'award', 'prize', 'work study', 'part time', 'job', 'employment',
                    'income', 'expense', 'textbook', 'book', 'material', 'supply'
                ],
                'description': 'Financial aid, scholarships, and cost-related inquiries'
            },
            'campus_life': {
                'keywords': [
                    'club', 'organization', 'student life', 'activity', 'event', 'social',
                    'friend', 'network', 'community', 'campus', 'facility', 'gym', 'fitness',
                    'sport', 'recreation', 'entertainment', 'party', 'celebration', 'festival',
                    'concert', 'performance', 'art', 'culture', 'diversity', 'inclusion'
                ],
                'description': 'Campus life, clubs, and social activities'
            },
            'technology_it': {
                'keywords': [
                    'computer', 'laptop', 'software', 'program', 'coding', 'programming',
                    'internet', 'wifi', 'network', 'email', 'account', 'password', 'login',
                    'quercus', 'acorn', 'portal', 'online', 'digital', 'tech', 'technology',
                    'it', 'support', 'help desk', 'virus', 'security', 'backup'
                ],
                'description': 'Technology and IT support inquiries'
            },
            'health_wellness': {
                'keywords': [
                    'health', 'medical', 'doctor', 'nurse', 'clinic', 'hospital', 'sick',
                    'illness', 'injury', 'mental health', 'counseling', 'therapy', 'stress',
                    'anxiety', 'depression', 'wellness', 'fitness', 'exercise', 'nutrition',
                    'diet', 'sleep', 'wellbeing', 'covid', 'vaccine', 'test'
                ],
                'description': 'Health, wellness, and medical services'
            },
            'transportation': {
                'keywords': [
                    'transport', 'transportation', 'bus', 'subway', 'ttc', 'transit', 'train',
                    'car', 'parking', 'bike', 'bicycle', 'walk', 'walking', 'route', 'map',
                    'direction', 'location', 'address', 'street', 'avenue', 'road', 'drive'
                ],
                'description': 'Transportation and commuting inquiries'
            },
            'career_employment': {
                'keywords': [
                    'career', 'job', 'employment', 'work', 'internship', 'co-op', 'coop',
                    'placement', 'opportunity', 'position', 'hire', 'hiring', 'recruit',
                    'resume', 'cv', 'interview', 'application', 'apply', 'company', 'industry',
                    'professional', 'network', 'linkedin', 'experience', 'skill'
                ],
                'description': 'Career and employment opportunities'
            },
            'general_inquiry': {
                'keywords': [
                    'question', 'inquiry', 'information', 'help', 'assist', 'support',
                    'welcome', 'hello', 'hi', 'thank', 'thanks', 'appreciate', 'grateful',
                    'excited', 'nervous', 'anxious', 'worry', 'concern', 'confused', 'lost'
                ],
                'description': 'General inquiries and questions'
            }
        }
    
    def classify_emails(self, df):
        """
        Classify emails into categories based on content analysis.
        
        Args:
            df (pd.DataFrame): DataFrame with 'content' and 'subject' columns
            
        Returns:
            pd.DataFrame: DataFrame with added 'category' and 'confidence' columns
        """
        df_copy = df.copy()
        df_copy['category'] = 'uncategorized'
        df_copy['confidence'] = 0.0
        df_copy['category_keywords'] = ''
        
        for idx, row in df_copy.iterrows():
            # Combine subject and content for analysis
            text = f"{row['subject']} {row['content']}".lower()
            
            best_category = 'uncategorized'
            best_score = 0
            matched_keywords = []
            
            for category, config in self.categories.items():
                score = 0
                category_matches = []
                
                for keyword in config['keywords']:
                    # Count keyword occurrences
                    pattern = r'\b' + re.escape(keyword) + r'\b'
                    matches = len(re.findall(pattern, text))
                    if matches > 0:
                        score += matches
                        category_matches.append(keyword)
                
                # Normalize score by text length
                if len(text.split()) > 0:
                    score = score / len(text.split()) * 100
                
                if score > best_score:
                    best_score = score
                    best_category = category
                    matched_keywords = category_matches
            
            df_copy.at[idx, 'category'] = best_category
            df_copy.at[idx, 'confidence'] = best_score
            df_copy.at[idx, 'category_keywords'] = ', '.join(matched_keywords)
        
        return df_copy
    
    def get_category_statistics(self, df):
        """
        Get statistics for email categories.
        
        Args:
            df (pd.DataFrame): DataFrame with email data
            
        Returns:
            dict: Category statistics
        """
        classified_df = self.classify_emails(df)
        
        # Category counts
        category_counts = classified_df['category'].value_counts().to_dict()
        
        # Category descriptions
        category_descriptions = {
            category: config['description'] 
            for category, config in self.categories.items()
        }
        
        # Average confidence by category
        confidence_by_category = classified_df.groupby('category')['confidence'].mean().to_dict()
        
        # Top keywords by category
        top_keywords_by_category = {}
        for category in self.categories.keys():
            category_emails = classified_df[classified_df['category'] == category]
            if not category_emails.empty:
                all_keywords = []
                for keywords in category_emails['category_keywords']:
                    if keywords:
                        all_keywords.extend([kw.strip() for kw in keywords.split(',')])
                
                keyword_counts = Counter(all_keywords)
                top_keywords_by_category[category] = dict(keyword_counts.most_common(5))
        
        return {
            'category_counts': category_counts,
            'category_descriptions': category_descriptions,
            'confidence_by_category': confidence_by_category,
            'top_keywords_by_category': top_keywords_by_category,
            'total_emails': len(df),
            'categorized_emails': len(classified_df[classified_df['category'] != 'uncategorized'])
        }
    
    def get_category_timeline(self, df, time_period='days'):
        """
        Get email volume by category over time.
        
        Args:
            df (pd.DataFrame): DataFrame with email data
            time_period (str): Time period for grouping ('days', 'weeks', 'months')
            
        Returns:
            dict: Timeline data by category
        """
        classified_df = self.classify_emails(df)
        
        # Group by time period
        if time_period == 'days':
            classified_df['time_group'] = classified_df['received'].dt.date
        elif time_period == 'weeks':
            classified_df['time_group'] = classified_df['received'].dt.to_period('W')
        elif time_period == 'months':
            classified_df['time_group'] = classified_df['received'].dt.to_period('M')
        
        # Get timeline for each category
        timeline_data = {}
        for category in self.categories.keys():
            category_emails = classified_df[classified_df['category'] == category]
            if not category_emails.empty:
                timeline = category_emails.groupby('time_group').size().to_dict()
                timeline_data[category] = timeline
        
        return timeline_data 