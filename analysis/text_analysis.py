import pandas as pd
import numpy as np
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
from nltk import pos_tag
from textblob import TextBlob
from rake_nltk import Rake
from gensim import corpora
from gensim.models.ldamodel import LdaModel
from sklearn.feature_extraction.text import TfidfVectorizer, CountVectorizer
from sklearn.cluster import KMeans
from sklearn.metrics import pairwise_distances
from sklearn.decomposition import TruncatedSVD
import string
import re
import os

# Download required NLTK data
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('wordnet')
nltk.download('averaged_perceptron_tagger')

class StudentEmailAnalyzer:
    def __init__(self, csv_path):
        """Initialize the analyzer with the path to the CSV file."""
        self.df = pd.read_csv(csv_path)
        self.stop_words = set(stopwords.words('english'))
        # Add UofT-specific terms to stop words
        self.stop_words.update(['uoft', 'toronto', 'university', 'college', 'campus'])
        self.lemmatizer = WordNetLemmatizer()
        self.exclude = set(string.punctuation)
        
    def clean_text(self, text):
        """Clean text by removing stopwords and punctuation."""
        # Convert to lowercase and split
        words = text.lower().split()
        # Remove stopwords
        stop_free = " ".join([i for i in words if i not in self.stop_words])
        # Remove punctuation
        punc_free = ''.join(ch for ch in stop_free if ch not in self.exclude)
        return punc_free
    
    def lemmatize_text(self, text):
        """Lemmatize text using NLTK's WordNetLemmatizer."""
        empty = []
        for word, tag in pos_tag(word_tokenize(text)):
            wntag = tag[0].lower()
            wntag = wntag if wntag in ['a', 'r', 'n', 'v'] else None
            if not wntag:
                lemma = word
                empty.append(lemma)
            else:
                lemma = self.lemmatizer.lemmatize(word, wntag)
                empty.append(lemma)
        return ' '.join(empty)
    
    def perform_sentiment_analysis(self):
        """Perform sentiment analysis on the email content."""
        # Use TextBlob for sentiment analysis
        self.df['polarity'] = self.df['content'].apply(lambda x: TextBlob(x).sentiment.polarity)
        self.df['subjectivity'] = self.df['content'].apply(lambda x: TextBlob(x).sentiment.subjectivity)
        
        # Add sentiment categories
        self.df['sentiment_category'] = pd.cut(
            self.df['polarity'],
            bins=[-1, -0.1, 0.1, 1],
            labels=['Negative', 'Neutral', 'Positive']
        )
        
        return self.df[['subject', 'polarity', 'subjectivity', 'sentiment_category']]
    
    def extract_keywords(self, text):
        """Extract keywords using RAKE."""
        r = Rake(include_repeated_phrases=False, min_length=1, max_length=3)
        r.extract_keywords_from_text(text)
        keyword_rank = [keyword for keyword in r.get_ranked_phrases_with_scores() if keyword[0] > 5]
        return [keyword[1] for keyword in keyword_rank]
    
    def perform_topic_modeling(self, num_topics=5, num_words=4):
        """Perform topic modeling using LDA on all emails."""
        # Clean and lemmatize all emails
        cleaned_texts = [self.clean_text(text) for text in self.df['content']]
        lemmatized_texts = [self.lemmatize_text(text) for text in cleaned_texts]
        
        # Create dictionary and corpus
        dictionary = corpora.Dictionary([text.split() for text in lemmatized_texts])
        corpus = [dictionary.doc2bow(text.split()) for text in lemmatized_texts]
        
        # Train LDA model
        ldamodel = LdaModel(corpus, num_topics=num_topics, id2word=dictionary, passes=50, random_state=1)
        
        # Get topics
        topics = ldamodel.print_topics(num_topics=num_topics, num_words=num_words)
        topics_list = []
        for elm in topics:
            content = elm[1]
            no_digits = ''.join([i for i in content if not i.isdigit()])
            topics_list.append(re.findall(r'\w+', no_digits, flags=re.IGNORECASE))
        
        return topics_list
    
    def cluster_emails(self, n_clusters=5):
        """Cluster emails using K-means."""
        cleaned_texts = [self.clean_text(email) for email in self.df['content']]
        lemmatized_texts = [self.lemmatize_text(text) for text in cleaned_texts]
        tfidf = TfidfVectorizer(stop_words=list(self.stop_words))
        X = tfidf.fit_transform(lemmatized_texts).toarray()
        kmeans = KMeans(n_clusters=n_clusters, random_state=42)
        clusters = kmeans.fit_predict(X)
        self.df['cluster'] = clusters
        cluster_stats = self.df.groupby('cluster').size()
        # Extract top keywords for each cluster
        cluster_keywords = {}
        for i in range(n_clusters):
            cluster_texts = ' '.join(self.df[self.df['cluster'] == i]['content'])
            r = Rake(stopwords=list(self.stop_words))
            r.extract_keywords_from_text(cluster_texts)
            cluster_keywords[i] = r.get_ranked_phrases()[:5]
        return cluster_stats, cluster_keywords
    
    def analyze_concordance(self, word, lines=10):
        """Analyze concordance for a specific word."""
        text = nltk.Text(word_tokenize(' '.join(self.df['content'])))
        return text.concordance(word, lines=lines)
    
    def answer_query(self, query):
        """Answer a query using similarity search."""
        # Clean and lemmatize all texts
        cleaned_texts = [self.clean_text(text) for text in self.df['content']]
        lemmatized_texts = [self.lemmatize_text(text) for text in cleaned_texts]
        
        # Create document-term matrix
        dtm = CountVectorizer(max_df=0.7, min_df=5, token_pattern="[a-z']+",
                            stop_words=list(self.stop_words), max_features=6000)
        dtm_mat = dtm.fit_transform(lemmatized_texts)
        
        # Apply SVD for dimensionality reduction
        n_features = dtm_mat.shape[1]
        n_components = min(200, n_features - 1) if n_features > 1 else 1
        tsvd = TruncatedSVD(n_components=n_components)
        tsvd_mat = tsvd.fit_transform(dtm_mat)
        
        # Transform query
        query_mat = tsvd.transform(dtm.transform([self.clean_text(query)]))
        
        # Find most similar document
        dist = pairwise_distances(X=tsvd_mat, Y=query_mat, metric='cosine')
        most_similar_idx = np.argmin(dist.flatten())
        
        return {
            'answer': self.df['content'].iloc[most_similar_idx],
            'subject': self.df['subject'].iloc[most_similar_idx],
            'similarity_score': 1 - dist.flatten()[most_similar_idx]
        }

def main():
    # Initialize analyzer with the path to the CSV file
    analyzer = StudentEmailAnalyzer('csv_files/sample_student_emails.csv')
    
    # Perform sentiment analysis
    print("\nSentiment Analysis Results:")
    sentiment_results = analyzer.perform_sentiment_analysis()
    print("\nSentiment Distribution:")
    print(sentiment_results['sentiment_category'].value_counts())
    print("\nSample of Sentiment Analysis:")
    print(sentiment_results.head())
    
    # Extract keywords from all emails
    print("\nTop Keywords Across All Emails:")
    all_text = ' '.join(analyzer.df['content'])
    keywords = analyzer.extract_keywords(all_text)
    print(keywords[:10])
    
    # Perform topic modeling
    print("\nTopic Modeling Results:")
    topics = analyzer.perform_topic_modeling()
    for i, topic in enumerate(topics):
        print(f"Topic {i+1}: {', '.join(topic)}")
    
    # Cluster emails
    print("\nEmail Clustering Results:")
    cluster_stats, cluster_keywords = analyzer.cluster_emails()
    print("\nCluster Statistics:")
    print(cluster_stats)
    print("\nTop Keywords by Cluster:")
    for cluster, keywords in cluster_keywords.items():
        print(f"Cluster {cluster}: {', '.join(keywords)}")
    
    # Analyze concordance for common words
    print("\nConcordance Analysis for 'course':")
    analyzer.analyze_concordance("course")
    
    # Answer sample queries
    print("\nQuery Answering Examples:")
    queries = [
        "How do I register for courses?",
        "What housing options are available?",
        "How do I get a student ID card?",
        "What are the computer science program requirements?"
    ]
    
    for query in queries:
        print(f"\nQuery: {query}")
        result = analyzer.answer_query(query)
        print(f"Answer: {result['answer']}")
        print(f"Subject: {result['subject']}")
        print(f"Similarity Score: {result['similarity_score']:.2f}")

if __name__ == "__main__":
    main() 