# ==========================================
# ml_analyzer.py - Machine Learning Module
# Shila Restaurant QFD Dashboard
# ==========================================

import pandas as pd
import numpy as np
from collections import Counter
import warnings
warnings.filterwarnings('ignore')

# ML Imports
try:
    from sklearn.model_selection import train_test_split, cross_val_score
    from sklearn.preprocessing import StandardScaler, LabelEncoder
    from sklearn.ensemble import RandomForestClassifier, IsolationForest, GradientBoostingClassifier
    from sklearn.cluster import KMeans, DBSCAN
    from sklearn.metrics import classification_report, confusion_matrix, accuracy_score, precision_recall_fscore_support
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.decomposition import PCA
    ML_AVAILABLE = True
except ImportError:
    ML_AVAILABLE = False

try:
    from mlxtend.frequent_patterns import apriori, association_rules
    from mlxtend.preprocessing import TransactionEncoder
    MLXTEND_AVAILABLE = True
except ImportError:
    MLXTEND_AVAILABLE = False

try:
    from imblearn.over_sampling import SMOTE
    IMBLEARN_AVAILABLE = True
except ImportError:
    IMBLEARN_AVAILABLE = False


class ShilaMLAnalyzer:
    """Machine Learning Analyzer for Shila QFD Dashboard"""
    
    def __init__(self, df, config_cols):
        """
        Initialize ML Analyzer
        
        Args:
            df: DataFrame with customer feedback data
            config_cols: Column configuration from config.py
        """
        self.df = df.copy()
        self.COLS = config_cols
        self.models = {}
        self.scalers = {}
        
    # ==========================================
    # 1. DETRACTOR PREDICTION MODEL
    # ==========================================
    
    def prepare_classification_features(self):
        """Prepare features for classification"""
        if not ML_AVAILABLE:
            return None, None, None
        
        df = self.df.copy()
        
        # Target: Is Detractor (NPS 0-6)
        nps_col = self.COLS.get('NPS')
        if not nps_col or nps_col not in df.columns:
            return None, None, None
        
        df['is_detractor'] = (df[nps_col] <= 6).astype(int)
        
        # Features
        features = []
        feature_names = []
        
        # Rating
        rating_col = self.COLS.get('RATING')
        if rating_col and rating_col in df.columns:
            features.append(df[rating_col].fillna(3))
            feature_names.append('rating')
        
        # Branch (encoded)
        branch_col = self.COLS.get('BRANCH')
        if branch_col and branch_col in df.columns:
            le = LabelEncoder()
            df['branch_encoded'] = le.fit_transform(df[branch_col].fillna('Unknown'))
            features.append(df['branch_encoded'])
            feature_names.append('branch')
        
        # Time features (if date column exists)
        date_col = self.COLS.get('DATE')
        if date_col and date_col in df.columns:
            try:
                df['date_parsed'] = pd.to_datetime(df[date_col], errors='coerce')
                df['day_of_week'] = df['date_parsed'].dt.dayofweek.fillna(0)
                df['hour'] = df['date_parsed'].dt.hour.fillna(12)
                features.append(df['day_of_week'])
                features.append(df['hour'])
                feature_names.extend(['day_of_week', 'hour'])
            except:
                pass
        
        # Issue flags
        weakness_col = self.COLS.get('WEAKNESS')
        if weakness_col and weakness_col in df.columns:
            # Count issues
            df['issue_count'] = df[weakness_col].fillna('').apply(
                lambda x: len(str(x).split('،')) if pd.notna(x) and str(x).strip() else 0
            )
            features.append(df['issue_count'])
            feature_names.append('issue_count')
            
            # Specific issue flags
            common_issues = ['کیفیت پایین غذا', 'عدم تناسب حجم و قیمت', 'تاخیر در ارسال', 
                          'زمان آماده سازی سفارش', 'بسته‌بندی نامناسب']
            for issue in common_issues:
                col_name = f'has_{issue[:10]}'
                df[col_name] = df[weakness_col].fillna('').str.contains(issue, na=False).astype(int)
                features.append(df[col_name])
                feature_names.append(col_name)
        
        # Strength flags
        strength_col = self.COLS.get('STRENGTH')
        if strength_col and strength_col in df.columns:
            df['strength_count'] = df[strength_col].fillna('').apply(
                lambda x: len(str(x).split('،')) if pd.notna(x) and str(x).strip() else 0
            )
            features.append(df['strength_count'])
            feature_names.append('strength_count')
        
        if not features:
            return None, None, None
        
        X = pd.concat(features, axis=1)
        X.columns = feature_names
        X = X.fillna(0)
        y = df['is_detractor']
        
        return X, y, feature_names
    
    def train_detractor_model(self):
        """Train model to predict detractors"""
        if not ML_AVAILABLE:
            return {'error': 'scikit-learn not installed'}
        
        X, y, feature_names = self.prepare_classification_features()
        
        if X is None:
            return {'error': 'Could not prepare features'}
        
        # Split data
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.2, random_state=42, stratify=y
        )
        
        # Scale features
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_test_scaled = scaler.transform(X_test)
        
        # Handle imbalanced data
        if IMBLEARN_AVAILABLE:
            try:
                smote = SMOTE(random_state=42)
                X_train_balanced, y_train_balanced = smote.fit_resample(X_train_scaled, y_train)
            except:
                X_train_balanced, y_train_balanced = X_train_scaled, y_train
        else:
            X_train_balanced, y_train_balanced = X_train_scaled, y_train
        
        # Train Random Forest
        model = RandomForestClassifier(
            n_estimators=100,
            max_depth=10,
            min_samples_split=5,
            random_state=42,
            class_weight='balanced'
        )
        model.fit(X_train_balanced, y_train_balanced)
        
        # Evaluate
        y_pred = model.predict(X_test_scaled)
        y_proba = model.predict_proba(X_test_scaled)[:, 1]
        
        # Metrics
        accuracy = accuracy_score(y_test, y_pred)
        precision, recall, f1, _ = precision_recall_fscore_support(y_test, y_pred, average='binary')
        
        # Confusion matrix
        cm = confusion_matrix(y_test, y_pred)
        
        # Feature importance
        importance = pd.DataFrame({
            'feature': feature_names,
            'importance': model.feature_importances_
        }).sort_values('importance', ascending=False)
        
        # Cross-validation
        cv_scores = cross_val_score(model, X_train_scaled, y_train, cv=5, scoring='f1')
        
        # Store model
        self.models['detractor'] = model
        self.scalers['detractor'] = scaler
        
        return {
            'accuracy': round(accuracy, 3),
            'precision': round(precision, 3),
            'recall': round(recall, 3),
            'f1_score': round(f1, 3),
            'cv_mean': round(cv_scores.mean(), 3),
            'cv_std': round(cv_scores.std(), 3),
            'confusion_matrix': cm.tolist(),
            'feature_importance': importance.to_dict('records'),
            'train_size': len(X_train),
            'test_size': len(X_test),
            'detractor_rate': round(y.mean() * 100, 1)
        }
    
    def predict_detractor_risk(self, top_n=100):
        """Predict which customers are at risk of being detractors"""
        if 'detractor' not in self.models:
            self.train_detractor_model()
        
        if 'detractor' not in self.models:
            return pd.DataFrame()
        
        X, y, feature_names = self.prepare_classification_features()
        if X is None:
            return pd.DataFrame()
        
        model = self.models['detractor']
        scaler = self.scalers['detractor']
        
        X_scaled = scaler.transform(X)
        probabilities = model.predict_proba(X_scaled)[:, 1]
        
        # Add risk scores to dataframe
        result = self.df.copy()
        result['detractor_risk'] = probabilities
        result['risk_level'] = pd.cut(
            probabilities, 
            bins=[0, 0.3, 0.6, 1.0], 
            labels=['Low', 'Medium', 'High']
        )
        
        # Get high risk customers
        high_risk = result.nlargest(top_n, 'detractor_risk')
        
        # Select relevant columns
        cols_to_show = ['detractor_risk', 'risk_level']
        if self.COLS.get('BRANCH') in result.columns:
            cols_to_show.append(self.COLS['BRANCH'])
        if self.COLS.get('RATING') in result.columns:
            cols_to_show.append(self.COLS['RATING'])
        if self.COLS.get('NPS') in result.columns:
            cols_to_show.append(self.COLS['NPS'])
        
        return high_risk[cols_to_show].round(3)
    
    # ==========================================
    # 2. CUSTOMER CLUSTERING
    # ==========================================
    
    def prepare_clustering_features(self):
        """Prepare features for clustering"""
        if not ML_AVAILABLE:
            return None
        
        df = self.df.copy()
        features = []
        
        # Rating
        rating_col = self.COLS.get('RATING')
        if rating_col and rating_col in df.columns:
            features.append(df[rating_col].fillna(3).values.reshape(-1, 1))
        
        # NPS
        nps_col = self.COLS.get('NPS')
        if nps_col and nps_col in df.columns:
            features.append(df[nps_col].fillna(5).values.reshape(-1, 1))
        
        # Issue count
        weakness_col = self.COLS.get('WEAKNESS')
        if weakness_col and weakness_col in df.columns:
            issue_counts = df[weakness_col].fillna('').apply(
                lambda x: len(str(x).split('،')) if str(x).strip() else 0
            ).values.reshape(-1, 1)
            features.append(issue_counts)
        
        # Strength count
        strength_col = self.COLS.get('STRENGTH')
        if strength_col and strength_col in df.columns:
            strength_counts = df[strength_col].fillna('').apply(
                lambda x: len(str(x).split('،')) if str(x).strip() else 0
            ).values.reshape(-1, 1)
            features.append(strength_counts)
        
        if not features:
            return None
        
        X = np.hstack(features)
        return X
    
    def perform_clustering(self, n_clusters=5):
        """Perform K-Means clustering to find customer segments"""
        if not ML_AVAILABLE:
            return {'error': 'scikit-learn not installed'}
        
        X = self.prepare_clustering_features()
        if X is None:
            return {'error': 'Could not prepare features'}
        
        # Scale features
        scaler = StandardScaler()
        X_scaled = scaler.fit_transform(X)
        
        # Find optimal number of clusters (Elbow method)
        inertias = []
        k_range = range(2, min(10, len(X) // 100 + 2))
        for k in k_range:
            kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
            kmeans.fit(X_scaled)
            inertias.append(kmeans.inertia_)
        
        # Train final model
        kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
        clusters = kmeans.fit_predict(X_scaled)
        
        # Analyze clusters
        df = self.df.copy()
        df['cluster'] = clusters
        
        cluster_stats = []
        for i in range(n_clusters):
            cluster_data = df[df['cluster'] == i]
            
            stats = {
                'cluster': i,
                'size': len(cluster_data),
                'percentage': round(len(cluster_data) / len(df) * 100, 1)
            }
            
            # Rating stats
            rating_col = self.COLS.get('RATING')
            if rating_col and rating_col in cluster_data.columns:
                stats['avg_rating'] = round(cluster_data[rating_col].mean(), 2)
            
            # NPS stats
            nps_col = self.COLS.get('NPS')
            if nps_col and nps_col in cluster_data.columns:
                stats['avg_nps'] = round(cluster_data[nps_col].mean(), 2)
                stats['promoter_pct'] = round((cluster_data[nps_col] >= 9).mean() * 100, 1)
                stats['detractor_pct'] = round((cluster_data[nps_col] <= 6).mean() * 100, 1)
            
            cluster_stats.append(stats)
        
        # Assign cluster names based on characteristics
        cluster_df = pd.DataFrame(cluster_stats)
        cluster_df = cluster_df.sort_values('avg_rating', ascending=False)
        
        # Name clusters
        names = ['⭐ Champions', '😊 Satisfied', '😐 Neutral', '😟 At Risk', '🚨 Critical']
        cluster_df['cluster_name'] = names[:len(cluster_df)]
        
        # PCA for visualization
        pca = PCA(n_components=2)
        X_pca = pca.fit_transform(X_scaled)
        
        return {
            'cluster_stats': cluster_df.to_dict('records'),
            'elbow_data': {'k': list(k_range), 'inertia': inertias},
            'pca_data': {
                'x': X_pca[:, 0].tolist(),
                'y': X_pca[:, 1].tolist(),
                'cluster': clusters.tolist()
            },
            'n_clusters': n_clusters,
            'total_samples': len(df)
        }
    
    # ==========================================
    # 3. ASSOCIATION RULES
    # ==========================================
    
    def get_association_rules(self, min_support=0.01, min_confidence=0.3):
        """Find association rules between issues"""
        if not MLXTEND_AVAILABLE:
            return {'error': 'mlxtend not installed. Run: pip install mlxtend'}
        
        weakness_col = self.COLS.get('WEAKNESS')
        if not weakness_col or weakness_col not in self.df.columns:
            return {'error': 'Weakness column not found'}
        
        # Extract issues as transactions
        transactions = []
        for _, row in self.df.iterrows():
            issues = str(row.get(weakness_col, '')).replace('،', ',').split(',')
            issues = [i.strip() for i in issues if i.strip()]
            if issues:
                transactions.append(issues)
        
        if len(transactions) < 100:
            return {'error': 'Not enough data for association rules'}
        
        # Encode transactions
        te = TransactionEncoder()
        te_ary = te.fit_transform(transactions)
        df_encoded = pd.DataFrame(te_ary, columns=te.columns_)
        
        # Find frequent itemsets
        frequent_itemsets = apriori(df_encoded, min_support=min_support, use_colnames=True)
        
        if len(frequent_itemsets) == 0:
            return {'error': 'No frequent itemsets found. Try lowering min_support.'}
        
        # Generate rules
        rules = association_rules(frequent_itemsets, metric='confidence', min_threshold=min_confidence)
        
        if len(rules) == 0:
            return {'error': 'No rules found. Try lowering min_confidence.'}
        
        # Format rules
        rules_list = []
        for _, row in rules.head(20).iterrows():
            antecedents = ', '.join(list(row['antecedents']))
            consequents = ', '.join(list(row['consequents']))
            rules_list.append({
                'if': antecedents,
                'then': consequents,
                'support': round(row['support'], 3),
                'confidence': round(row['confidence'], 3),
                'lift': round(row['lift'], 2)
            })
        
        # Top itemsets
        top_itemsets = []
        for _, row in frequent_itemsets.nlargest(15, 'support').iterrows():
            items = ', '.join(list(row['itemsets']))
            top_itemsets.append({
                'items': items,
                'support': round(row['support'], 3),
                'count': int(row['support'] * len(transactions))
            })
        
        return {
            'rules': rules_list,
            'frequent_itemsets': top_itemsets,
            'total_transactions': len(transactions),
            'unique_items': len(te.columns_)
        }
    
    # ==========================================
    # 4. ANOMALY DETECTION
    # ==========================================
    
    def detect_anomalies(self, contamination=0.05):
        """Detect anomalous customer feedback patterns"""
        if not ML_AVAILABLE:
            return {'error': 'scikit-learn not installed'}
        
        X = self.prepare_clustering_features()
        if X is None:
            return {'error': 'Could not prepare features'}
        
        # Scale features
        scaler = StandardScaler()
        X_scaled = scaler.fit_transform(X)
        
        # Isolation Forest
        iso_forest = IsolationForest(
            contamination=contamination,
            random_state=42,
            n_estimators=100
        )
        anomaly_labels = iso_forest.fit_predict(X_scaled)
        anomaly_scores = iso_forest.decision_function(X_scaled)
        
        # -1 = anomaly, 1 = normal
        df = self.df.copy()
        df['is_anomaly'] = (anomaly_labels == -1).astype(int)
        df['anomaly_score'] = -anomaly_scores  # Higher = more anomalous
        
        # Get anomalies
        anomalies = df[df['is_anomaly'] == 1].copy()
        
        # Analyze anomalies
        normal = df[df['is_anomaly'] == 0]
        
        anomaly_stats = {
            'total_anomalies': len(anomalies),
            'anomaly_rate': round(len(anomalies) / len(df) * 100, 2),
        }
        
        # Compare anomalies vs normal
        rating_col = self.COLS.get('RATING')
        if rating_col and rating_col in df.columns:
            anomaly_stats['anomaly_avg_rating'] = round(anomalies[rating_col].mean(), 2)
            anomaly_stats['normal_avg_rating'] = round(normal[rating_col].mean(), 2)
        
        nps_col = self.COLS.get('NPS')
        if nps_col and nps_col in df.columns:
            anomaly_stats['anomaly_avg_nps'] = round(anomalies[nps_col].mean(), 2)
            anomaly_stats['normal_avg_nps'] = round(normal[nps_col].mean(), 2)
        
        # Anomaly types
        anomaly_types = []
        
        # Low rating + high NPS (suspicious)
        if rating_col in df.columns and nps_col in df.columns:
            suspicious = anomalies[(anomalies[rating_col] <= 2) & (anomalies[nps_col] >= 9)]
            if len(suspicious) > 0:
                anomaly_types.append({
                    'type': 'Low Rating + High NPS',
                    'description': 'Gave low rating but would recommend',
                    'count': len(suspicious),
                    'icon': '🤔'
                })
            
            # High rating + low NPS
            suspicious2 = anomalies[(anomalies[rating_col] >= 4) & (anomalies[nps_col] <= 3)]
            if len(suspicious2) > 0:
                anomaly_types.append({
                    'type': 'High Rating + Low NPS',
                    'description': 'Gave high rating but would not recommend',
                    'count': len(suspicious2),
                    'icon': '⚠️'
                })
        
        # Extreme scores
        extreme_low = anomalies[anomalies['anomaly_score'] > anomalies['anomaly_score'].quantile(0.9)]
        if len(extreme_low) > 0:
            anomaly_types.append({
                'type': 'Extreme Outliers',
                'description': 'Highly unusual patterns',
                'count': len(extreme_low),
                'icon': '🚨'
            })
        
        # Top anomalies to review
        top_anomalies = anomalies.nlargest(20, 'anomaly_score')
        
        cols_to_show = ['anomaly_score']
        for col in [rating_col, nps_col, self.COLS.get('BRANCH')]:
            if col and col in top_anomalies.columns:
                cols_to_show.append(col)
        
        return {
            'stats': anomaly_stats,
            'anomaly_types': anomaly_types,
            'top_anomalies': top_anomalies[cols_to_show].round(3).to_dict('records'),
            'score_distribution': {
                'normal_mean': round(normal['anomaly_score'].mean(), 3),
                'anomaly_mean': round(anomalies['anomaly_score'].mean(), 3)
            }
        }
    
    # ==========================================
    # 5. CHURN PREDICTION
    # ==========================================
    
    def prepare_churn_features(self):
        """
        Prepare features for churn prediction.
        Note: True churn prediction requires repeat customer data.
        This creates a proxy based on available signals.
        """
        if not ML_AVAILABLE:
            return None, None
        
        df = self.df.copy()
        
        # Create churn proxy: Low rating + Low NPS + Multiple issues
        rating_col = self.COLS.get('RATING')
        nps_col = self.COLS.get('NPS')
        weakness_col = self.COLS.get('WEAKNESS')
        
        if not all([rating_col, nps_col]):
            return None, None
        
        # Churn signals
        df['low_rating'] = (df[rating_col] <= 2).astype(int)
        df['low_nps'] = (df[nps_col] <= 6).astype(int)
        
        if weakness_col and weakness_col in df.columns:
            df['issue_count'] = df[weakness_col].fillna('').apply(
                lambda x: len(str(x).split('،')) if str(x).strip() else 0
            )
            df['has_issues'] = (df['issue_count'] > 0).astype(int)
        else:
            df['issue_count'] = 0
            df['has_issues'] = 0
        
        # Churn proxy score (0-3)
        df['churn_score'] = df['low_rating'] + df['low_nps'] + df['has_issues']
        
        # Binary churn target (score >= 2)
        df['likely_churn'] = (df['churn_score'] >= 2).astype(int)
        
        # Features
        features = pd.DataFrame({
            'rating': df[rating_col].fillna(3),
            'nps': df[nps_col].fillna(5),
            'issue_count': df['issue_count']
        })
        
        # Add branch if available
        branch_col = self.COLS.get('BRANCH')
        if branch_col and branch_col in df.columns:
            le = LabelEncoder()
            features['branch_encoded'] = le.fit_transform(df[branch_col].fillna('Unknown'))
        
        return features, df['likely_churn']
    
    def train_churn_model(self):
        """Train churn prediction model"""
        if not ML_AVAILABLE:
            return {'error': 'scikit-learn not installed'}
        
        X, y = self.prepare_churn_features()
        
        if X is None:
            return {'error': 'Could not prepare features'}
        
        # Split
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.2, random_state=42, stratify=y
        )
        
        # Scale
        scaler = StandardScaler()
        X_train_scaled = scaler.fit_transform(X_train)
        X_test_scaled = scaler.transform(X_test)
        
        # Train Gradient Boosting
        model = GradientBoostingClassifier(
            n_estimators=100,
            max_depth=5,
            random_state=42
        )
        model.fit(X_train_scaled, y_train)
        
        # Evaluate
        y_pred = model.predict(X_test_scaled)
        y_proba = model.predict_proba(X_test_scaled)[:, 1]
        
        accuracy = accuracy_score(y_test, y_pred)
        precision, recall, f1, _ = precision_recall_fscore_support(y_test, y_pred, average='binary')
        cm = confusion_matrix(y_test, y_pred)
        
        # Feature importance
        importance = pd.DataFrame({
            'feature': X.columns,
            'importance': model.feature_importances_
        }).sort_values('importance', ascending=False)
        
        # Store model
        self.models['churn'] = model
        self.scalers['churn'] = scaler
        
        return {
            'accuracy': round(accuracy, 3),
            'precision': round(precision, 3),
            'recall': round(recall, 3),
            'f1_score': round(f1, 3),
            'confusion_matrix': cm.tolist(),
            'feature_importance': importance.to_dict('records'),
            'churn_rate': round(y.mean() * 100, 1),
            'note': 'This is a proxy model based on rating/NPS/issues. True churn requires repeat customer data.'
        }
    
    def predict_churn_risk(self, top_n=100):
        """Predict churn risk for customers"""
        if 'churn' not in self.models:
            self.train_churn_model()
        
        if 'churn' not in self.models:
            return pd.DataFrame()
        
        X, y = self.prepare_churn_features()
        if X is None:
            return pd.DataFrame()
        
        model = self.models['churn']
        scaler = self.scalers['churn']
        
        X_scaled = scaler.transform(X)
        probabilities = model.predict_proba(X_scaled)[:, 1]
        
        result = self.df.copy()
        result['churn_risk'] = probabilities
        result['churn_level'] = pd.cut(
            probabilities,
            bins=[0, 0.3, 0.6, 1.0],
            labels=['Low', 'Medium', 'High']
        )
        
        high_risk = result.nlargest(top_n, 'churn_risk')
        
        cols_to_show = ['churn_risk', 'churn_level']
        for col_name in ['BRANCH', 'RATING', 'NPS']:
            col = self.COLS.get(col_name)
            if col and col in result.columns:
                cols_to_show.append(col)
        
        return high_risk[cols_to_show].round(3)
    
    # ==========================================
    # 6. ML SUMMARY
    # ==========================================
    
    def get_ml_summary(self):
        """Get summary of all ML analyses"""
        summary = {
            'ml_available': ML_AVAILABLE,
            'mlxtend_available': MLXTEND_AVAILABLE,
            'imblearn_available': IMBLEARN_AVAILABLE,
            'models_trained': list(self.models.keys()),
            'data_size': len(self.df)
        }
        
        # Quick stats
        nps_col = self.COLS.get('NPS')
        if nps_col and nps_col in self.df.columns:
            summary['detractor_rate'] = round((self.df[nps_col] <= 6).mean() * 100, 1)
            summary['promoter_rate'] = round((self.df[nps_col] >= 9).mean() * 100, 1)
        
        return summary