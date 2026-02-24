# -*- coding: utf-8 -*-
"""
QFD Analyzer Module for Shila Dashboard
"""

import pandas as pd
import numpy as np
from collections import Counter
from itertools import combinations
import re
import datetime
import streamlit as st

from config import COLS, STOPWORDS, ASPECTS, EXCLUDE_PRODUCTS

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_PERSIAN_SUPPORT = True
except ImportError:
    HAS_PERSIAN_SUPPORT = False

try:
    from hazm import Normalizer, word_tokenize, stopwords_list
    HAZM_AVAILABLE = True
except ImportError:
    HAZM_AVAILABLE = False

class ShilaAnalyzer:
    def __init__(self, df, cols):
        self.df = df.copy()
        self.cols = cols
        self._preprocess_data()
    
    def _preprocess_data(self):
        # Auto-normalize branch names FIRST
        self._normalize_branch_names()
        
        nps_col = self.cols.get('NPS', 'NPS') # Use self.cols now that it's initialized
        if COLS['NPS'] in self.df and self.df[COLS['NPS']].notna().any():
            self.df['NPS_Segment'] = pd.cut(
                self.df[COLS['NPS']],
                bins=[-1, 6, 8, 10],
                labels=['Detractor', 'Passive', 'Promoter']
            )   
        else:
            # Create an empty column so other functions don't crash looking for the header
            self.df['NPS_Segment'] = None
        
        # 3. Handle Date and Time Processing
        date_col = self.cols.get('DATE', 'Date')
        created_col = self.cols.get('CREATED_AT', 'Order Created At')
        
        if date_col in self.df.columns:
            # Check if we are dealing with SnappFood (standard datetime) or Shila (Persian strings)
            # We can check if the first non-null value is already a datetime object
            first_val = self.df[date_col].dropna().iloc[0] if not self.df[date_col].dropna().empty else None
        
            if isinstance(first_val, (pd.Timestamp, datetime.datetime)):
                # SNAPPFOOD LOGIC: Preserving Time
                self.df['parsed_date'] = self.df[date_col] # Keep original for time extraction
            
                # Create the strings for daily/monthly grouping
                self.df['date_str'] = self.df[date_col].dt.strftime('%Y/%m/%d')
                self.df['year_month'] = self.df[date_col].dt.strftime('%Y/%m')
            
                # Ensure the Hourly Chart has access to the full datetime
                if created_col in self.df.columns:
                    self.df[created_col] = pd.to_datetime(self.df[created_col], errors='coerce')
            else:
                # ORIGINAL LOGIC: Persian Date Parsing
                self.df['parsed_date'] = self.df[date_col].apply(self._parse_persian_date)
                valid = self.df['parsed_date'].notna()
                if valid.any():
                    self.df.loc[valid, 'date_str'] = self.df.loc[valid, 'parsed_date'].apply(
                        lambda x: f"{x[0]}/{x[1]:02d}/{x[2]:02d}" if x else None)
                    self.df.loc[valid, 'year_month'] = self.df.loc[valid, 'parsed_date'].apply(
                        lambda x: f"{x[0]}/{x[1]:02d}" if x else None)
    
    def _normalize_branch_names(self):
        """Auto-detect and normalize branch name variations"""
        if COLS['BRANCH'] not in self.df.columns:
            return
    
        # Get all unique branch names
        all_branches = self.df[COLS['BRANCH']].dropna().unique()
    
        # Group by spaceless version
        groups = {}
        for branch in all_branches:
            # Clean and create spaceless key
            clean = str(branch).strip()
            clean = clean.replace('ي', 'ی').replace('ك', 'ک').replace('ە', 'ه')
            spaceless = clean.replace(' ', '').replace('\u200c', '')
        
            if spaceless not in groups:
                groups[spaceless] = []
            groups[spaceless].append(branch)
    
        # Create mapping: all variations → best version
        mapping = {}
        for spaceless, variations in groups.items():
            # Pick the best version:
            # - Prefer one WITH space (more readable)
            # - If tie, pick most frequent
        
            best = None
            for v in variations:
                if ' ' in str(v) or '\u200c' in str(v):  # Has space or half-space
                    best = v
                    break
        
            # If none has space, pick most frequent
            if best is None:
                counts = self.df[COLS['BRANCH']].value_counts()
                best = max(variations, key=lambda x: counts.get(x, 0))
        
            # Map all variations to best
            for v in variations:
                mapping[v] = best
    
        # Apply mapping
        self.df[COLS['BRANCH']] = self.df[COLS['BRANCH']].map(mapping).fillna(self.df[COLS['BRANCH']])
    
    def _parse_persian_date(self, date_str):
        if pd.isna(date_str): return None
        try:
            match = re.match(r'(\d{4})[/\-](\d{1,2})[/\-](\d{1,2})', str(date_str).strip())
            if match: return (int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except: pass
        return None
    
    def _extract_tags(self, series):
        tags = []
        for text in series.dropna():
            text = str(text).replace(',', '،')
            tags.extend([t.strip() for t in text.split('،') if t.strip()])
        return Counter(tags)
    
    def get_kpis(self):
        total = len(self.df)
        nps_col, rating_col = COLS['NPS'], COLS['RATING']
    
        # Check if NPS column exists AND has at least one non-null value
        if nps_col in self.df.columns and self.df[nps_col].notna().any():
            promoters = (self.df[nps_col] >= 9).sum()
            passives = ((self.df[nps_col] >= 7) & (self.df[nps_col] <= 8)).sum()
            detractors = (self.df[nps_col] <= 6).sum()
        
            # Calculate NPS based on rows that actually have NPS data
            nps_count = promoters + passives + detractors
            nps_score = ((promoters - detractors) / nps_count) * 100 if nps_count > 0 else 0
        else:
            # Fallback for SnappFood files
            promoters = passives = detractors = nps_score = 0
    
        avg_rating = self.df[rating_col].mean() if rating_col in self.df.columns else 0
        response_rate = (self.df[COLS['COMMENT']].notna().sum() / total * 100) if COLS['COMMENT'] in self.df.columns and total > 0 else 0
    
        return {
            'total_orders': total, 
            'nps_score': round(nps_score, 1), 
            'avg_rating': round(avg_rating, 2),
            'promoters': promoters, 
            'passives': passives, 
            'detractors': detractors,
            'response_rate': round(response_rate, 1)
        }
    
    def get_rating_distribution(self):
        if COLS['RATING'] not in self.df.columns: return pd.DataFrame()
        dist = self.df[COLS['RATING']].value_counts().sort_index().reset_index()
        dist.columns = ['Rating', 'Count']
        dist['Percentage'] = (dist['Count'] / dist['Count'].sum() * 100).round(1)
        return dist
    
    def get_nps_distribution(self):
        if COLS['NPS'] not in self.df.columns: return pd.DataFrame()
        dist = self.df[COLS['NPS']].value_counts().sort_index().reset_index()
        dist.columns = ['NPS', 'Count']
        dist['Segment'] = dist['NPS'].apply(lambda x: 'Promoter' if x >= 9 else ('Passive' if x >= 7 else 'Detractor'))
        return dist
    
    def get_pareto_analysis(self):
        if COLS['WEAKNESS'] not in self.df.columns: return pd.DataFrame()
        records = []
        for _, row in self.df.iterrows():
            if pd.notna(row[COLS['WEAKNESS']]):
                tags = [t.strip() for t in str(row[COLS['WEAKNESS']]).replace(',', '،').split('،') if t.strip()]
                for tag in tags:
                    records.append({'tag': tag, 'damage': 5 - row[COLS['RATING']], 'rating': row[COLS['RATING']]})
        if not records: return pd.DataFrame()
        df_tags = pd.DataFrame(records)
        pareto = df_tags.groupby('tag').agg(total_damage=('damage', 'sum'), frequency=('damage', 'count'), avg_rating=('rating', 'mean')).reset_index()
        pareto = pareto.sort_values('total_damage', ascending=False)
        pareto['cumulative_damage'] = pareto['total_damage'].cumsum()
        pareto['cumulative_pct'] = (pareto['cumulative_damage'] / pareto['total_damage'].sum() * 100).round(1)
        pareto['avg_rating'] = pareto['avg_rating'].round(2)
        return pareto
    
    def get_kano_analysis(self):
        if COLS['RATING'] not in self.df.columns: return pd.DataFrame()
        all_attrs = set()
        for col in [COLS['STRENGTH'], COLS['WEAKNESS']]:
            if col in self.df.columns:
                for text in self.df[col].dropna():
                    all_attrs.update([t.strip() for t in str(text).replace(',', '،').split('،') if t.strip()])
        baseline = self.df[COLS['RATING']].mean()
        kano_data = []
        for attr in all_attrs:
            s_mask = self.df[COLS['STRENGTH']].fillna('').str.contains(attr, regex=False) if COLS['STRENGTH'] in self.df.columns else pd.Series([False]*len(self.df))
            w_mask = self.df[COLS['WEAKNESS']].fillna('').str.contains(attr, regex=False) if COLS['WEAKNESS'] in self.df.columns else pd.Series([False]*len(self.df))
            s_cnt, w_cnt = s_mask.sum(), w_mask.sum()
            s_rat = self.df.loc[s_mask, COLS['RATING']].mean() if s_cnt > 0 else np.nan
            w_rat = self.df.loc[w_mask, COLS['RATING']].mean() if w_cnt > 0 else np.nan
            if pd.notna(s_rat) and pd.notna(w_rat) and (s_cnt + w_cnt) >= 10:
                lift, drop = s_rat - baseline, baseline - w_rat
                ktype = 'Must-Be' if drop > 0.8 and lift < 0.3 else ('Delighter' if lift > 0.5 and drop < 0.3 else 'Performance')
                kano_data.append({'attribute': attr, 'kano_type': ktype, 'lift_as_strength': round(lift, 3), 'drop_as_weakness': round(drop, 3), 'strength_mentions': s_cnt, 'weakness_mentions': w_cnt})
        return pd.DataFrame(kano_data)
    
    def get_branch_product_performance(self, min_orders=1):
        """Explodes SnappFood product strings to analyze individual item performance."""
        df = self.df.copy()
    
        # Use the column names from your config
        prod_col = self.cols.get('Items')
        branch_col = self.cols.get('BRANCH')
        rating_col = self.cols.get('RATING')

        if not all(col in df.columns for col in [prod_col, branch_col, rating_col]):
            return pd.DataFrame()

        # Split the SnappFood string (e.g., "Item 1، Item 2")
        df[prod_col] = df[prod_col].astype(str).str.split(r'[،,]')
        df = df.explode(prod_col).reset_index(drop=True)
        df[prod_col] = df[prod_col].str.strip()
        
        # Filter out empty strings
        df = df[df[prod_col] != '']
        
        return df.pivot_table(index=branch_col, columns=prod_col, values=rating_col, aggfunc='mean')
    
    def get_product_analysis(self):
        """Analyze performance by product"""
        product_col = COLS['PRODUCT']
        rating_col = COLS['RATING']
    
        if not product_col or product_col not in self.df.columns:
            return pd.DataFrame()
    
        # Extract individual products (comma-separated)
        records = []
        for _, row in self.df.iterrows():
            if pd.notna(row[product_col]):
                products = str(row[product_col]).replace(',', '،').split('،')  # ✅ CORRECT!
                for product in products:
                    product = product.strip()
                    if product and product not in EXCLUDE_PRODUCTS:
                        records.append({
                            'product': product,
                            'rating': row[rating_col],
                            'nps': row.get(COLS['NPS'], None)
                        })
    
        if not records:
            return pd.DataFrame()
    
        df_products = pd.DataFrame(records)
    
        # Aggregate
        product_stats = df_products.groupby('product').agg(
            avg_rating=('rating', 'mean'),
            order_count=('rating', 'count'),
            rating_std=('rating', 'std')
        ).reset_index()
    
        product_stats = product_stats[product_stats['order_count'] >= 5]  # Min orders
        product_stats = product_stats.sort_values('avg_rating', ascending=False)
        product_stats['rank'] = range(1, len(product_stats) + 1)
    
        return product_stats.round(2)
    
    def get_branch_analysis(self, min_orders=10):
        if COLS['BRANCH'] not in self.df.columns: return pd.DataFrame(), pd.DataFrame()
        stats = self.df.groupby(COLS['BRANCH']).agg({COLS['RATING']: ['mean', 'std', 'count']}).reset_index()
        stats.columns = ['branch', 'avg_rating', 'rating_std', 'order_count']
        if COLS['NPS'] in self.df.columns:
            nps = self.df.groupby(COLS['BRANCH'])[COLS['NPS']].apply(lambda g: ((g >= 9).sum() - (g <= 6).sum()) / len(g) * 100 if len(g) > 0 else 0).reset_index()
            nps.columns = ['branch', 'nps_score']
            stats = stats.merge(nps, on='branch')
        stats = stats[stats['order_count'] >= min_orders]
        overall = self.df[COLS['RATING']].mean()
        stats['rating_vs_avg'] = stats['avg_rating'] - overall
        stats = stats.sort_values('avg_rating', ascending=False)
        stats['rank'] = range(1, len(stats) + 1)
        issues = []
        if COLS['WEAKNESS'] in self.df.columns:
            for br in stats.tail(5)['branch']:
                br_tags = self._extract_tags(self.df[self.df[COLS['BRANCH']] == br][COLS['WEAKNESS']])
                for tag, cnt in br_tags.most_common(5):
                    issues.append({'branch': br, 'issue': tag, 'count': cnt})
        return stats, pd.DataFrame(issues)
    
    def get_branch_product_matrix(self):
        """Which products perform best at which branches (supports both Shila and SnappFood)"""
        from config import COLS, EXCLUDE_PRODUCTS, EXCLUDE_BRANCHES
        import re
        if self.df.empty:
            return pd.DataFrame()
        
        df = self.df.copy()
    
        # 1. Determine which column to use (Check ORDER_ITEMS first for SnappFood)
        sf_col = COLS.get('ORDER_ITEMS')
        original_col = COLS.get('PRODUCT')
        branch_col = COLS.get('BRANCH')
        rating_col = COLS.get('RATING')

        # Find the active column: Check if sf_col exists and isn't entirely empty
        if sf_col in self.df.columns and self.df[sf_col].dropna().astype(str).str.len().sum() > 0:
            active_product_col = sf_col
        elif original_col in self.df.columns:
            active_product_col = original_col
        else:
            return pd.DataFrame()

        if branch_col not in self.df.columns:
            return pd.DataFrame()

        records = []
        # 2. Iterate and process strings (Handle both standard and Persian commas)
        for _, row in self.df.iterrows():
            raw_val = row.get(active_product_col)
            branch_val = row.get(branch_col)
            rating_val = row.get(rating_col)
            
            if any(keyword in branch_val for keyword in EXCLUDE_BRANCHES):
                continue
        
            if pd.notna(raw_val):
                # Split items using regex to catch both standard (,) and Persian (،) commas
                products = re.split(r'[,،]', str(raw_val))
            
                for product in products:
                    product = product.strip()
                    # 3. Filter out excluded products (Side dishes, drinks, etc.)
                    if product and product not in EXCLUDE_PRODUCTS:
                        records.append({
                            'branch': branch_val,
                            'product': product,
                            'rating': rating_val
                        })

        if not records:
            return pd.DataFrame()
    
        # 4. Create the heatmap matrix
        df_bp = pd.DataFrame(records)
        matrix = df_bp.pivot_table(
            values='rating',
            index='branch',
            columns='product',
            aggfunc='mean'
        ).round(2)

        return matrix
    
    def get_low_rating_deep_dive(self):
        """Analyzes 1-3 star reviews to find recurring themes across branches."""
        # 1. Filter for low ratings
        low_df = self.df[self.df[self.cols.get('RATING')] <= 3].copy()
        if low_df.empty:
            return pd.DataFrame(), pd.DataFrame()

        from config import ASPECTS
        branch_col = self.cols.get('BRANCH')
        comment_col = self.cols.get('COMMENT')
    
        # 2. Map comments to Aspects (Topics)
        def identify_topics(text):
            if not text or pd.isna(text): return "Other"
            found = []
            for aspect, keywords in ASPECTS.items():
                if any(kw in str(text) for kw in keywords):
                    found.append(aspect)
            return found if found else ["Uncategorized"]

        low_df['topics'] = low_df[comment_col].apply(identify_topics)
        exploded_low = low_df.explode('topics')

        # 3. Aggregate by Branch and Topic
        topic_summary = exploded_low.groupby([branch_col, 'topics']).size().reset_index(name='count')
    
        # 4. Weekly Trend of Low Ratings
        date_col = self.cols.get('CREATED_AT')
        weekly_trend = low_df.groupby(low_df[date_col].dt.date).size().reset_index(name='complaint_count')
    
        return topic_summary, weekly_trend
    
    def get_aspect_sentiment(self):
        if COLS['COMMENT'] not in self.df.columns: return pd.DataFrame()
        df_v = self.df[[COLS['COMMENT'], COLS['RATING']]].dropna()
        if len(df_v) < 20: return pd.DataFrame()
        results = []
        for aspect, keywords in ASPECTS.items():
            mask = df_v[COLS['COMMENT']].str.contains('|'.join(keywords), regex=True, na=False)
            asp_df = df_v[mask]
            n = len(asp_df)
            if n >= 5:
                avg = asp_df[COLS['RATING']].mean()
                pos, neg = (asp_df[COLS['RATING']] >= 4).sum(), (asp_df[COLS['RATING']] <= 2).sum()
                results.append({'aspect': aspect, 'mentions': n, 'avg_rating': round(avg, 2), 'positive_pct': round(pos/n*100, 1), 'negative_pct': round(neg/n*100, 1), 'sentiment_score': round((pos-neg)/n, 3)})
        return pd.DataFrame(results).sort_values('mentions', ascending=False)
    
    def get_hourly_trends(self):
        # Ensure the date column is datetime objects
        df = self.df.copy()
    
        # Force conversion of SnappFood format: 26/12/2025 18:37:39
        date_col = self.cols.get('CREATED_AT', 'Order Created At')
        rating_col = self.cols.get('RATING') # This will be the Persian string from config
        
        # 1. Validation: Check if columns exist
        if rating_col not in df.columns:
            st.error(f"Rating column not found. Expected: {rating_col}")
            return pd.DataFrame()
    
        if date_col not in df.columns or rating_col not in df.columns:
            return pd.DataFrame()
    
        # 1. Convert to datetime if not already
        if not pd.api.types.is_datetime64_any_dtype(df[date_col]):
            df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    
        # 2. Extract hour
        df['hour'] = df[date_col].dt.hour
    
        # 3. Group by hour - USE THE MAPPED COLUMN NAME HERE
        hourly_stats = df.groupby('hour').agg(
            avg_rating=(rating_col, 'mean'), # FIXED: Use rating_col variable
            order_count=(rating_col, 'count') # FIXED: Use rating_col variable
        ).reset_index()
    
        # 4. Fill in missing hours (0-23)
        all_hours = pd.DataFrame({'hour': range(24)})
        return all_hours.merge(hourly_stats, on='hour', how='left').fillna(0)
    
    def get_peak_hour_analysis(self):
        """Calculate busiest and best/worst performing hours."""
        df = self.df.copy()
    
        date_col = self.cols.get('CREATED_AT', 'Order Created At')
        rating_col = self.cols.get('RATING')
    
        if date_col not in df.columns or rating_col not in df.columns:
            return None

        # Ensure datetime and extract hour
        if not pd.api.types.is_datetime64_any_dtype(df[date_col]):
            df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    
        df['hour'] = df[date_col].dt.hour
    
        # Group stats
        stats = df.groupby('hour').agg(
            avg_rating=(rating_col, 'mean'),
            order_count=(rating_col, 'count')
        )
    
        if stats.empty:
            return None
        
        return {
            'busiest_hour': int(stats['order_count'].idxmax()), # Cast to int here
            'peak_volume': stats['order_count'].max(),
            'best_hour': int(stats['avg_rating'].idxmax()),    # Cast to int here
            'best_rating': stats['avg_rating'].max(),
            'worst_hour': int(stats['avg_rating'].idxmin()),   # Cast to int here
            'worst_rating': stats['avg_rating'].min()
        }   

    def get_daily_trends(self):
        if 'date_str' not in self.df.columns: return pd.DataFrame()
        df = self.df[self.df['date_str'].notna()]
        if len(df) < 7: return pd.DataFrame()
        daily = df.groupby('date_str').agg({COLS['RATING']: ['mean', 'count']}).reset_index()
        daily.columns = ['date', 'avg_rating', 'order_count']
        daily = daily.sort_values('date')
        daily['rating_7day_avg'] = daily['avg_rating'].rolling(window=7, min_periods=1).mean()
        daily['orders_7day_avg'] = daily['order_count'].rolling(window=7, min_periods=1).mean()
        if COLS['NPS'] in self.df.columns:
            nps = df.groupby('date_str')[COLS['NPS']].apply(lambda g: ((g >= 9).sum() - (g <= 6).sum()) / len(g) * 100 if len(g) > 0 else 0).reset_index()
            nps.columns = ['date', 'nps_score']
            daily = daily.merge(nps, on='date', how='left')
        return daily
    
    def get_day_of_week_analysis(self):
        """Analyze patterns by day of week"""
        if 'parsed_date' not in self.df.columns:
            return pd.DataFrame()
    
        # Persian day names
        PERSIAN_DAYS = {
        0: 'شنبه', 1: 'یکشنبه', 2: 'دوشنبه', 3: 'سه‌شنبه',
        4: 'چهارشنبه', 5: 'پنجشنبه', 6: 'جمعه'
    }
    
        df = self.df[self.df['parsed_date'].notna()].copy()
    
        # Convert Persian date to day of week (simplified - assumes recent dates)
        # For accurate conversion, use jdatetime library
        try:
            import jdatetime
            df['day_of_week'] = df['parsed_date'].apply(
                lambda x: jdatetime.date(x[0], x[1], x[2]).weekday() if x else None
            )
        except:
            return pd.DataFrame()
    
        day_stats = df.groupby('day_of_week').agg({
            COLS['RATING']: ['mean', 'count'],
            COLS['NPS']: 'mean'
        }).reset_index()
        day_stats.columns = ['day_num', 'avg_rating', 'order_count', 'avg_nps']
        day_stats['day_name'] = day_stats['day_num'].map(PERSIAN_DAYS)
    
        return day_stats.round(2)
    
    def get_period_analysis(self):
        """Analyze patterns by period of month (early/mid/late)"""
        if 'parsed_date' not in self.df.columns:
            return pd.DataFrame()
        
        df = self.df[self.df['parsed_date'].notna()].copy()
        
        if len(df) < 30:
            return pd.DataFrame()
        
        df['day_of_month'] = df['parsed_date'].apply(lambda x: x[2] if x else None)
        
        def get_period(day):
            if day <= 10: return ('Early', 'اول ماه (۱-۱۰)')
            elif day <= 20: return ('Mid', 'میانه ماه (۱۱-۲۰)')
            else: return ('Late', 'آخر ماه (۲۱-۳۱)')
        
        df['period'], df['period_fa'] = zip(*df['day_of_month'].apply(get_period))
        
        period_stats = df.groupby(['period', 'period_fa']).agg({
            COLS['RATING']: ['mean', 'count']
        }).reset_index()
        period_stats.columns = ['period', 'period_fa', 'avg_rating', 'order_count']
        
        if COLS['NPS'] in df.columns:
            nps_stats = df.groupby('period')[COLS['NPS']].mean().reset_index()
            nps_stats.columns = ['period', 'avg_nps']
            period_stats = period_stats.merge(nps_stats, on='period')
        
        order = {'Early': 0, 'Mid': 1, 'Late': 2}
        period_stats['sort_order'] = period_stats['period'].map(order)
        period_stats = period_stats.sort_values('sort_order').drop('sort_order', axis=1)
        
        return period_stats.round(2)
    
    def get_mom_comparison(self):
        """Month-over-month performance comparison"""
        if 'year_month' not in self.df.columns:
            return pd.DataFrame()
        
        df = self.df[self.df['year_month'].notna()].copy()
        
        if len(df) < 30:
            return pd.DataFrame()

        monthly_ym = df.groupby('year_month').size().reset_index(name='order_count')
        monthly_ym['avg_rating'] = df.groupby('year_month')[COLS['RATING']].mean().values
        
        if COLS['NPS'] in df.columns:
            monthly_nps = df.groupby('year_month')[COLS['NPS']].apply(
                lambda g: ((g >= 9).sum() - (g <= 6).sum()) / len(g) * 100 if len(g) > 0 else 0
            ).reset_index()
            monthly_nps.columns = ['year_month', 'nps_score']
            monthly_ym = monthly_ym.merge(monthly_nps, on='year_month')        

        monthly_ym = monthly_ym.sort_values('year_month')
        monthly_ym['rating_change'] = monthly_ym['avg_rating'].diff()
        monthly_ym['orders_change_pct'] = monthly_ym['order_count'].pct_change() * 100
        
        if 'nps_score' in monthly_ym.columns:
            monthly_ym['nps_change'] = monthly_ym['nps_score'].diff()
        
        return monthly_ym.round(2)
   
    def get_rating_nps_correlation(self):
        """Analyze relationship between rating and NPS"""
        rating_col = COLS['RATING']
        nps_col = COLS['NPS']
    
        if rating_col not in self.df.columns or nps_col not in self.df.columns:
            return {}
    
        df_valid = self.df[[rating_col, nps_col]].dropna()
        
        # Correlation
        correlation = df_valid[rating_col].corr(df_valid[nps_col])
    
        # Cross-tab: Rating vs NPS Segment
        df_valid['nps_segment'] = pd.cut(df_valid[nps_col], bins=[-1, 6, 8, 10], 
                                      labels=['Detractor', 'Passive', 'Promoter'])
    
        crosstab = pd.crosstab(df_valid[rating_col], df_valid['nps_segment'], normalize='index') * 100
    
        # Insights
        low_rating_promoters = df_valid[(df_valid[rating_col] <= 3) & (df_valid[nps_col] >= 9)]
        high_rating_detractors = df_valid[(df_valid[rating_col] >= 4) & (df_valid[nps_col] <= 6)]
        
        return {
            'correlation': round(correlation, 3),
            'crosstab': crosstab.round(1),
            'anomaly_low_rating_promoters': len(low_rating_promoters),
            'anomaly_high_rating_detractors': len(high_rating_detractors)
            }
    
    def get_low_rating_comments_by_hour(self, min_rating=1, max_rating=3):
        """Filter and return low-rating comments with their hours."""
        df = self.df.copy()
    
        date_col = self.cols.get('CREATED_AT', 'Order Created At')
        rating_col = self.cols.get('RATING')
        comment_col = self.cols.get('COMMENT')
    
        # Validation
        if not all(col in df.columns for col in [date_col, rating_col]):
            return pd.DataFrame()

        # Ensure datetime
        if not pd.api.types.is_datetime64_any_dtype(df[date_col]):
            df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
    
        df['hour'] = df[date_col].dt.hour
    
        # Filter for low ratings (1, 2, 3) and existing comments
        low_ratings = df[
            (df[rating_col] >= min_rating) & 
            (df[rating_col] <= max_rating) &
            (df[comment_col].notna())
        ].copy()
    
        return low_ratings[[date_col, 'hour', rating_col, comment_col, self.cols.get('BRANCH', 'Branch')]]
    
    def get_issue_category_analysis(self):
        """Detailed analysis of delivery, packaging, personnel issues"""
        categories = {
            'delivery': COLS['DELIVERY'],
            'packaging': COLS['PACKAGING'],
            'personnel': COLS['PERSONNEL']
        }
    
        results = []
        for category, col in categories.items():
            if col not in self.df.columns:
                continue
        
            # Count issues
            has_issue = self.df[col].notna() & (self.df[col] != '')
            issue_count = has_issue.sum()
            
            if issue_count == 0:
                continue
            
            issue_pct = (issue_count / len(self.df)) * 100
        
            # Average rating when this issue exists
            avg_rating_with_issue = self.df.loc[has_issue, COLS['RATING']].mean()
            avg_rating_without = self.df.loc[~has_issue, COLS['RATING']].mean()
            rating_impact = avg_rating_without - avg_rating_with_issue
            
            issue_tags = self._extract_tags(self.df.loc[has_issue, col])
            top_issues = [tag for tag, _ in issue_tags.most_common(3)]
            
            results.append({
                'category': category,
                'category_fa': {'delivery': 'پیک', 'packaging': 'بسته‌بندی', 'personnel': 'پرسنل'}[category],
                'issue_count': issue_count,
                'issue_pct': round(issue_pct, 1),
                'avg_rating_with_issue': round(avg_rating_with_issue, 2),
                'avg_rating_without': round(avg_rating_without, 2),
                'rating_impact': round(rating_impact, 2),
                'top_issues': '، '.join(top_issues) if top_issues else ''
            })
        
        if not results:
        # Return empty DataFrame with correct columns so downstream code (like AI Summary) doesn't crash
            return pd.DataFrame(columns=[
                'category', 'category_fa', 'issue_count', 'issue_pct', 
                'avg_rating_with_issue', 'avg_rating_without', 'rating_impact', 'top_issues'
            ])
        return pd.DataFrame(results).sort_values('rating_impact', ascending=False)
    
    def get_comment_keywords(self, top_n=20):
        """Extract most frequent keywords from comments"""
        comment_col = COLS['COMMENT']
        if comment_col not in self.df.columns:
            return pd.DataFrame(), pd.DataFrame()
    
        # Positive comments (rating >= 4)
        positive_comments = self.df[self.df[COLS['RATING']] >= 4][comment_col].dropna()
        # Negative comments (rating <= 2)
        negative_comments = self.df[self.df[COLS['RATING']] <= 2][comment_col].dropna()
    
        def extract_words(comments):
            words = []
            for comment in comments:
                tokens = str(comment).split()
                for token in tokens:
                    token = token.strip('.,!?،؟؛:»«')
                    if len(token) > 2 and token not in STOPWORDS:
                        words.append(token)
            return Counter(words)
    
        positive_words = extract_words(positive_comments)
        negative_words = extract_words(negative_comments)
    
        df_positive = pd.DataFrame(positive_words.most_common(top_n), columns=['word', 'count'])
        df_negative = pd.DataFrame(negative_words.most_common(top_n), columns=['word', 'count'])
    
        return df_positive, df_negative
    
    def get_top_issues(self, n=10):
        if COLS['WEAKNESS'] not in self.df.columns: return pd.DataFrame()
        return pd.DataFrame(self._extract_tags(self.df[COLS['WEAKNESS']]).most_common(n), columns=['Issue', 'Count'])
    
    def get_top_strengths(self, n=10):
        if COLS['STRENGTH'] not in self.df.columns: return pd.DataFrame()
        return pd.DataFrame(self._extract_tags(self.df[COLS['STRENGTH']]).most_common(n), columns=['Strength', 'Count'])
    
    def get_cooccurrence(self, n=15):
        if COLS['WEAKNESS'] not in self.df.columns: return pd.DataFrame()
        cooccur = Counter()
        for text in self.df[COLS['WEAKNESS']].dropna():
            tags = [t.strip() for t in str(text).replace(',', '،').split('،') if t.strip()]
            if len(tags) >= 2:
                for pair in combinations(sorted(tags), 2): cooccur[pair] += 1
        if not cooccur: return pd.DataFrame()
        return pd.DataFrame([{'issue_1': k[0], 'issue_2': k[1], 'count': v} for k, v in cooccur.most_common(n)])
    
    def get_recovery_opportunities(self):
        """Find customers who gave low ratings but high NPS (salvageable)"""
        df = self.df.copy()
        
        if COLS['NPS'] not in df.columns or COLS['RATING'] not in df.columns:
            return pd.DataFrame()
        
        df['segment'] = 'Neutral'
        df['segment_fa'] = 'خنثی'
        df['emoji'] = '😐'
    
        mask_happy = (df[COLS['RATING']] >= 4) & (df[COLS['NPS']] >= 9)
        df.loc[mask_happy, 'segment'] = 'Happy'
        df.loc[mask_happy, 'segment_fa'] = 'راضی'
        df.loc[mask_happy, 'emoji'] = '😊'
        
        mask_risk = (df[COLS['RATING']] <= 2) & (df[COLS['NPS']] <= 6)
        df.loc[mask_risk, 'segment'] = 'At Risk'
        df.loc[mask_risk, 'segment_fa'] = 'در خطر'
        df.loc[mask_risk, 'emoji'] = '🚨'
        
        mask_recovery = (df[COLS['RATING']] <= 3) & (df[COLS['NPS']] >= 7) & ~mask_happy
        df.loc[mask_recovery, 'segment'] = 'Recovery'
        df.loc[mask_recovery, 'segment_fa'] = 'قابل بازیابی'
        df.loc[mask_recovery, 'emoji'] = '🔄'
        
        mask_churner = (df[COLS['RATING']] >= 3) & (df[COLS['NPS']] <= 6) & ~mask_risk
        df.loc[mask_churner, 'segment'] = 'Silent Churner'
        df.loc[mask_churner, 'segment_fa'] = 'ریزش خاموش'
        df.loc[mask_churner, 'emoji'] = '⚠️'
        
        segment_counts = df.groupby(['segment', 'segment_fa', 'emoji']).size().reset_index(name='count')
        segment_counts['percentage'] = (segment_counts['count'] / len(df) * 100).round(1)
        
        order = {'At Risk': 0, 'Silent Churner': 1, 'Recovery': 2, 'Neutral': 3, 'Happy': 4}
        segment_counts['sort_order'] = segment_counts['segment'].map(order)
        segment_counts = segment_counts.sort_values('sort_order').drop('sort_order', axis=1)
        
        return segment_counts
    
    def get_unmapped_comments(self, category_type="Other"):
        """
        Returns rows where the topic was identified as 'Other' or 'Uncategorized'.
        category_type: "Other" or "Uncategorized"
        """
        df = self.df.copy()
        comment_col = self.cols.get('COMMENT')
        rating_col = self.cols.get('RATING')
        branch_col = self.cols.get('BRANCH')
    
        # We apply the same identification logic used in your deep dive
        from config import ASPECTS
        def is_other(text):
            if not text or pd.isna(text) or len(str(text).strip()) < 5:
                return "Uncategorized"
        
            text_clean = str(text).replace('ي', 'ی').replace('ك', 'ک')
            for keywords in ASPECTS.values():
                if any(kw in text_clean for kw in keywords):
                    return "Mapped"
            return "Other"

        df['mapping_status'] = df[comment_col].apply(is_other)
    
        # Filter for the requested type and return relevant columns
        unmapped = df[df['mapping_status'] == category_type]
        return unmapped[[branch_col, rating_col, comment_col]]
    
    def get_summary_for_ai(self):
        kpis = self.get_kpis()
        pareto = self.get_pareto_analysis()
        br_stats, _ = self.get_branch_analysis()
        issue_cats = self.get_issue_category_analysis()
        recovery = self.get_recovery_opportunities()
        mom = self.get_mom_comparison()
        
        # 2. Get top issues/strengths safely
        top_issues_df = self.get_top_issues(5)
        top_strengths_df = self.get_top_strengths(5)
        
        return {
            'kpis': kpis,
        
            # Use .empty check for DataFrames
            'top_issues': top_issues_df.to_dict('records') if not top_issues_df.empty else [],
            'top_strengths': top_strengths_df.to_dict('records') if not top_strengths_df.empty else [],
        
            'pareto_top5': pareto.head(5).to_dict('records') if not pareto.empty else [],
        
            # Safe access for Branch Stats
            'best_branch': br_stats.iloc[0].to_dict() if not br_stats.empty else {},
            'worst_branch': br_stats.iloc[-1].to_dict() if not br_stats.empty and len(br_stats) > 1 else {},
        
            'aspect_sentiment': self.get_aspect_sentiment().to_dict('records') if hasattr(self, 'get_aspect_sentiment') else [],
            
            # This now uses the safe DataFrame we fixed in the previous step
            'issue_categories': issue_cats.to_dict('records') if not issue_cats.empty else [],
            
            'customer_segments': recovery.to_dict('records') if not recovery.empty else [],
            
            # Handle Month-over-Month logic
            'latest_month': mom.iloc[-1].to_dict() if not mom.empty else {},
            'mom_trend': 'improving' if not mom.empty and len(mom) > 1 and mom.iloc[-1].get('rating_change', 0) > 0 else 'stable/declining'
        }   
    
    def get_text_column(self):
        """Get the main text/comment column"""
        # Try common column names for comments
        possible_cols = [
        'لطفا نظر و انتفادات خود را برای ما بنویسید',             
        'نظر',            
        'توضیحات',            
            'comment',
            'text'
        ]
        for col in possible_cols:
            if col in self.df.columns:
                return col
        return None

    def preprocess_persian_text(self, text):
        """Clean and normalize Persian text"""
        if pd.isna(text) or not isinstance(text, str):
            return ""
    
        # Normalize if Hazm is available
        if HAZM_AVAILABLE:
            normalizer = Normalizer()
            text = normalizer.normalize(text)
    
        # Remove English characters and numbers
        text = re.sub(r'[a-zA-Z0-9]', '', text)
        # Remove special characters but keep Persian
        text = re.sub(r'[^\u0600-\u06FF\s]', '', text)
        # Remove extra whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text

    def get_persian_stopwords(self):
        """Get Persian stopwords list"""
        custom_stopwords = [
        'و', 'در', 'به', 'از', 'که', 'این', 'را', 'با', 'است', 'برای',
        'آن', 'یک', 'خود', 'تا', 'کرد', 'بر', 'هم', 'نیز', 'گفت', 'می',
        'شد', 'او', 'ما', 'اما', 'یا', 'شده', 'باید', 'هر', 'آنها', 'بود',
        'دارد', 'من', 'دیگر', 'اگر', 'همه', 'وی', 'شود', 'کند', 'چه', 'بی',
        'کنند', 'ها', 'های', 'شما', 'بسیار', 'پس', 'چون', 'زیرا', 'همین',
        'کرده', 'باشد', 'داشت', 'نه', 'هیچ', 'انها', 'توسط', 'سپس', 'ولی',
        'خیلی', 'بودم', 'بودن', 'داشته', 'کردم', 'کنید', 'میشه', 'بشه',
        'خوب', 'بد', 'هست', 'نیست', 'شدم', 'داره', 'نداره', 'کنم', 'کنه',
        'بودند', 'بوده', 'شدن', 'کردن', 'کردند', 'میکنه', 'میشد', 'همچنین'
    ]
    
        if HAZM_AVAILABLE:
            try:
                return list(set(stopwords_list() + custom_stopwords))
            except:
                pass
    
        return custom_stopwords

    def tokenize_text(self, text):
        """Tokenize Persian text"""
        if not text:
            return []
    
        if HAZM_AVAILABLE:
            try:
                return word_tokenize(text)
            except:
                pass
    
        # Fallback: simple split
        return text.split()

    def get_word_frequency(self, min_freq=5, top_n=50):
        """Get word frequency for word cloud"""
        text_col = self.get_text_column()
        if not text_col:
            return {}
    
        stopwords = self.get_persian_stopwords()
        word_counts = Counter()
    
        for _, row in self.df.iterrows():
            text = self.preprocess_persian_text(row.get(text_col, ''))
            tokens = self.tokenize_text(text)
        
            for token in tokens:
                if len(token) > 1 and token not in stopwords:
                    word_counts[token] += 1
    
        # Filter by minimum frequency
        filtered = {k: v for k, v in word_counts.items() if v >= min_freq}
    
        # Return top N
        return dict(Counter(filtered).most_common(top_n))

    def get_ngram_analysis(self, n=2, min_freq=3, top_n=30):
        """Find common n-gram phrases"""
        text_col = self.get_text_column()
        if not text_col:
            return pd.DataFrame()
    
        stopwords = self.get_persian_stopwords()
        ngram_counts = Counter()
    
        for _, row in self.df.iterrows():
            text = self.preprocess_persian_text(row.get(text_col, ''))
            tokens = self.tokenize_text(text)
        
            # Filter stopwords
            tokens = [t for t in tokens if t not in stopwords and len(t) > 1]
            
            # Create n-grams
            for i in range(len(tokens) - n + 1):
                ngram = ' '.join(tokens[i:i+n])
                ngram_counts[ngram] += 1
                
        # Filter and sort
        filtered = [(k, v) for k, v in ngram_counts.items() if v >= min_freq]
        filtered.sort(key=lambda x: x[1], reverse=True)
    
        df_ngrams = pd.DataFrame(filtered[:top_n], columns=['phrase', 'count'])
        return df_ngrams

    def get_keywords_by_rating(self, top_n=20):
        """Find distinctive keywords for each rating level"""
        text_col = self.get_text_column()
        rating_col = COLS['RATING']
        
        if not text_col or rating_col not in self.df.columns:
            return {}
    
        stopwords = self.get_persian_stopwords()
    
        # Group ratings
        rating_groups = {
            'low': [1, 2],      # 1-2 stars
            'mid': [3],         # 3 stars
            'high': [4, 5]      # 4-5 stars
        }
    
        results = {}
        all_words = Counter()
        group_words = {g: Counter() for g in rating_groups}
    
        for _, row in self.df.iterrows():
            text = self.preprocess_persian_text(row.get(text_col, ''))
            rating = row.get(rating_col)
            tokens = self.tokenize_text(text)
        
            for token in tokens:
                if len(token) > 1 and token not in stopwords:
                    all_words[token] += 1
                    for group, ratings in rating_groups.items():
                        if rating in ratings:
                            group_words[group][token] += 1
    
        # Calculate TF-IDF-like score (relative frequency)
        for group, words in group_words.items():
            scored_words = []
            for word, count in words.items():
                if all_words[word] >= 5:  # Min frequency
                # Score = group frequency / total frequency
                    score = count / all_words[word]
                    if count >= 3:  # Min count in group
                        scored_words.append({
                            'word': word,
                            'count': count,
                            'score': round(score, 3),
                            'total': all_words[word]
                        })
        
            # Sort by score (distinctiveness)
            scored_words.sort(key=lambda x: (x['score'], x['count']), reverse=True)
            results[group] = scored_words[:top_n]
    
        return results

    def get_topic_keywords(self, n_topics=5, n_words=10):
        """Simple topic discovery using word co-occurrence"""
        text_col = self.get_text_column()
        if not text_col:
            return []
    
        stopwords = self.get_persian_stopwords()
    
        # Collect all documents as word lists
        documents = []
        for _, row in self.df.iterrows():
            text = self.preprocess_persian_text(row.get(text_col, ''))
            tokens = self.tokenize_text(text)
            tokens = [t for t in tokens if t not in stopwords and len(t) > 1]
            if tokens:
                documents.append(tokens)
    
        if not documents:
            return []
    
        # Simple approach: cluster by seed words
        seed_topics = {
        'کیفیت غذا': ['غذا', 'کیفیت', 'طعم', 'مزه', 'خوشمزه', 'تازه', 'سرد', 'گرم', 'پخت', 'خام'],
        'قیمت و ارزش': ['قیمت', 'گران', 'ارزان', 'حجم', 'اندازه', 'پرس', 'هزینه', 'تخفیف'],
        'تحویل و زمان': ['تاخیر', 'دیر', 'سریع', 'زمان', 'انتظار', 'تحویل', 'ارسال', 'پیک'],
        'بسته‌بندی': ['بسته', 'بندی', 'کارتن', 'جعبه', 'ظرف', 'پاره', 'ریخته', 'نشت'],
        'پرسنل و خدمات': ['پرسنل', 'برخورد', 'رفتار', 'پشتیبان', 'خدمات', 'مودب', 'بی‌ادب']
        }
    
        topic_counts = {topic: Counter() for topic in seed_topics}
    
        for doc in documents:
            doc_text = ' '.join(doc)
            for topic, seeds in seed_topics.items():
                for seed in seeds:
                    if seed in doc_text:
                        for word in doc:
                            topic_counts[topic][word] += 1
                        break
    
        # Format results
        results = []
        for topic, words in topic_counts.items():
            top_words = [w for w, c in words.most_common(n_words) if c >= 3]
            if top_words:
                results.append({
                    'topic': topic,
                    'keywords': top_words,
                    'count': sum(words.values())
                })
    
        results.sort(key=lambda x: x['count'], reverse=True)
        return results

    def get_comment_sentiment_distribution(self):
        """Analyze sentiment distribution of comments"""
        text_col = self.get_text_column()
        rating_col = COLS['RATING']
    
        if not text_col or rating_col not in self.df.columns:
            return pd.DataFrame()
    
        # Positive and negative word lists
        positive_words = [
        'عالی', 'خوب', 'عالیه', 'خوشمزه', 'تازه', 'سریع', 'مودب', 'تمیز',
        'عالی‌بود', 'راضی', 'ممنون', 'متشکر', 'بهترین', 'فوق‌العاده',
        'محشر', 'دوست‌داشتنی', 'لذیذ', 'گرم', 'مناسب', 'حرفه‌ای'
        ]
    
        negative_words = [
        'بد', 'افتضاح', 'سرد', 'دیر', 'گران', 'کم', 'کثیف', 'خراب',
        'بی‌کیفیت', 'ناراضی', 'متاسف', 'بدترین', 'زشت', 'خام', 'سوخته',
        'تاخیر', 'اشتباه', 'بی‌ادب', 'پاره', 'ریخته', 'نامناسب'
        ]
    
        results = []
        for _, row in self.df.iterrows():
            text = str(row.get(text_col, '')).lower()
            rating = row.get(rating_col)
        
            if not text or pd.isna(rating):
                continue
        
            pos_count = sum(1 for w in positive_words if w in text)
            neg_count = sum(1 for w in negative_words if w in text)
        
            if pos_count + neg_count > 0:
                sentiment = 'positive' if pos_count > neg_count else ('negative' if neg_count > pos_count else 'mixed')
            else:
                sentiment = 'neutral'
        
            results.append({
                'rating': rating,
                'sentiment': sentiment,
                'pos_words': pos_count,
                'neg_words': neg_count
            })
    
        df_sent = pd.DataFrame(results)
    
        if len(df_sent) == 0:
            return pd.DataFrame()
    
        # Aggregate
        summary = df_sent.groupby('sentiment').agg(
            count=('rating', 'count'),
            avg_rating=('rating', 'mean')
        ).reset_index()
    
        summary['percentage'] = (summary['count'] / summary['count'].sum() * 100).round(1)
    
        return summary

    def get_rating_sentiment_matrix(self):
        """Cross-tabulation of rating vs detected sentiment"""
        text_col = self.get_text_column()
        rating_col = COLS['RATING']
    
        if not text_col or rating_col not in self.df.columns:
            return pd.DataFrame()
    
        positive_words = ['عالی', 'خوب', 'خوشمزه', 'تازه', 'سریع', 'مودب', 'تمیز', 'راضی', 'ممنون', 'بهترین', 'فوق‌العاده']
        negative_words = ['بد', 'افتضاح', 'سرد', 'دیر', 'گران', 'کم', 'کثیف', 'خراب', 'بی‌کیفیت', 'ناراضی', 'بدترین', 'تاخیر']
    
        results = []
        for _, row in self.df.iterrows():
            text = str(row.get(text_col, '')).lower()
            rating = row.get(rating_col)
        
            if pd.isna(rating):
                continue
        
            pos_count = sum(1 for w in positive_words if w in text)
            neg_count = sum(1 for w in negative_words if w in text)
        
            if pos_count > neg_count:
                sentiment = 'Positive'
            elif neg_count > pos_count:
                sentiment = 'Negative'
            else:
                sentiment = 'Neutral'
        
            results.append({'rating': int(rating), 'sentiment': sentiment})
    
        df = pd.DataFrame(results)
    
        if len(df) == 0:
            return pd.DataFrame()
    
        # Create cross-tab
        matrix = pd.crosstab(df['rating'], df['sentiment'], margins=True)
    
        return matrix

    def get_weekly_trends(self):
        """Aggregate trends by week (ISO week number)"""
        try:
            df = self.df.copy()
            date_col = self.cols.get('CREATED_AT') or self.cols.get('DATE')
            rating_col = self.cols.get('RATING')
            
            if date_col not in df.columns or rating_col not in df.columns:
                return pd.DataFrame()
            
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            df = df.dropna(subset=[date_col, rating_col])
            
            # Create a week label like "2025-W03" for sorting + readability
            df['week'] = df[date_col].dt.strftime('%Y-W%V')
            
            weekly = df.groupby('week').agg(
                avg_rating=(rating_col, 'mean'),
                order_count=(rating_col, 'count')
            ).reset_index()
            
            weekly['avg_rating'] = weekly['avg_rating'].round(2)
            weekly = weekly.sort_values('week')
            
            # Add a rolling 4-week average if enough data
            if len(weekly) >= 4:
                weekly['rating_4week_avg'] = weekly['avg_rating'].rolling(4, min_periods=1).mean().round(2)
            
            return weekly
            
        except Exception as e:
            print(f"Weekly trends error: {e}")
            return pd.DataFrame()

    def get_monthly_trends(self):
        """Aggregate trends by calendar month"""
        try:
            df = self.df.copy()
            date_col = self.cols.get('CREATED_AT') or self.cols.get('DATE')
            rating_col = self.cols.get('RATING')
            
            if date_col not in df.columns or rating_col not in df.columns:
                return pd.DataFrame()
            
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            df = df.dropna(subset=[date_col, rating_col])
            
            # Create month label like "2025-01"
            df['month'] = df[date_col].dt.strftime('%Y-%m')
            
            monthly = df.groupby('month').agg(
                avg_rating=(rating_col, 'mean'),
                order_count=(rating_col, 'count')
            ).reset_index()
            
            monthly['avg_rating'] = monthly['avg_rating'].round(2)
            monthly = monthly.sort_values('month')
            
            # Month-over-month changes
            monthly['rating_change'] = monthly['avg_rating'].diff().round(2)
            monthly['orders_change_pct'] = monthly['order_count'].pct_change().mul(100).round(1)
            
            return monthly
            
        except Exception as e:
            print(f"Monthly trends error: {e}")
            return pd.DataFrame()
    
