# -*- coding: utf-8 -*-
"""
Configuration and Bilingual Labels for Shila Dashboard
"""

import os

def get_secret(key, default=""):
    """Get secret from Streamlit Cloud secrets or environment variable."""
    # First try Streamlit secrets (works on Streamlit Cloud)
    try:
        import streamlit as st
        if hasattr(st, 'secrets') and key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    # Then try environment variable (works locally)
    env_val = os.getenv(key)
    if env_val:
        return env_val
    return default

# ==========================================
# PATHS
# ==========================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data", "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
REPORTS_DIR = os.path.join(OUTPUT_DIR, "reports")
NOTEBOOKLM_DIR = os.path.join(OUTPUT_DIR, "notebooklm")

# Create directories if they don't exist
for dir_path in [DATA_DIR, REPORTS_DIR, NOTEBOOKLM_DIR]:
    os.makedirs(dir_path, exist_ok=True)

# ==========================================
# API KEYS & SECRETS (Streamlit Cloud secrets or environment variables)
# ==========================================
ANTHROPIC_API_KEY = get_secret("ANTHROPIC_API_KEY", "")
DASHBOARD_PASSWORD = get_secret("DASHBOARD_PASSWORD", "shila2026")  # Default for development

# ==========================================
# COLUMN MAPPING - UPDATE THESE TO MATCH YOUR DATA
# ==========================================
COLS = {
    'RATING': 'تجربه ای که از این سفارش داشتید چطور بود؟',
    'FEEDBACK_TYPE': 'لطفا حداقل یکی از گزینه ها را انتخاب کنید',
    'STRENGTH': 'دلایل نقاط قوت',
    'WEAKNESS': 'دلایل نقاط ضعف',
    'DELIVERY': 'دلایل نارضایتی شما از پیک ...',
    'PACKAGING': 'دلایل نارضایتی شما از بسته بندی ...',
    'PERSONNEL': 'دلایل نارضایتی شما از پرسنل را خلاصه بنویسید:',
    'NPS': 'چقدر احتمال داره که ما را به دوستان و آشنایان خود معرفی نمایید',
    'COMMENT': 'لطفا نظر و انتفادات خود را برای ما بنویسید',
    'BRANCH': 'نام شعبه',
    'DATE': 'تاریخ',
    'PRODUCT': 'اقلام فاکتور',
    'CREATED_AT': 'Order Created At',
    'ORDER_ITEMS': 'Order Items',
}

# ==========================================
# BILINGUAL LABELS
# ==========================================
LABELS = {
    'en': {
        'app_title': '🍕 Shila Restaurant - QFD Dashboard',
        'language': 'Language',
        'sidebar_title': '⚙️ Settings & Filters',
        'upload_csv': '📁 Upload CSV File',
        'or_select': 'Or select from existing files:',
        'no_files': 'No files in data folder',
        'date_range': '📅 Date Range',
        'branch_filter': '🏪 Branch',
        'product_filter': '🍔 Product',
        'all_branches': 'All Branches',
        'all_products': 'All Products',
        'apply_filters': '🔍 Apply Filters',
        'kpi_section': '📊 Key Performance Indicators',
        'nps_score': 'NPS Score',
        'avg_rating': 'Avg Rating',
        'total_orders': 'Total Orders',
        'promoters': 'Promoters',
        'detractors': 'Detractors',
        'response_rate': 'Response Rate',
        'tab_overview': '📈 Overview',
        'tab_pareto': '📊 Pareto Analysis',
        'tab_kano': '🎨 Kano Model',
        'tab_branches': '🏪 Branch Analysis',
        'tab_aspects': '🎭 Aspect Sentiment',
        'tab_trends': '📅 Time Trends',
        'tab_ai': '🤖 AI Insights',
        'rating_distribution': 'Rating Distribution',
        'nps_distribution': 'NPS Score Distribution',
        'daily_trend': 'Daily Rating Trend',
        'issues_by_damage': 'Issues by Rating Damage (Pareto)',
        'kano_classification': 'Kano Model Classification',
        'branch_comparison': 'Branch Performance Comparison',
        'aspect_sentiment': 'Aspect-Based Sentiment Analysis',
        'top_issues': 'Top Issues',
        'top_strengths': 'Top Strengths',
        'branch_ranking': 'Branch Ranking',
        'ai_title': '🤖 AI-Powered Insights',
        'ai_rule_based': '📋 Automated Analysis',
        'ai_claude': '🧠 Claude AI Analysis',
        'generate_insights': 'Generate AI Insights',
        'api_key_missing': '⚠️ Anthropic API key not set. Add ANTHROPIC_API_KEY to environment variables.',
        'export_section': '📤 Export',
        'export_excel': '📊 Export to Excel',
        'export_pptx': '📽️ Export to PowerPoint',
        'export_notebooklm': '📝 Export for NotebookLM',
        'export_success': '✅ Export successful!',
        'upload_prompt': '👆 Please upload a CSV file or select from existing files to begin.',
        'loading': 'Loading...',
        'no_data': 'No data available for selected filters.',
        'records_loaded': 'records loaded',
        'filtered_records': 'Filtered records',
        'hourly_rating_trend': 'Hourly Rating Trend',
        'mom_comparison': 'Month-over-Month Comparison',
    },
    'fa': {
        'app_title': '🍕 داشبورد QFD رستوران شیلا',
        'language': 'زبان',
        'sidebar_title': '⚙️ تنظیمات و فیلترها',
        'upload_csv': '📁 بارگذاری فایل CSV',
        'or_select': 'یا انتخاب از فایل‌های موجود:',
        'no_files': 'فایلی در پوشه وجود ندارد',
        'date_range': '📅 بازه زمانی',
        'branch_filter': '🏪 شعبه',
        'product_filter': '🍔 محصول',
        'all_branches': 'همه شعب',
        'all_products': 'همه محصولات',
        'apply_filters': '🔍 اعمال فیلترها',
        'kpi_section': '📊 شاخص‌های کلیدی عملکرد',
        'nps_score': 'امتیاز NPS',
        'avg_rating': 'میانگین امتیاز',
        'total_orders': 'کل سفارشات',
        'promoters': 'ترویج‌کنندگان',
        'detractors': 'منتقدین',
        'response_rate': 'نرخ پاسخ‌دهی',
        'tab_overview': '📈 نمای کلی',
        'tab_pareto': '📊 تحلیل پارتو',
        'tab_kano': '🎨 مدل کانو',
        'tab_branches': '🏪 تحلیل شعب',
        'tab_aspects': '🎭 تحلیل جنبه‌ها',
        'tab_trends': '📅 روند زمانی',
        'tab_ai': '🤖 تحلیل هوش مصنوعی',
        'rating_distribution': 'توزیع امتیازات',
        'nps_distribution': 'توزیع امتیاز NPS',
        'daily_trend': 'روند روزانه امتیاز',
        'issues_by_damage': 'مشکلات بر اساس آسیب (پارتو)',
        'kano_classification': 'طبقه‌بندی مدل کانو',
        'branch_comparison': 'مقایسه عملکرد شعب',
        'aspect_sentiment': 'تحلیل احساسات بر اساس جنبه',
        'top_issues': 'مشکلات اصلی',
        'top_strengths': 'نقاط قوت اصلی',
        'branch_ranking': 'رتبه‌بندی شعب',
        'ai_title': '🤖 تحلیل‌های هوش مصنوعی',
        'ai_rule_based': '📋 تحلیل خودکار',
        'ai_claude': '🧠 تحلیل Claude AI',
        'generate_insights': 'تولید تحلیل هوش مصنوعی',
        'api_key_missing': '⚠️ کلید API آنتروپیک تنظیم نشده.',
        'export_section': '📤 خروجی',
        'export_excel': '📊 خروجی اکسل',
        'export_pptx': '📽️ خروجی پاورپوینت',
        'export_notebooklm': '📝 خروجی برای NotebookLM',
        'export_success': '✅ خروجی با موفقیت ذخیره شد!',
        'upload_prompt': '👆 لطفاً یک فایل CSV بارگذاری کنید.',
        'loading': 'در حال بارگذاری...',
        'no_data': 'داده‌ای موجود نیست.',
        'records_loaded': 'رکورد بارگذاری شد',
        'filtered_records': 'رکوردهای فیلتر شده',
        'hourly_rating_trend': 'روند ساعتی',
        'mom_comparison': 'مقایسه ماهانه',
    }
}

# ==========================================
# ASPECT KEYWORDS
# ==========================================
ASPECTS = {
    'Food Quality / کیفیت غذا': ['کیفیت', 'غذا', 'طعم', 'مزه', 'خوشمزه', 'بدمزه', 'بی مزه', 'تازه', 'سرد', 'گرم', 'پخت', 'شور', 'تند', 'نپخته', 'سوخته', 'خمیر', 'مونده', 'چرب', 'ماسیده'],
    'Price / قیمت': ['قیمت', 'گران', 'ارزان', 'هزینه', 'حجم', 'اندازه', 'پرس', 'ارزش', 'کم', 'کوچک'],
    'Delivery / تحویل': ['پیک', 'تحویل', 'ارسال', 'تاخیر', 'سریع', 'دیر', 'رسید', 'برخورد پیک', 'موتور'],
    'Packaging / بسته‌بندی': ['بسته', 'بندی', 'کارتن', 'جعبه', 'ظرف', 'پاره', 'مو', 'کثیف', 'چیدمان', 'له'],
    'Staff / پرسنل': ['پرسنل', 'برخورد', 'رفتار', 'مودب', 'پشتیبان'],
    'Hygiene / بهداشت': ['بهداشت', 'تمیز', 'کثیف', 'سالم'],
    'Accuracy / دقت سفارش': ['اشتباه', 'مغایرت', 'فراموش', 'جابجا', 'نیست', 'کمبود', 'اضافه', 'سس'],
}

# ==========================================
# COLORS
# ==========================================
COLORS = {
    'primary': '#1f77b4',
    'success': '#4caf50',
    'warning': '#ff9800',
    'danger': '#d32f2f',
    'promoter': '#4caf50',
    'passive': '#ff9800',
    'detractor': '#d32f2f',
}

# ==========================================
# STOPWORDS
# ==========================================
STOPWORDS = set([
    'از', 'به', 'با', 'که', 'در', 'و', 'برای', 'را', 'می', 'است',
    'آن', 'این', 'ها', 'های', 'یک', 'بود', 'شد', 'دارد', 'نیست',
    'غذا', 'سفارش', 'رستوران', 'شیلا', 'مرسی', 'ممنون', 'لطفا', 
    # Optional additions:
    'خیلی', 'هم', 'بسیار', 'کاملا', 'واقعا', 'اصلا', 'همه', 'چون',
    'اگر', 'ولی', 'اما', 'البته', 'فقط', 'حتی', 'یا', 'هر', 'چه'
])

# ==========================================
# PRODUCTS TO EXCLUDE (Side Dishes, Drinks, Extras)
# ==========================================
EXCLUDE_PRODUCTS = [
# Cakes & Desserts
    'جار کیک شکلات فندق',
    'پای سیب آمریکایی',
    'شیرینی پذیرایی 600 گرمی',
    'چیزکیک تنوری کوچک',
    'کیک کاراموند مینی',
    'کیک دبل چاکلت مینی 500 گرمی',
    'کیک آلبالویی مینی 500 گرمی',
    'کیک شکلات هلندی مینی 500 گرمی',
    
# Bread & Extras
    'نان همبرگری',
    'نان هات داگی',
    'نان سیر',
    'نان لقمه اضافه',
    'قارچ اضافه پیتزا',
    'پنير پيتزا اضافه',
    
# Noon Packs
    'نون پک ژامبون میکس',
    'نون پک مرغ و ریحون',
    'نون پک رست بیف',
    
# Sauces
    'کچاپ تکنفره اضافه',
    'کچاپ تند تکنفره اضافه',
    'مایونز تکنفره اضافه',
    'سس سیر تکنفره اضافه',
    'فرانسوی تکنفره اضافه',
    'سس سوخاري',
    'سس هالوپينو',
    'سس سزار',
    'سس پنير گودا',
    'سس آلفردو',
    'سس آيولي',
    
# Salads
    'سالاد کلم',
    
# Soft Drinks
    'نوشابه قوطی لیموناد',
    'نوشابه لیموناد 1 لیتری',
    'نوشابه لیوانی',
    'نوشابه خانواده فانتا پرتقالی',
    'نوشابه قوطی کوکا کولا',
    'نوشابه قوطی فانتا پرتقالی',
    'نوشابه قوطی کوکا کولا زیرو',
    'نوشابه خانواده اسپرایت',
    'نوشابه خانواده کوکا کولا زیرو',
    'نوشابه خانواده کوکا کولا',
'نوشابه قوطی اسپرایت',
   
# Non-Alcoholic Beer
    'ماءالشعیر قوطی هی دی استوایی',
    'ماءالشعیر قوطی هی دی لیمو',
    'ماءالشعیر قوطی هی دی ساده',
    'ماءالشعیر قوطی هی دی سیب',
    'ماءالشعیر قوطی هی دی هلو',
    
# Other Drinks
    'دوغ قوطی خوشگوار',
    'آب معدنی کوچک',
]

# Branches to ignore (Non-Shila locations or partners)
EXCLUDE_BRANCHES = [
    'کیک خونه',
    'کشمون',
]

# ==========================================
# SNAPPFOOD FORMAT SUPPORT
# ==========================================

# SnappFood column positions (0-indexed) in Reviews sheet
SNAPPFOOD_COLS = {
    'BRANCH': 1,           # Column B
    'CUSTOMER_NAME': 3,    # Column D
    'REVIEWED_AT': 6,      # Column G
    'ORDER_CODE': 9,       # Column J
    'CREATED_AT': 12,      # Column M
    'RATING': 14,          # Column O
    'REVIEW_TAG': 17,      # Column R
    'COMMENT': 20,         # Column U
    'DELIVERY_COMMENT': 24,# Column Y
    'ORDER_ITEMS': 28      # Column AC
}

# SnappFood issue columns (column index: issue name)
SNAPPFOOD_ISSUES = {
    5: 'برخورد نامناسب پیک',
    7: 'بسته بندی نامناسب',
    10: 'تأخیر',
    13: 'حفظ نشدن دمای سفارش',
    15: 'دریافت وجه اضافه',
    16: 'طولانی بودن آماده‌سازی',
    18: 'عدم تحویل درب خانه',
    22: 'مشکلات بهداشتی',
    23: 'مشکلات بهداشتی پیک',
    25: 'مغایرت با عکس و توضیحات',
    26: 'مغایرت در سفارش',
    27: 'مغایرت در قیمت'
}

# Positive tags in SnappFood (treated as strengths)
SNAPPFOOD_POSITIVE_TAGS = [
    'کیفیت بالا',
    'تناسب حجم و قیمت',
    'بسته‌بندی مناسب',
    'خوش‌برخوردی پیک',
    'سرعت ارسال',
    'رعایت بهداشت'
]