# -*- coding: utf-8 -*-
"""AI Insights Module - Rule-based and Claude API analysis"""

import os
from config import ANTHROPIC_API_KEY

try:
    from anthropic import Anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False


class InsightsGenerator:
    def __init__(self, lang='en'):
        self.lang = lang
        self.client = Anthropic(api_key=ANTHROPIC_API_KEY) if HAS_ANTHROPIC and ANTHROPIC_API_KEY else None
    
    def generate_rule_based_insights(self, summary):
        insights = []
        kpis = summary.get('kpis', {})
        nps = kpis.get('nps_score', 0)
        avg_rating = kpis.get('avg_rating', 0)
        
        # NPS insights
        if nps >= 50: insights.append(self._t('nps_excellent', nps))
        elif nps >= 20: insights.append(self._t('nps_good', nps))
        elif nps >= 0: insights.append(self._t('nps_average', nps))
        else: insights.append(self._t('nps_critical', nps))
        
        # Rating insights
        if avg_rating >= 4.5: insights.append(self._t('rating_excellent', avg_rating))
        elif avg_rating >= 4.0: insights.append(self._t('rating_good', avg_rating))
        elif avg_rating >= 3.5: insights.append(self._t('rating_average', avg_rating))
        else: insights.append(self._t('rating_critical', avg_rating))
        
        # Top issue
        top_issues = summary.get('top_issues', [])
        if top_issues:
            insights.append(self._t('top_issue', top_issues[0].get('Issue', 'Unknown')))
        
        return insights
    
    def _t(self, key, *args):
        texts = {
            'en': {
                'nps_excellent': f"🌟 **Excellent NPS ({args[0]})**: Customers are highly loyal!",
                'nps_good': f"✅ **Good NPS ({args[0]})**: Solid loyalty.",
                'nps_average': f"⚠️ **Average NPS ({args[0]})**: Focus on improving.",
                'nps_critical': f"🚨 **Critical NPS ({args[0]})**: Urgent action needed.",
                'rating_excellent': f"🌟 **Excellent Rating ({args[0]})**: Outstanding!",
                'rating_good': f"✅ **Good Rating ({args[0]})**: Above average.",
                'rating_average': f"⚠️ **Average Rating ({args[0]})**: Improvements needed.",
                'rating_critical': f"🚨 **Low Rating ({args[0]})**: Urgent intervention required.",
                'top_issue': f"🔴 **Top Issue**: '{args[0]}' is most mentioned.",
            },
            'fa': {
                'nps_excellent': f"🌟 **NPS عالی ({args[0]})**: مشتریان وفادار!",
                'nps_good': f"✅ **NPS خوب ({args[0]})**: وفاداری خوب.",
                'nps_average': f"⚠️ **NPS متوسط ({args[0]})**: نیاز به بهبود.",
                'nps_critical': f"🚨 **NPS بحرانی ({args[0]})**: اقدام فوری لازم است.",
                'rating_excellent': f"🌟 **امتیاز عالی ({args[0]})**: فوق‌العاده!",
                'rating_good': f"✅ **امتیاز خوب ({args[0]})**: بالاتر از میانگین.",
                'rating_average': f"⚠️ **امتیاز متوسط ({args[0]})**: بهبود لازم است.",
                'rating_critical': f"🚨 **امتیاز پایین ({args[0]})**: مداخله فوری لازم است.",
                'top_issue': f"🔴 **مشکل اصلی**: '{args[0]}' بیشترین گزارش.",
            }
        }
        return texts.get(self.lang, texts['en']).get(key, '')
    
    def generate_claude_insights(self, summary, custom_question=None):
        if not self.client:
            return {'success': False, 'error': 'API not configured', 'insights': ''}
        
        context = self._build_context(summary)
        prompt = f"""Analyze this restaurant feedback data:

{context}

{"User Question: " + custom_question if custom_question else "Provide: 1) Executive Summary, 2) Top 3 Actions, 3) Quick Wins, 4) Risks"}

Respond in {'Persian' if self.lang == 'fa' else 'English'}. Be specific and actionable."""

        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-20250514", max_tokens=2000,
                messages=[{"role": "user", "content": prompt}]
            )
            return {'success': True, 'insights': response.content[0].text}
        except Exception as e:
            return {'success': False, 'error': str(e), 'insights': ''}
    
    def _build_context(self, summary):
        kpis = summary.get('kpis', {})
        ctx = f"NPS: {kpis.get('nps_score')}, Rating: {kpis.get('avg_rating')}/5, Orders: {kpis.get('total_orders')}\n"
        ctx += "Top Issues: " + ", ".join([i.get('Issue', '') for i in summary.get('top_issues', [])]) + "\n"
        ctx += "Top Strengths: " + ", ".join([s.get('Strength', '') for s in summary.get('top_strengths', [])])
        return ctx


def get_api_setup_instructions():
    return """
## Setting Up Claude AI

1. Get API key from [console.anthropic.com](https://console.anthropic.com/)
2. Set environment variable:
   - Windows: `set ANTHROPIC_API_KEY=your-key`
   - Mac/Linux: `export ANTHROPIC_API_KEY=your-key`
3. Restart the dashboard
"""
