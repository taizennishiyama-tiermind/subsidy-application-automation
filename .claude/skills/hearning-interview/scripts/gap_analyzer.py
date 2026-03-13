#!/usr/bin/env python3
"""
ギャップアナライザー: ヒアリング内容とテンプレート項目を照合
"""
import json
import re
from typing import Dict, List, Any

class GapAnalyzer:
    """既存情報と必要項目のギャップを分析"""
    
    def __init__(self):
        self.keyword_cache = {}
        
    def analyze(self, template_items: List[Dict], interview_content: str) -> Dict[str, Any]:
        """
        ギャップ分析を実行
        
        Args:
            template_items: テンプレートから抽出した項目リスト
            interview_content: ヒアリング内容のテキスト
            
        Returns:
            分析結果(完了、部分的、未記入の項目分類)
        """
        result = {
            'completed': [],
            'partial': [],
            'missing': [],
            'stats': {
                'total': len(template_items),
                'completed_count': 0,
                'partial_count': 0,
                'missing_count': 0
            }
        }
        
        for item in template_items:
            coverage = self._check_coverage(item, interview_content)
            
            item_with_coverage = {
                **item,
                'coverage_score': coverage['score'],
                'found_info': coverage['found_info']
            }
            
            if coverage['score'] >= 0.8:
                result['completed'].append(item_with_coverage)
                result['stats']['completed_count'] += 1
            elif coverage['score'] >= 0.3:
                result['partial'].append(item_with_coverage)
                result['stats']['partial_count'] += 1
            else:
                result['missing'].append(item_with_coverage)
                result['stats']['missing_count'] += 1
        
        return result
    
    def _check_coverage(self, item: Dict, content: str) -> Dict[str, Any]:
        """
        項目に対する情報カバレッジをチェック
        
        Returns:
            {'score': 0.0-1.0, 'found_info': 抽出された情報}
        """
        item_name = item['name']
        keywords = self._extract_keywords(item_name)
        
        found_sentences = []
        score = 0.0
        
        # キーワードが含まれる文を抽出
        sentences = re.split(r'[。\n]', content)
        for sentence in sentences:
            if any(kw in sentence for kw in keywords):
                found_sentences.append(sentence.strip())
        
        if found_sentences:
            # 項目タイプに応じてスコアを調整
            if item['type'] in ['date', 'currency', 'number']:
                # 具体的な数値/日付が含まれているか
                if re.search(r'\d+', ' '.join(found_sentences)):
                    score = 0.9
                else:
                    score = 0.4
            elif item['type'] == 'choice':
                # 選択肢的な表現があるか
                score = 0.7 if found_sentences else 0.3
            elif item['type'] == 'long_text':
                # 十分な説明文があるか(50文字以上)
                total_length = sum(len(s) for s in found_sentences)
                if total_length >= 50:
                    score = 0.9
                elif total_length >= 20:
                    score = 0.5
                else:
                    score = 0.3
            else:
                # 短文テキスト
                score = 0.8 if found_sentences else 0.0
        
        return {
            'score': score,
            'found_info': ' '.join(found_sentences[:2])  # 最初の2文まで
        }
    
    def _extract_keywords(self, item_name: str) -> List[str]:
        """項目名から検索キーワードを抽出"""
        
        if item_name in self.keyword_cache:
            return self.keyword_cache[item_name]
        
        keywords = []
        
        # 項目名そのもの
        keywords.append(item_name)
        
        # カッコ内のテキスト
        bracket_content = re.findall(r'[「『(（]([^」』)）]+)[」』)）]', item_name)
        keywords.extend(bracket_content)
        
        # 主要な名詞を抽出(簡易版)
        # 「〜について」「〜に関して」などを除去
        cleaned = re.sub(r'(について|に関して|の詳細|の説明)$', '', item_name)
        if cleaned != item_name:
            keywords.append(cleaned)
        
        self.keyword_cache[item_name] = keywords
        return keywords


def main():
    """CLIエントリーポイント"""
    import sys
    
    if len(sys.argv) < 3:
        print("Usage: python gap_analyzer.py <template_items.json> <interview.txt>")
        sys.exit(1)
    
    with open(sys.argv[1], 'r', encoding='utf-8') as f:
        template_items = json.load(f)
    
    with open(sys.argv[2], 'r', encoding='utf-8') as f:
        interview_content = f.read()
    
    analyzer = GapAnalyzer()
    result = analyzer.analyze(template_items, interview_content)
    
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
