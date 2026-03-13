#!/usr/bin/env python3
"""
テンプレートパーサー: 補助金申請書テンプレートから項目を抽出
"""
import json
import re
from pathlib import Path
from typing import Dict, List, Any

class TemplateParser:
    """申請書テンプレートを解析し、必要項目を構造化"""
    
    def __init__(self):
        self.items = []
        
    def parse_text(self, content: str) -> List[Dict[str, Any]]:
        """
        テキストベースのテンプレートを解析
        
        Args:
            content: テンプレートのテキスト内容
            
        Returns:
            項目リスト
        """
        items = []
        
        # パターン1: 「項目名:」または「項目名:_____」
        pattern1 = r'([^:\n]{2,50}):\s*[_\s]*(?:\n|$)'
        
        # パターン2: 「【項目名】」
        pattern2 = r'【([^】]+)】'
        
        # パターン3: 「■項目名」
        pattern3 = r'■\s*([^\n]{2,50})'
        
        for pattern in [pattern1, pattern2, pattern3]:
            matches = re.finditer(pattern, content)
            for match in matches:
                item_name = match.group(1).strip()
                
                # 重要度を推測
                importance = self._infer_importance(item_name, content)
                
                items.append({
                    'name': item_name,
                    'type': self._infer_field_type(item_name),
                    'importance': importance,
                    'context': self._extract_context(content, match.start(), match.end())
                })
        
        return items
    
    def _infer_importance(self, item_name: str, full_content: str) -> str:
        """項目名から重要度を推測"""
        
        # 必須キーワード
        required_keywords = [
            '必須', '事業名', '申請者', '代表者', '所在地', 
            '事業内容', '事業目的', '補助金額', '総事業費'
        ]
        
        # 重要キーワード
        important_keywords = [
            '効果', '計画', 'スケジュール', '体制', '実績',
            '売上', '経費', '雇用', '設備'
        ]
        
        for keyword in required_keywords:
            if keyword in item_name:
                return '必須'
                
        for keyword in important_keywords:
            if keyword in item_name:
                return '重要'
                
        return '任意'
    
    def _infer_field_type(self, item_name: str) -> str:
        """項目名からフィールドタイプを推測"""
        
        if any(kw in item_name for kw in ['日付', '年月日', '期間', '時期']):
            return 'date'
        elif any(kw in item_name for kw in ['金額', '費用', '円', '予算']):
            return 'currency'
        elif any(kw in item_name for kw in ['人数', '件数', '数']):
            return 'number'
        elif any(kw in item_name for kw in ['選択', 'いずれか', '該当']):
            return 'choice'
        elif any(kw in item_name for kw in ['説明', '内容', '理由', '詳細']):
            return 'long_text'
        else:
            return 'short_text'
    
    def _extract_context(self, content: str, start: int, end: int, window: int = 100) -> str:
        """項目の前後の文脈を抽出"""
        context_start = max(0, start - window)
        context_end = min(len(content), end + window)
        return content[context_start:context_end].strip()


def main():
    """CLIエントリーポイント"""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python template_parser.py <template_file>")
        sys.exit(1)
    
    template_path = Path(sys.argv[1])
    
    if not template_path.exists():
        print(f"Error: File not found: {template_path}")
        sys.exit(1)
    
    content = template_path.read_text(encoding='utf-8')
    
    parser = TemplateParser()
    items = parser.parse_text(content)
    
    print(json.dumps(items, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
