#!/usr/bin/env python3
"""
質問ジェネレーター: ギャップ分析結果から最適な質問を生成
"""
import json
from typing import Dict, List, Any

class QuestionGenerator:
    """回答しやすい質問を生成"""
    
    QUESTION_TEMPLATES = {
        'date': [
            "{name}はいつ頃を予定していますか?\n  [ ] 2025年4月\n  [ ] 2025年7月\n  [ ] 2025年10月\n  [ ] その他: _____",
            "{name}について教えてください: ____年__月頃"
        ],
        'currency': [
            "{name}の概算額を教えてください:\n  [ ] 100万円未満\n  [ ] 100-300万円\n  [ ] 300-500万円\n  [ ] 500万円以上: 約____万円",
            "{name}: 約_____円(概算可)"
        ],
        'number': [
            "{name}について:\n  [ ] 1-2{unit}\n  [ ] 3-5{unit}\n  [ ] 6-10{unit}\n  [ ] 10{unit}以上",
            "{name}: 約_____{unit}"
        ],
        'choice': [
            "{name}について該当するものを選択してください:\n{options}",
            "{name}: _____"
        ],
        'long_text': [
            "{name}について、簡潔にご説明ください(3-5行程度):\n\n_____________________",
        ],
        'short_text': [
            "{name}: _____",
        ]
    }
    
    def generate_questions(self, gap_analysis: Dict, round_num: int = 1) -> Dict[str, Any]:
        """
        質問リストを生成
        
        Args:
            gap_analysis: ギャップ分析結果
            round_num: 質問ラウンド番号(1, 2, 3...)
            
        Returns:
            質問グループのリスト
        """
        if round_num == 1:
            # 必須項目の未記入・部分的項目
            items = self._filter_by_importance(
                gap_analysis['missing'] + gap_analysis['partial'],
                '必須'
            )
        elif round_num == 2:
            # 重要項目の未記入・部分的項目
            items = self._filter_by_importance(
                gap_analysis['missing'] + gap_analysis['partial'],
                '重要'
            )
        else:
            # 任意項目
            items = self._filter_by_importance(
                gap_analysis['missing'],
                '任意'
            )
        
        # 質問をグループ化
        groups = self._group_questions(items)
        
        # 各グループに質問を生成
        result = {
            'round': round_num,
            'groups': []
        }
        
        for group_name, group_items in groups.items():
            questions = []
            for item in group_items:
                question = self._create_question(item)
                questions.append(question)
            
            result['groups'].append({
                'name': group_name,
                'questions': questions
            })
        
        return result
    
    def _filter_by_importance(self, items: List[Dict], importance: str) -> List[Dict]:
        """重要度でフィルタリング"""
        return [item for item in items if item['importance'] == importance]
    
    def _group_questions(self, items: List[Dict]) -> Dict[str, List[Dict]]:
        """関連する質問をグループ化"""
        groups = {}
        
        # グループ化のキーワード
        group_keywords = {
            '事業概要': ['事業', '目的', '内容', '概要'],
            '財務情報': ['金額', '費用', '予算', '売上', '経費'],
            '実施体制': ['体制', '人員', '組織', '役割', '担当'],
            'スケジュール': ['期間', '日程', '時期', 'スケジュール'],
            '効果・実績': ['効果', '成果', '実績', '目標'],
        }
        
        for item in items:
            item_name = item['name']
            assigned = False
            
            for group_name, keywords in group_keywords.items():
                if any(kw in item_name for kw in keywords):
                    if group_name not in groups:
                        groups[group_name] = []
                    groups[group_name].append(item)
                    assigned = True
                    break
            
            if not assigned:
                if 'その他' not in groups:
                    groups['その他'] = []
                groups['その他'].append(item)
        
        return groups
    
    def _create_question(self, item: Dict) -> Dict[str, Any]:
        """個別の質問を生成"""
        field_type = item['type']
        templates = self.QUESTION_TEMPLATES.get(field_type, self.QUESTION_TEMPLATES['short_text'])
        
        # 項目に応じてテンプレートを選択
        if field_type == 'number':
            # 単位を推測
            unit = self._infer_unit(item['name'])
            template = templates[0].format(name=item['name'], unit=unit)
        elif field_type == 'choice':
            # 選択肢を生成(項目名から推測)
            options = self._infer_options(item['name'])
            template = templates[0].format(name=item['name'], options=options)
        else:
            template = templates[0].format(name=item['name'])
        
        question = {
            'item_name': item['name'],
            'question_text': template,
            'field_type': field_type,
            'importance': item['importance'],
            'hint': self._generate_hint(item)
        }
        
        return question
    
    def _infer_unit(self, item_name: str) -> str:
        """数値項目の単位を推測"""
        if '人' in item_name:
            return '名'
        elif '件' in item_name:
            return '件'
        elif '社' in item_name:
            return '社'
        else:
            return ''
    
    def _infer_options(self, item_name: str) -> str:
        """選択項目の選択肢を推測(プレースホルダー)"""
        return "  [ ] 選択肢A\n  [ ] 選択肢B\n  [ ] 選択肢C\n  [ ] その他: _____"
    
    def _generate_hint(self, item: Dict) -> str:
        """回答のヒントを生成"""
        if item.get('found_info'):
            return f"💡 参考: ヒアリングメモに「{item['found_info'][:50]}...」という記載がありました"
        
        if item['importance'] == '必須':
            return "⚠️ この項目は申請書の必須項目です"
        elif item['importance'] == '重要':
            return "📌 この項目は審査で重視されます"
        else:
            return ""


def main():
    """CLIエントリーポイント"""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python question_generator.py <gap_analysis.json> [round_num]")
        sys.exit(1)
    
    with open(sys.argv[1], 'r', encoding='utf-8') as f:
        gap_analysis = json.load(f)
    
    round_num = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    
    generator = QuestionGenerator()
    questions = generator.generate_questions(gap_analysis, round_num)
    
    print(json.dumps(questions, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
