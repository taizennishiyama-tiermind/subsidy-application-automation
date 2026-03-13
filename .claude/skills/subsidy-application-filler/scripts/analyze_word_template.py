#!/usr/bin/env python3
"""
Word Template Analyzer for Subsidy Application Forms

このスクリプトはWordテンプレート(.docx)を分析し、
以下の情報をJSON形式で出力します:
- プレースホルダー({{variable}}形式)
- 表の構造
- スタイル情報
- 動的セクション
"""

import zipfile
import xml.etree.ElementTree as ET
import json
import sys
import re
from pathlib import Path
from collections import defaultdict


class WordTemplateAnalyzer:
    def __init__(self, template_path):
        self.template_path = template_path
        self.analysis = {
            "file_name": Path(template_path).name,
            "placeholders": [],
            "tables": [],
            "styles": {},
            "dynamic_sections": []
        }
        
        # Word XML名前空間
        self.namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
    
    def analyze(self):
        """完全な分析を実行"""
        print(f"📊 Wordテンプレート分析開始: {self.template_path}")
        print("=" * 60)
        
        with zipfile.ZipFile(self.template_path, 'r') as docx:
            # document.xmlを読み込み
            document_xml = docx.read('word/document.xml')
            root = ET.fromstring(document_xml)
            
            # プレースホルダーを検出
            self._find_placeholders(root)
            
            # 表を分析
            self._analyze_tables(root)
            
            # スタイルを分析(オプション)
            try:
                styles_xml = docx.read('word/styles.xml')
                self._analyze_styles(ET.fromstring(styles_xml))
            except KeyError:
                print("  ⚠️ styles.xml が見つかりません")
            
            # 動的セクションを検出
            self._detect_dynamic_sections(root)
        
        return self.analysis
    
    def _find_placeholders(self, root):
        """
        プレースホルダーを検出
        
        対応パターン:
        - {{variable}}
        - {variable}
        - [variable]
        - <<variable>>
        """
        print("\n[プレースホルダー検出]")
        
        # すべてのテキストノードを取得
        text_elements = root.findall('.//w:t', self.namespaces)
        
        placeholder_patterns = [
            r'\{\{([^}]+)\}\}',  # {{variable}}
            r'\{([^}]+)\}',      # {variable}
            r'\[([^\]]+)\]',     # [variable]
            r'<<([^>]+)>>',      # <<variable>>
        ]
        
        placeholders_found = {}
        
        for text_elem in text_elements:
            if text_elem.text:
                for pattern in placeholder_patterns:
                    matches = re.finditer(pattern, text_elem.text)
                    for match in matches:
                        placeholder = match.group(0)
                        variable_name = match.group(1).strip()
                        
                        if placeholder not in placeholders_found:
                            placeholders_found[placeholder] = {
                                "placeholder": placeholder,
                                "variable": variable_name,
                                "pattern": pattern,
                                "count": 1,
                                "suggested_data_path": self._guess_data_path(variable_name)
                            }
                        else:
                            placeholders_found[placeholder]["count"] += 1
        
        self.analysis["placeholders"] = list(placeholders_found.values())
        
        print(f"  検出されたプレースホルダー: {len(placeholders_found)}種類")
        for ph in self.analysis["placeholders"]:
            print(f"    {ph['placeholder']} → {ph['suggested_data_path']} ({ph['count']}箇所)")
    
    def _guess_data_path(self, variable_name):
        """変数名からデータパスを推測"""
        var_lower = variable_name.lower().replace(' ', '_').replace('-', '_')
        
        # 会社情報
        if any(kw in var_lower for kw in ['company', '会社', '企業']):
            if any(kw in var_lower for kw in ['name', '名', '名称']):
                return "company_info.name"
            elif any(kw in var_lower for kw in ['address', '住所', '所在地']):
                return "company_info.address"
            elif any(kw in var_lower for kw in ['tel', '電話', 'phone']):
                return "company_info.tel"
            elif any(kw in var_lower for kw in ['representative', '代表', '社長']):
                return "company_info.representative"
        
        # 事業情報
        if any(kw in var_lower for kw in ['project', '事業', 'プロジェクト']):
            if any(kw in var_lower for kw in ['title', 'name', '名', '名称']):
                return "project_info.title"
            elif any(kw in var_lower for kw in ['purpose', '目的']):
                return "project_info.purpose"
            elif any(kw in var_lower for kw in ['budget', '予算', '金額']):
                return "project_info.total_budget"
        
        # 日付
        if any(kw in var_lower for kw in ['date', '日付', '年月日']):
            if 'start' in var_lower or '開始' in var_lower:
                return "project_info.period_start"
            elif 'end' in var_lower or '終了' in var_lower:
                return "project_info.period_end"
        
        # デフォルト
        return f"UNKNOWN.{var_lower}"
    
    def _analyze_tables(self, root):
        """表の構造を分析"""
        print("\n[表の分析]")
        
        tables = root.findall('.//w:tbl', self.namespaces)
        
        for idx, table in enumerate(tables):
            rows = table.findall('.//w:tr', self.namespaces)
            
            if not rows:
                continue
            
            # ヘッダー行(最初の行)を取得
            header_row = rows[0]
            header_cells = header_row.findall('.//w:tc', self.namespaces)
            
            header_texts = []
            for cell in header_cells:
                text_parts = []
                for t in cell.findall('.//w:t', self.namespaces):
                    if t.text:
                        text_parts.append(t.text)
                header_texts.append(''.join(text_parts).strip())
            
            # プレースホルダーを含む表かチェック
            is_dynamic = False
            for row in rows[1:]:  # データ行をチェック
                for cell in row.findall('.//w:tc', self.namespaces):
                    cell_text = ''.join(t.text or '' for t in cell.findall('.//w:t', self.namespaces))
                    if re.search(r'\{\{|\{|\[|<<', cell_text):
                        is_dynamic = True
                        break
                if is_dynamic:
                    break
            
            table_info = {
                "table_id": f"table_{idx + 1}",
                "row_count": len(rows),
                "column_count": len(header_cells),
                "headers": header_texts,
                "is_dynamic": is_dynamic,
                "note": "動的表: プレースホルダーを含む行を繰り返す" if is_dynamic else "静的表"
            }
            
            self.analysis["tables"].append(table_info)
            
            print(f"  表{idx + 1}: {len(rows)}行 × {len(header_cells)}列")
            print(f"    ヘッダー: {', '.join(header_texts[:3])}...")
            if is_dynamic:
                print(f"    ⚠️ 動的表(プレースホルダー含む)")
    
    def _analyze_styles(self, styles_root):
        """スタイル情報を分析"""
        print("\n[スタイル分析]")
        
        style_elements = styles_root.findall('.//w:style', self.namespaces)
        
        for style in style_elements:
            style_id = style.get('{' + self.namespaces['w'] + '}styleId')
            
            name_elem = style.find('.//w:name', self.namespaces)
            style_name = name_elem.get('{' + self.namespaces['w'] + '}val') if name_elem is not None else "Unnamed"
            
            self.analysis["styles"][style_id] = {
                "name": style_name,
                "type": style.get('{' + self.namespaces['w'] + '}type', 'unknown')
            }
        
        print(f"  スタイル数: {len(self.analysis['styles'])}")
    
    def _detect_dynamic_sections(self, root):
        """
        動的セクション(繰り返しが必要な部分)を検出
        
        連続する段落にプレースホルダーがある場合、
        それを動的セクションと判定
        """
        print("\n[動的セクション検出]")
        
        paragraphs = root.findall('.//w:p', self.namespaces)
        
        current_section = None
        section_start = None
        
        for idx, para in enumerate(paragraphs):
            para_text = ''.join(t.text or '' for t in para.findall('.//w:t', self.namespaces))
            
            # プレースホルダーを含むか
            has_placeholder = bool(re.search(r'\{\{|\{|\[|<<', para_text))
            
            if has_placeholder:
                if current_section is None:
                    # 新しいセクション開始
                    current_section = {
                        "start_paragraph": idx,
                        "placeholders": []
                    }
                    section_start = idx
                
                # プレースホルダーを抽出
                for pattern in [r'\{\{([^}]+)\}\}', r'\{([^}]+)\}']:
                    matches = re.finditer(pattern, para_text)
                    for match in matches:
                        current_section["placeholders"].append(match.group(1).strip())
            
            else:
                # プレースホルダーがない段落
                if current_section is not None:
                    # セクション終了
                    current_section["end_paragraph"] = idx - 1
                    current_section["paragraph_count"] = idx - section_start
                    
                    # 複数段落にまたがる場合のみ記録
                    if current_section["paragraph_count"] > 1:
                        self.analysis["dynamic_sections"].append(current_section)
                    
                    current_section = None
        
        # 最後のセクション処理
        if current_section is not None:
            current_section["end_paragraph"] = len(paragraphs) - 1
            current_section["paragraph_count"] = len(paragraphs) - section_start
            if current_section["paragraph_count"] > 1:
                self.analysis["dynamic_sections"].append(current_section)
        
        if self.analysis["dynamic_sections"]:
            print(f"  動的セクション: {len(self.analysis['dynamic_sections'])}個")
            for sec in self.analysis["dynamic_sections"]:
                print(f"    段落 {sec['start_paragraph']}-{sec['end_paragraph']}: {len(sec['placeholders'])}個のプレースホルダー")
        else:
            print("  動的セクションなし")
    
    def save_to_file(self, output_path):
        """分析結果をJSONファイルに保存"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.analysis, f, indent=2, ensure_ascii=False)
        
        print(f"\n✅ 分析完了: {output_path}")
    
    def print_summary(self):
        """分析結果のサマリーを表示"""
        print("\n" + "=" * 60)
        print("📋 分析サマリー")
        print("=" * 60)
        
        print(f"プレースホルダー: {len(self.analysis['placeholders'])}種類")
        print(f"表: {len(self.analysis['tables'])}個")
        dynamic_tables = sum(1 for t in self.analysis['tables'] if t['is_dynamic'])
        if dynamic_tables > 0:
            print(f"  └─ 動的表: {dynamic_tables}個")
        print(f"動的セクション: {len(self.analysis['dynamic_sections'])}個")
        
        print("\n⚠️ 重要な注意点:")
        print("  1. プレースホルダーは完全一致で置換される")
        print("  2. 動的表は行テンプレートを繰り返す")
        print("  3. スタイルとフォーマットは自動的に保持される")


def main():
    if len(sys.argv) < 2:
        print("使い方: python analyze_word_template.py <テンプレートファイル.docx> [出力先.json]")
        sys.exit(1)
    
    template_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "word_template_analysis.json"
    
    analyzer = WordTemplateAnalyzer(template_path)
    analyzer.analyze()
    analyzer.print_summary()
    analyzer.save_to_file(output_path)


if __name__ == "__main__":
    main()
