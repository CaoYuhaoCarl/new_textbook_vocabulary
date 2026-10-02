#!/usr/bin/env python3
"""
词汇数据清洗脚本
功能: 统一音标格式、删除空字段、生成清洗报告
"""

import re
from pathlib import Path
from collections import defaultdict

class VocabularyDataCleaner:
    def __init__(self, input_file):
        self.input_file = Path(input_file)
        self.output_file = self.input_file.parent / f"zk_vocabulary_v2024_cleaned.md"
        self.stats = defaultdict(int)

    def clean_data(self):
        """主清洗函数"""
        print("🚀 开始清洗数据...")

        with open(self.input_file, 'r', encoding='utf-8') as f:
            content = f.read()

        original_lines = len(content.split('\n'))

        # 阶段1: 音标格式统一 [] -> //
        content = self._unify_phonetics(content)

        # 阶段2: 删除空派生词行
        content = self._remove_empty_derivatives(content)

        # 阶段3: 删除"搭配: 无"行
        content = self._remove_no_collocations(content)

        # 阶段4: 修复多余空行
        content = self._fix_empty_lines(content)

        # 保存清洗后的文件
        with open(self.output_file, 'w', encoding='utf-8') as f:
            f.write(content)

        cleaned_lines = len(content.split('\n'))
        self.stats['删除的行数'] = original_lines - cleaned_lines

        print(f"✅ 清洗完成！输出文件: {self.output_file}")
        return content

    def _unify_phonetics(self, content):
        """统一音标格式: [音标] -> /音标/"""
        print("  📝 阶段1: 统一音标格式...")

        # 匹配模式: **单词** [音标] [词频: 数字]
        pattern = r'(\*\*[a-zA-Z\-]+\*\*) \[([^\]]+)\] (\[词频: \d+\])'

        def replace_phonetic(match):
            word = match.group(1)
            phonetic = match.group(2)
            freq = match.group(3)
            self.stats['音标格式转换'] += 1
            return f"{word} /{phonetic}/ {freq}"

        content = re.sub(pattern, replace_phonetic, content)

        print(f"    ✓ 转换了 {self.stats['音标格式转换']} 个音标")
        return content

    def _remove_empty_derivatives(self, content):
        """删除空派生词行"""
        print("  🧹 阶段2: 删除空派生词...")

        # 匹配"派生词:"后紧跟空行的情况
        pattern = r'\n派生词:\s*\n\n'

        matches = re.findall(pattern, content)
        self.stats['删除空派生词'] = len(matches)

        content = re.sub(pattern, '\n', content)

        print(f"    ✓ 删除了 {self.stats['删除空派生词']} 个空派生词字段")
        return content

    def _remove_no_collocations(self, content):
        """删除"搭配: 无"行"""
        print("  🧹 阶段3: 删除无效搭配...")

        # 匹配"搭配: 无"整行
        pattern = r'\n搭配: 无\s*\n'

        matches = re.findall(pattern, content)
        self.stats['删除无效搭配'] = len(matches)

        content = re.sub(pattern, '\n', content)

        print(f"    ✓ 删除了 {self.stats['删除无效搭配']} 个无效搭配字段")
        return content

    def _fix_empty_lines(self, content):
        """修复多余的空行（3+连续空行 -> 2个空行）"""
        print("  🔧 阶段4: 修复多余空行...")

        original_content = content
        content = re.sub(r'\n{4,}', '\n\n\n', content)

        if content != original_content:
            self.stats['修复空行'] = 1
            print(f"    ✓ 修复了多余空行")

        return content

    def generate_report(self, content):
        """生成详细清洗报告"""
        print("\n" + "="*50)
        print("📊 数据清洗报告")
        print("="*50)

        # 统计信息
        total_words = len(re.findall(r'\*\*[a-zA-Z\-]+\*\* /', content))
        phonetic_slash = len(re.findall(r'\*\*[a-zA-Z\-]+\*\* /', content))
        phonetic_bracket = len(re.findall(r'\*\*[a-zA-Z\-]+\*\* \[', content))

        print(f"\n📈 清洗统计:")
        print(f"  • 总词汇量: {total_words}")
        print(f"  • 音标格式转换: {self.stats['音标格式转换']} 个")
        print(f"  • 删除空派生词: {self.stats['删除空派生词']} 个")
        print(f"  • 删除无效搭配: {self.stats['删除无效搭配']} 个")
        print(f"  • 删除的总行数: {self.stats['删除的行数']}")

        print(f"\n✅ 验证结果:")
        print(f"  • 斜杠音标 //: {phonetic_slash} 个")
        print(f"  • 方括号音标 []: {phonetic_bracket} 个")

        if phonetic_bracket == 0:
            print(f"  ✓ 所有音标已统一为 // 格式")
        else:
            print(f"  ⚠️ 仍有 {phonetic_bracket} 个方括号音标需要检查")

        print("\n" + "="*50)

        # 保存报告到文件
        report_file = self.input_file.parent / "cleaning_report.txt"
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write("数据清洗报告\n")
            f.write("="*50 + "\n\n")
            f.write(f"原始文件: {self.input_file}\n")
            f.write(f"清洗文件: {self.output_file}\n\n")
            f.write(f"总词汇量: {total_words}\n")
            f.write(f"音标格式转换: {self.stats['音标格式转换']}\n")
            f.write(f"删除空派生词: {self.stats['删除空派生词']}\n")
            f.write(f"删除无效搭配: {self.stats['删除无效搭配']}\n")
            f.write(f"删除的总行数: {self.stats['删除的行数']}\n")

        print(f"\n📄 详细报告已保存至: {report_file}")

if __name__ == "__main__":
    input_file = "/Users/mac/Desktop/new_textbook_vocabulary/zk_vocabulary_v2024_20226.md"

    cleaner = VocabularyDataCleaner(input_file)
    cleaned_content = cleaner.clean_data()
    cleaner.generate_report(cleaned_content)

    print("\n🎉 数据清洗任务完成！")
