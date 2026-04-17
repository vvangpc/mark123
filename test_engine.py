# -*- coding: utf-8 -*-
import os
import sys

from doc_parser import parse_document
from mark_extractor import extract_marks_from_paragraph
from annotator import (
    smart_annotate_section,
    smart_remove_section,
)

def test_file(file_path):
    print(f"==================================================")
    print(f"Testing: {file_path}")
    
    # 1. Parse document
    doc_data = parse_document(file_path)
    paragraphs = doc_data['paragraphs']
    sections = doc_data['sections']
    mark_para = doc_data.get('mark_para')
    doc = doc_data['document']
    
    if not mark_para:
        print("未找到附图标记段落")
        return
        
    marks = extract_marks_from_paragraph(mark_para)
    print(f"提取到 {len(marks)} 个标记: {marks}")
    
    # 2. Add Annotations
    claims_count = 0
    impl_count = 0
    if '权利要求书' in sections:
        claims_count = smart_annotate_section(paragraphs, sections['权利要求书'], marks, mode="claims")
    if '具体实施方式' in sections:
        impl_count = smart_annotate_section(paragraphs, sections['具体实施方式'], marks, mode="implementation")
        
    print(f"标注完成 - 权利要求书: {claims_count} 段, 具体实施方式: {impl_count} 段")
    
    # Save annotated
    annotated_path = file_path.replace('.docx', '_已标注测试.docx')
    doc.save(annotated_path)
    print(f"已经保存至: {annotated_path}")
    
    # 3. Reload and Remove annotations
    print("--------------------------------------------------")
    print("Testing Removal...")
    doc_data2 = parse_document(annotated_path)
    paragraphs2 = doc_data2['paragraphs']
    sections2 = doc_data2['sections']
    doc2 = doc_data2['document']
    
    rem_claims_count = 0
    rem_impl_count = 0
    if '权利要求书' in sections2:
        rem_claims_count = smart_remove_section(paragraphs2, sections2['权利要求书'], marks, mode="claims")
    if '具体实施方式' in sections2:
        rem_impl_count = smart_remove_section(paragraphs2, sections2['具体实施方式'], marks, mode="implementation")
    
    print(f"清洗完成 - 权利要求书: {rem_claims_count} 段, 具体实施方式: {rem_impl_count} 段")
    
    cleaned_path = file_path.replace('.docx', '_已清洗测试.docx')
    doc2.save(cleaned_path)
    print(f"清洗文件保存至: {cleaned_path}")


if __name__ == "__main__":
    files = [
        r"d:\AI\mark123\FE26Y1586-【无标记初稿】一种新能源汽车专用轴系轴承拆解的夹紧定位工装 - 副本.docx",
        r"d:\AI\mark123\LGKWFE26Y0955-【无标记稿】一种动水环境衬砌注浆修复模拟装置及模拟方法.docx"
    ]
    for f in files:
        if os.path.exists(f):
            test_file(f)
        else:
            print(f"Cannot find {f}")
