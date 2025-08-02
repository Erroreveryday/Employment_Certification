import os
import random
from datetime import datetime, timedelta
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

def read_name_list(file_path):
    """读取List.txt中的姓名列表"""
    with open(file_path, 'r', encoding='utf-8') as f:
        # 读取并过滤空行
        names = [line.strip() for line in f if line.strip()]
    return names

def read_student_info(file_path):
    """读取学生信息Excel文件，返回以姓名为键的字典"""
    # 读取Excel文件
    df = pd.read_excel(file_path)
    
    # 构建学生信息字典
    student_info = {}
    for _, row in df.iterrows():
        name = row['姓名']
        student_info[name] = {
            '性别': row['性别'],
            '身份证号': row['身份证号'],
            '专业': row['专业']
        }
    return student_info

def get_workdays():
    """生成2025年7月上半月和下半月的工作日列表"""
    # 2025年7月1日是星期二
    start_date = datetime(2025, 7, 1)
    
    # 上半月工作日（1-15日）
    first_half_workdays = []
    for i in range(15):
        current_date = start_date + timedelta(days=i)
        # 排除周末（周六和周日）
        if current_date.weekday() < 5:  # 0-4表示周一到周五
            first_half_workdays.append(current_date)
    
    # 下半月工作日（16-31日）
    second_half_workdays = []
    for i in range(16, 32):
        try:
            current_date = datetime(2025, 7, i)
            if current_date.weekday() < 5:
                second_half_workdays.append(current_date)
        except ValueError:
            break  # 处理月份天数不足的情况
    
    return first_half_workdays, second_half_workdays

def replace_placeholders(doc, name, gender, id_card, major, join_date, issue_date):
    """替换文档中的占位符，仅修改目标段落的字体样式"""
    # 需要处理的目标文本片段（用于定位需要修改的段落）
    target_texts = [
        '****，男/女，系湖南理工学院信息科学与工程学院*****专业2025届毕业生（身份证号：******）',
        '**年**月**日',
        '年   月   日'
    ]
    
    # 遍历所有段落
    for para in doc.paragraphs:
        # 仅处理包含目标文本的段落（避免修改标题等其他内容）
        if any(text in para.text for text in target_texts):
            # 替换姓名、性别、专业和身份证号
            if target_texts[0] in para.text:
                new_text = f'{name}，{gender}，系湖南理工学院信息科学与工程学院{major}专业2025届毕业生（身份证号：{id_card}）'
                para.text = para.text.replace(target_texts[0], new_text)
            
            # 替换入职日期
            if target_texts[1] in para.text:
                date_str = join_date.strftime('%Y年%m月%d日')
                para.text = para.text.replace(target_texts[1], date_str)
            
            # 替换落款日期
            if target_texts[2] in para.text:
                date_str = issue_date.strftime('%Y年%m月%d日')
                para.text = para.text.replace(target_texts[2], date_str)
            
            # 仅对当前修改的段落设置字体（宋体三号）
            for run in para.runs:
                run.font.name = '宋体'
                run.font.size = Pt(16)  # 三号字体对应16磅
                # 设置中文字体（确保宋体生效）
                r = run._element
                r.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    
    return doc

def process_certificates(name_list_path, student_info_path, template_path):
    """处理整个流程：读取名单、查找信息、生成证明"""
    # 读取姓名列表
    names = read_name_list(name_list_path)
    print(f"读取到的姓名列表: {names}")
    
    # 读取学生信息
    student_info = read_student_info(student_info_path)
    
    # 获取工作日列表
    first_half, second_half = get_workdays()
    print(f"7月上半月工作日: {[d.strftime('%Y-%m-%d') for d in first_half]}")
    print(f"7月下半月工作日: {[d.strftime('%Y-%m-%d') for d in second_half]}")
    
    # 为每个学生生成用工证明
    for name in names:
        if name not in student_info:
            print(f"警告：未找到{name}的信息，跳过该学生")
            continue
        
        # 获取学生信息
        info = student_info[name]
        print(f"找到{name}的信息: {info}")
        
        # 随机选择日期
        join_date = random.choice(first_half)
        issue_date = random.choice(second_half)
        
        # 读取模板
        doc = Document(template_path)
        
        # 替换占位符
        doc = replace_placeholders(
            doc, 
            name, 
            info['性别'], 
            info['身份证号'], 
            info['专业'], 
            join_date, 
            issue_date
        )
        
        # 创建子文件夹（如果不存在）
        output_dir = "用工证明生成文件"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"已创建子文件夹：{output_dir}")

        # 保存为新文档
        output_path = os.path.join(output_dir, f"{name}.doc")
        doc.save(output_path)
        print(f"已生成：{output_path}")

if __name__ == "__main__":
    # 文件路径
    name_list_path = "名单.txt"
    student_info_path = "学生信息年级总表.xlsx"
    template_path = "用工证明[模板]"
    
    # 检查文件是否存在
    if not os.path.exists(name_list_path):
        print(f"错误：找不到文件 {name_list_path}")
    elif not os.path.exists(student_info_path):
        print(f"错误：找不到文件 {student_info_path}")
    elif not os.path.exists(template_path):
        print(f"错误：找不到文件 {template_path}")
    else:
        # 执行处理
        process_certificates(name_list_path, student_info_path, template_path)