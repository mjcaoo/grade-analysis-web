# -*- coding: utf-8 -*-
# @Author  : mjcaoo
# @File    : grade_analyzer.py
# @Time    : 2025/9/13

import pandas as pd
import numpy as np
import re

class GradeAnalyzer:
    """成绩分析器类"""
    
    def __init__(self):
        self.main_courses = []
    
    def convert_grade_to_score(self, grade, level):
        """
        转换成绩：
        若为0或无效成绩，返回0
        若为优秀、良好、中等、及格、不及格，分别转换为95, 85, 75, 60, 0
        若为数字，直接返回数字（限制在0-100之间）
        """
        if pd.isna(grade):
            return 0
        
        if isinstance(grade, str):
            grade = grade.strip()
            level_grades = [[95, 85, 75, 65, 55], [90, 80, 70, 60, 50]]
            level_id = 1 if level >= 2023 else 0
            if grade == '优秀':
                return level_grades[level_id][0]
            elif grade == '良好':
                return level_grades[level_id][1]
            elif grade == '中等':
                return level_grades[level_id][2]
            elif grade == '及格':
                return level_grades[level_id][3]
            elif grade == '不及格':
                return level_grades[level_id][4]
            else:
                # 尝试转换为数字
                try:
                    return float(grade)
                except:
                    return 0
        
        # 如果是数字类型
        try:
            score = float(grade)
            # 限制分数在合理范围内
            if score > 100:
                return 0  # 超过100分的认为是无效成绩
            return score if score > 0 else 0
        except:
            return 0

    def extract_credits_from_course_name(self, course_name):
        """
        从课程名称中提取学分数
        例如：花卉栽培与环境 【1.0】 -> 1.0
        """
        if pd.isna(course_name):
            return 0
        
        # 使用正则表达式匹配【数字】格式
        match = re.search(r'【(\d+\.?\d*)】', str(course_name))
        if match:
            return float(match.group(1))
        return 0

    def load_main_courses(self, main_course_file_path):
        """
        加载主要课程列表
        """
        try:
            df = pd.read_excel(main_course_file_path)
            if '主要课程' in df.columns:
                self.main_courses = df['主要课程'].tolist()
            else:
                # # 如果没有'主要课程'列，尝试其他可能的列名
                # possible_columns = ['课程名称', '课程', 'course', 'Course']
                # for col in possible_columns:
                #     if col in df.columns:
                #         self.main_courses = df[col].tolist()
                #         break
                print("警告：主要课程列表文件中未找到'主要课程'列，主要课程列表为空")
                self.main_courses = []
            
            # 清理课程名称
            self.main_courses = [str(course).strip() for course in self.main_courses if pd.notna(course)]
            print(f"成功加载 {len(self.main_courses)} 门主要课程")
            
        except Exception as e:
            print(f"加载主要课程列表失败: {e}")
            self.main_courses = []

    def process_combined_data(self, file_paths):
        """
        合并处理多个学期的数据
        """
        print(f"\n=== 合并处理 {len(file_paths)} 个文件的数据 ===")
        
        # 存储所有学生的课程数据
        all_students_data = {}
        file_idx = 0
        
        for file_path in file_paths:
            print(f"读取文件: {file_path}")
            file_idx += 1

            try:
                # 读取Excel文件
                df = pd.read_excel(file_path)
                print(f"数据形状: {df.shape}")
                
                # 查找学号列
                student_id_col = None
                for col in df.columns:
                    if '学号' in str(col):
                        student_id_col = col
                        break
                
                if student_id_col is None:
                    print(f"错误：在文件 {file_path} 中找不到学号列")
                    continue
                
                # 查找年级列
                level_col = None
                for col in df.columns:
                    if '年级' in str(col):
                        level_col = col
                        break

                # 查找姓名列
                student_name_col = None
                for col in df.columns:
                    if '姓名' in str(col):
                        student_name_col = col
                        break
                
                # 获取课程列（包含【】的列）
                course_columns = []
                for col in df.columns:
                    if '【' in str(col) and '】' in str(col):
                        # 提取学分确保是有效课程
                        credits = self.extract_credits_from_course_name(col)
                        if credits > 0:
                            course_columns.append(col)
                
                print(f"找到有效课程: {len(course_columns)}门")
                
                # 处理每个学生的数据
                for index, row in df.iterrows():
                    student_id = row[student_id_col]
                    
                    # 跳过无效的学号
                    if pd.isna(student_id):
                        continue

                    level = row[level_col] if level_col and not pd.isna(row[level_col]) else '未知'
                    level = int(str(level).strip())
                    
                    # 初始化学生数据
                    if student_id not in all_students_data:
                        student_name = row[student_name_col] if student_name_col and not pd.isna(row[student_name_col]) else f"学生{student_id}"
                        all_students_data[student_id] = {
                            'name': student_name,
                            'level': level,
                            'courses': {}
                        }
                    
                    # 收集该学生的课程成绩
                    for course_col in course_columns:
                        # 提取课程学分
                        credits = self.extract_credits_from_course_name(course_col)
                        if credits == 0:
                            continue
                            
                        # 获取成绩
                        grade = row[course_col]
                        score = self.convert_grade_to_score(grade, level)
                        
                        # 成绩为0或没有成绩说明该学生未学习该课程
                        if score > 0:
                            course_name = str(course_col).split('【')[0].strip()
                            
                            # 为每个课程创建唯一的标识符，包含原始列名以区分不同学期的同名课程
                            unique_course_key = f"{course_name}_{file_idx}"
                            
                            all_students_data[student_id]['courses'][unique_course_key] = {
                                'name': course_name,  # 保存原始课程名称用于分类
                                'score': score,
                                'credits': credits,
                                'original_name': course_col
                            }
                            
            except Exception as e:
                print(f"处理文件 {file_path} 时出错: {e}")
                continue
        
        # 计算每个学生的综合成绩
        results = []
        for student_id, student_data in all_students_data.items():
            total_weighted_score = 0
            total_credits = 0
            course_count = 0
            valid_courses = []
            
            # 分别收集主要课程和其他课程
            major_courses = []  # 主要课程
            other_courses = []  # 其他课程
            
            for course_key, course_data in student_data['courses'].items():
                course_info = {
                    'name': course_data['name'],  # 使用保存的原始课程名称
                    'score': course_data['score'],
                    'credits': course_data['credits'],
                    'original_name': course_data['original_name']
                }
                
                # 判断是否为主要课程（使用原始课程名称进行判断）
                is_major_course = False
                for major_course in self.main_courses:
                    if major_course in course_data['name']:
                        is_major_course = True
                        break
                
                if is_major_course:
                    major_courses.append(course_info)
                else:
                    other_courses.append(course_info)
            
            # 计算主要课程的学分和成绩
            for course in major_courses:
                total_weighted_score += course['score'] * course['credits']
                total_credits += course['credits']
                course_count += 1
                valid_courses.append(f"{course['name']}({course['score']}分,{course['credits']}学分)[主要课程]")
            
            # 对其他课程按成绩排序，只取前4门
            other_courses.sort(key=lambda x: x['score'], reverse=True)
            if len(other_courses) > 4:
                print(f"注意：学生 {student_id} 的其他课程超过4门，仅取成绩最高的4门")
                selected_other_courses = other_courses[:4]  # 只取成绩最高的4门
            else:
                selected_other_courses = other_courses

            for course in selected_other_courses:
                total_weighted_score += course['score'] * course['credits']
                total_credits += course['credits']
                course_count += 1
                valid_courses.append(f"{course['name']}({course['score']}分,{course['credits']}学分)[其他课程]")
            
            # 计算平均学分绩点
            weighted_avg = total_weighted_score / total_credits if total_credits > 0 else 0

            results.append({
                '学号': student_id,
                '姓名': student_data['name'],
                '年级': student_data['level'],
                '修读课程数': course_count,
                '总学分': total_credits,
                '学分加权平均分': round(weighted_avg, 2),
                '课程详情': "; ".join(valid_courses)
            })
        
        print(f"处理完成，共{len(results)}名学生")
        
        # 创建DataFrame并按照学分加权平均分降序排列
        df_results = pd.DataFrame(results)
        if not df_results.empty:
            df_results = df_results.sort_values(by='学分加权平均分', ascending=False)
        
        return df_results
    

if __name__ == "__main__":
    # 示例用法
    data_dir = "data"
    main_course_files = ["主要课程列表-计算机拔尖221.xlsx", "主要课程列表-数据科学231.xlsx"]
    data_files = [["计算机（拔尖）221-1.xlsx", "计算机（拔尖）221-2.xlsx"],
                  ["数据科学231-1.xlsx", "数据科学231-2.xlsx"]]
    test_case = 0  # 选择测试用例索引
    main_course_file = f"{data_dir}\\{main_course_files[test_case]}"
    analyzer = GradeAnalyzer()
    analyzer.load_main_courses(main_course_file)
    data_files = [f"{data_dir}\\{f}" for f in data_files[test_case]]
    combined_df = analyzer.process_combined_data(data_files)
    # 保存Excel结果
    result_path = f"{data_dir}\\成绩分析结果-{main_course_files[test_case]}.xlsx"
    with pd.ExcelWriter(result_path, engine='openpyxl') as writer:
        combined_df.to_excel(writer, sheet_name='综合成绩', index=False)