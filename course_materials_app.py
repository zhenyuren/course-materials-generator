import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
import shutil
from generate_renyu_materials import RenYuMaterialsGenerator

class CourseMaterialsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("课程期初资料生成工具")
        self.root.geometry("600x400")
        
        # 设置默认路径
        self.default_template_dir = "期初资料1"
        self.default_output_dir = "生成的课程资料"
        
        # 创建主框架
        self.main_frame = tk.Frame(root, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        self.title_label = tk.Label(self.main_frame, text="课程期初资料生成工具", font=("微软雅黑", 16, "bold"))
        self.title_label.pack(pady=10)
        
        # JSON文件选择
        self.json_frame = tk.LabelFrame(self.main_frame, text="JSON文件", padx=10, pady=10)
        self.json_frame.pack(fill=tk.X, pady=10)
        
        self.json_path_var = tk.StringVar()
        self.json_entry = tk.Entry(self.json_frame, textvariable=self.json_path_var, width=50)
        self.json_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.browse_json_btn = tk.Button(self.json_frame, text="浏览", command=self.browse_json_file)
        self.browse_json_btn.pack(side=tk.LEFT, padx=5)
        
        # 模板目录选择
        self.template_frame = tk.LabelFrame(self.main_frame, text="模板目录", padx=10, pady=10)
        self.template_frame.pack(fill=tk.X, pady=10)
        
        self.template_path_var = tk.StringVar(value=self.default_template_dir)
        self.template_entry = tk.Entry(self.template_frame, textvariable=self.template_path_var, width=50)
        self.template_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.browse_template_btn = tk.Button(self.template_frame, text="浏览", command=self.browse_template_dir)
        self.browse_template_btn.pack(side=tk.LEFT, padx=5)
        
        # 输出目录选择
        self.output_frame = tk.LabelFrame(self.main_frame, text="输出目录", padx=10, pady=10)
        self.output_frame.pack(fill=tk.X, pady=10)
        
        self.output_path_var = tk.StringVar(value=self.default_output_dir)
        self.output_entry = tk.Entry(self.output_frame, textvariable=self.output_path_var, width=50)
        self.output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.browse_output_btn = tk.Button(self.output_frame, text="浏览", command=self.browse_output_dir)
        self.browse_output_btn.pack(side=tk.LEFT, padx=5)
        
        # 状态显示
        self.status_var = tk.StringVar(value="就绪")
        self.status_bar = tk.Label(root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 生成按钮
        self.generate_btn = tk.Button(self.main_frame, text="生成课程资料", command=self.generate_materials, 
                                    font=("微软雅黑", 12), bg="#4CAF50", fg="white", padx=20, pady=10)
        self.generate_btn.pack(pady=20)
        
        # 日志文本框
        self.log_frame = tk.LabelFrame(self.main_frame, text="日志", padx=10, pady=10)
        self.log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.log_text = tk.Text(self.log_frame, height=8, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 重定向print函数到日志文本框
        import sys
        self.old_stdout = sys.stdout
        sys.stdout = self
        
    def write(self, text):
        """重定向print输出到日志文本框"""
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)
        
    def flush(self):
        """刷新输出"""
        pass
    
    def browse_json_file(self):
        """浏览JSON文件"""
        file_path = filedialog.askopenfilename(
            title="选择JSON文件",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if file_path:
            self.json_path_var.set(file_path)
    
    def browse_template_dir(self):
        """浏览模板目录"""
        dir_path = filedialog.askdirectory(title="选择模板目录")
        if dir_path:
            self.template_path_var.set(dir_path)
    
    def browse_output_dir(self):
        """浏览输出目录"""
        dir_path = filedialog.askdirectory(title="选择输出目录")
        if dir_path:
            self.output_path_var.set(dir_path)
    
    def generate_materials(self):
        """生成课程资料"""
        json_path = self.json_path_var.get()
        template_dir = self.template_path_var.get()
        output_dir = self.output_path_var.get()
        
        # 验证输入
        if not json_path:
            messagebox.showerror("错误", "请选择JSON文件")
            return
        
        if not os.path.exists(json_path):
            messagebox.showerror("错误", f"JSON文件不存在: {json_path}")
            return
        
        if not os.path.exists(template_dir):
            messagebox.showerror("错误", f"模板目录不存在: {template_dir}")
            return
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 读取JSON数据
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
        except Exception as e:
            messagebox.showerror("错误", f"读取JSON文件失败: {str(e)}")
            return
        
        # 检查JSON格式
        if isinstance(json_data, dict) and 'courses' in json_data:
            course_data = json_data['courses']
        elif isinstance(json_data, list):
            course_data = json_data
        else:
            messagebox.showerror("错误", "JSON文件格式不正确，需要包含courses列表")
            return
        
        # 获取教师姓名
        if course_data:
            teacher_name = course_data[0].get('teacherName', '未知教师')
        else:
            teacher_name = '未知教师'
        
        # 创建教师输出目录
        teacher_output_dir = os.path.join(output_dir, f"{teacher_name}_课程资料")
        os.makedirs(teacher_output_dir, exist_ok=True)
        
        # 使用现有的生成器类
        try:
            # 创建生成器实例
            generator = RenYuMaterialsGenerator()
            # 设置模板目录和输出目录
            generator.template_dir = template_dir
            generator.output_base_dir = teacher_output_dir
            
            # 设置课程数据
            generator.courses = course_data
            
            # 设置元数据
            from datetime import datetime
            generator.metadata = {
                'date': datetime.now().strftime('%Y年%m月%d日'),
                'semester': '2026年春季学期'
            }
            
            # 合并相同课程代码的课程信息
            merged_courses = {}
            for course in course_data:
                course_code = course.get('courseCode', '')
                if course_code not in merged_courses:
                    merged_courses[course_code] = course.copy()
                else:
                    if 'teacherName' in course and course['teacherName']:
                        existing_teachers = merged_courses[course_code].get('teacherName', '')
                        if existing_teachers:
                            merged_courses[course_code]['teacherName'] = f"{existing_teachers}、{course['teacherName']}"
                        else:
                            merged_courses[course_code]['teacherName'] = course['teacherName']
            
            # 为每门课程生成资料
            for course_code, course in merged_courses.items():
                course['formattedSemester'] = generator.metadata['semester']
                
                # 创建课程文件夹（替换特殊字符）
                course_name = course.get('courseName', '')
                safe_course_name = course_name.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                course_folder_name = f"{course_code}_{safe_course_name}"
                output_course_dir = os.path.join(teacher_output_dir, course_folder_name)
                os.makedirs(output_course_dir, exist_ok=True)
                
                # 生成教学大纲
                doc = generator.generate_syllabus(course)
                if doc:
                    teacher_name_course = course.get('teacherName', teacher_name)
                    safe_filename_name = course_name.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                    output_file = os.path.join(output_course_dir, f"2026春-{safe_filename_name}-教学大纲-{teacher_name_course}.docx")
                    doc.save(output_file)
                    print(f"✅ 教学大纲已生成: {output_file}")
                
                # 生成附件3
                wb3 = generator.generate_attachment3(course)
                if wb3:
                    course_leader = generator.get_course_leader(course_code)
                    safe_filename_name = course_name.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                    output_file = os.path.join(output_course_dir, f"{safe_filename_name}-课程组期初教学资料检查情况记录表-{course_leader}.xlsx")
                    wb3.save(output_file)
                    print(f"✅ 附件3已生成: {output_file}")
                
                # 生成附件1
                wb1 = generator.generate_attachment1(course)
                if wb1:
                    department = course.get('department', '智能金融学院')
                    course_leader = generator.get_course_leader(course_code)
                    output_file = os.path.join(output_course_dir, f"{department}-教学大纲审核汇总表-{course_leader}.xlsx")
                    wb1.save(output_file)
                    print(f"✅ 附件1已生成: {output_file}")
            
            print(f"\n🎉 {teacher_name}的课程资料生成完成！")
            messagebox.showinfo("成功", f"{teacher_name}的课程资料已成功生成！")
            
        except Exception as e:
            print(f"❌ 生成失败: {str(e)}")
            messagebox.showerror("错误", f"生成失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CourseMaterialsApp(root)
    root.mainloop()
