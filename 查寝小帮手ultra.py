import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import pandas as pd
import numpy as np
import random
import os
from datetime import datetime

# pip install pandas openpyxl numpy

class UltimateDormScoreApp:
    def __init__(self, root):
        self.root = root
        self.root.title("寝室卫生检查评分系统")
        self.root.geometry("900x850")
        
        self.setup_styles()
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        
        self.create_widgets()
        self.restore_default_dorms()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TButton', padding=5)
        style.configure('Title.TLabel', font=('微软雅黑', 16, 'bold'), foreground='#2E86C1')
        style.configure('Accent.TButton', background='#2E86C1', foreground='white')

    def get_default_dorms(self):
        return "\n".join(["10A201", "10B408", "10B414"])

    def get_current_date(self):
        return datetime.now().strftime("%y.%m.%d")

    def create_widgets(self):
        row = 0
        ttk.Label(self.main_frame, text="寝室卫生检查评分系统", style='Title.TLabel').grid(row=row, column=0, columnspan=2, pady=(0, 20))
        row += 1
        
        fields_frame = ttk.LabelFrame(self.main_frame, text="基础信息", padding="10")
        fields_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(fields_frame, text="检查日期:").grid(row=0, column=0, sticky=tk.W)
        self.date_var = tk.StringVar(value=self.get_current_date())
        ttk.Entry(fields_frame, textvariable=self.date_var, width=15).grid(row=0, column=1, sticky=tk.W, padx=5)

        ttk.Label(fields_frame, text="检查人员:").grid(row=0, column=2, sticky=tk.W, padx=10)
        self.inspector_var = tk.StringVar()
        ttk.Entry(fields_frame, textvariable=self.inspector_var, width=15).grid(row=0, column=3, sticky=tk.W)

        ttk.Label(fields_frame, text="随机种子:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.seed_var = tk.StringVar(value="666")
        ttk.Entry(fields_frame, textvariable=self.seed_var, width=15).grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Label(fields_frame, text="(相同种子生成结果一致)", font=('微软雅黑', 8), foreground='gray').grid(row=1, column=2, columnspan=2, sticky=tk.W)
        row += 1

        # 算法参数配置
        algo_frame = ttk.LabelFrame(self.main_frame, text="评分逻辑配置 (正态分布)", padding="10")
        algo_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(algo_frame, text="最低分:").grid(row=0, column=0)
        self.min_score_var = tk.StringVar(value="84")
        ttk.Entry(algo_frame, textvariable=self.min_score_var, width=8).grid(row=0, column=1, padx=5)

        ttk.Label(algo_frame, text="最高分:").grid(row=0, column=2, padx=5)
        self.max_score_var = tk.StringVar(value="96")
        ttk.Entry(algo_frame, textvariable=self.max_score_var, width=8).grid(row=0, column=3, padx=5)
        row += 1

        # 寝室列表
        ttk.Label(self.main_frame, text="寝室列表 (每行一个):").grid(row=row, column=0, sticky=tk.W, pady=(10, 0))
        row += 1
        self.dorm_text = scrolledtext.ScrolledText(self.main_frame, width=50, height=12)
        self.dorm_text.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        row += 1

        # 按钮组
        btn_frame = ttk.Frame(self.main_frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=5)
        ttk.Button(btn_frame, text="恢复默认寝室", command=self.restore_default_dorms).pack(side=tk.LEFT, padx=5)
        row += 1

        # 保存路径
        path_frame = ttk.Frame(self.main_frame)
        path_frame.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        self.path_var = tk.StringVar(value=os.path.join(os.getcwd(), "寝室卫生检查评分表.xlsx"))
        ttk.Entry(path_frame, textvariable=self.path_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="浏览", command=self.browse_file).pack(side=tk.RIGHT, padx=5)
        row += 1

        # 执行按钮
        ttk.Button(self.main_frame, text="生成评分表格", style='Accent.TButton', command=self.generate_scores).grid(row=row, column=0, columnspan=2, pady=10)
        row += 1

        # 状态栏
        self.status_var = tk.StringVar(value="准备就绪")
        ttk.Label(self.main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W).grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E))

    def restore_default_dorms(self):
        self.dorm_text.delete("1.0", tk.END)
        self.dorm_text.insert("1.0", self.get_default_dorms())

    def browse_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path: self.path_var.set(file_path)

    def generate_scores(self):
        try:
            # 1. 数据校验
            min_s = int(self.min_score_var.get())
            max_s = int(self.max_score_var.get())
            if min_s >= max_s: raise ValueError("最低分必须小于最高分")

            # 2. 初始化环境
            seed_val = int(self.seed_var.get()) if self.seed_var.get() else None
            np.random.seed(seed_val)
            random.seed(seed_val)

            scoring_items = {
                "礼貌": 5, "床上": 10, "床下": 10, "个人衣服": 10, "地面卫生": 10,
                "电线": 10, "窗台": 5, "乱挂杂物字画": 10, "桌椅": 5, "桌面": 10,
                "厕所": 10, "及时清理垃圾": 5
            }
            
            dorms = [line.strip() for line in self.dorm_text.get("1.0", tk.END).split('\n') if line.strip()]
            if not dorms: raise ValueError("寝室列表不能为空")

            # 3. 正态分布参数
            mu = (min_s + max_s) / 2
            sigma = (max_s - min_s) / 6
            
            results = []
            for dorm in dorms:
                # 目标总分生成
                target = np.clip(np.random.normal(mu, sigma), min_s, max_s)
                
                # 初始化分值 (礼貌固定满分)
                scores = {"日期": self.date_var.get(), "检查人员": self.inspector_var.get(), "寝室": dorm, "礼貌": 5}
                current_sum = 5
                
                # 其他项随机赋初值
                other_keys = [k for k in scoring_items.keys() if k != "礼貌"]
                for k in other_keys:
                    val = random.choice([3, 5]) if scoring_items[k] == 5 else random.choice([6, 8, 10])
                    scores[k] = val
                    current_sum += val
                
                # 算法逼近
                for _ in range(50):
                    if abs(current_sum - target) <= 1: break
                    item = random.choice(other_keys)
                    step = 2
                    if current_sum < target and scores[item] + step <= scoring_items[item]:
                        scores[item] += step
                        current_sum += step
                    elif current_sum > target and scores[item] - step >= (6 if scoring_items[item]==10 else 3):
                        scores[item] -= step
                        current_sum -= step
                
                scores["总分"] = current_sum
                results.append(scores)

            # 4. 保存与反馈
            df = pd.DataFrame(results)
            df.to_excel(self.path_var.get(), index=False)
            
            stats = f"成功生成！\n平均分: {df['总分'].mean():.2f}\n标准差: {df['总分'].std():.2f}\n\n各项均分：\n"
            for k in scoring_items.keys():
                stats += f"{k}: {df[k].mean():.2f}  "
            
            messagebox.showinfo("完成", stats)
            self.status_var.set(f"已保存至: {self.path_var.get()}")

        except Exception as e:
            messagebox.showerror("错误", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateDormScoreApp(root)
    root.mainloop()