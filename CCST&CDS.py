import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from tkinter import font
import re

class ScholarshipSystem:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("计算机科学与技术学院奖学金评定系统 V1.3")
        
        # --- 界面色彩与尺寸配置 ---
        self.bg_color = "#f5f6fa"      # 浅灰色背景
        self.header_color = "#2c3e50"  # 深蓝色头部
        self.accent_color = "#27ae60"  # 绿色动作按钮
        
        self.root.geometry("1000x600")
        self.root.configure(bg=self.bg_color)
        
        # 数据存储容器
        self.zongce_df = None
        self.honor_limits = {}
        self.scholar_limits = {}
        self.results_df = None
        self.tie_break_var = tk.StringVar(value="综测成绩") 

        self.create_widgets()

    def _safe_int_from_cell(self, v, default=0):
        """从Excel单元格解析整数（兼容空格/文本/小数/NaN）"""
        try:
            num = pd.to_numeric(v, errors="coerce")
            if pd.isna(num):
                return default
            return int(num)
        except Exception:
            return default

    def create_widgets(self):
        # 1. 顶部 Header
        header = tk.Frame(self.root, bg=self.header_color, height=70)
        header.pack(fill=tk.X)
        tk.Label(header, text="计算机科学与技术学院（大数据学院）", font=("微软雅黑", 16, "bold"), fg="white", bg=self.header_color).pack(pady=5)
        tk.Label(header, text="奖学金评定自动化系统", font=("微软雅黑", 10), fg="#bdc3c7", bg=self.header_color).pack()

        # 2. 主体内容容器
        main_container = tk.Frame(self.root, bg=self.bg_color)
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 定义通用按钮样式
        btn_opts = {"font": ("微软雅黑", 10), "width": 18, "cursor": "hand2"}

        # --- A. 数据导入区 (左侧) ---
        left_box = tk.LabelFrame(main_container, text=" 1. 数据准备 ", font=("微软雅黑", 11, "bold"), bg="white", padx=20, pady=20)
        left_box.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)

        tk.Button(left_box, text="📂 导入综测明细", command=self.load_zongce, **btn_opts).pack(pady=8, fill=tk.X)
        tk.Button(left_box, text="🏅 导入荣誉名额", command=self.load_honor, **btn_opts).pack(pady=8, fill=tk.X)
        tk.Button(left_box, text="🎓 导入奖金名额", command=self.load_scholar, **btn_opts).pack(pady=8, fill=tk.X)

        # --- B. 核心处理区 (中间) ---
        mid_box = tk.LabelFrame(main_container, text=" 2. 评定引擎 ", font=("微软雅黑", 11, "bold"), bg="white", padx=20, pady=20)
        mid_box.grid(row=0, column=1, sticky="nsew", padx=10, pady=5)

        # 同分排序参考依据
        sort_frame = tk.Frame(mid_box, bg="white")
        sort_frame.pack(fill=tk.X, pady=10)
        tk.Label(sort_frame, text="同分排序参考依据：", bg="white", font=("微软雅黑", 10, "bold")).pack(anchor="w", pady=5)
        
        radio_frame = tk.Frame(sort_frame, bg="white")
        radio_frame.pack(anchor="w", padx=20)
        tk.Radiobutton(radio_frame, text="综测总分优先", variable=self.tie_break_var, value="综测成绩", 
                      bg="white", font=("微软雅黑", 9)).pack(anchor="w", pady=2)
        tk.Radiobutton(radio_frame, text="学业成绩优先", variable=self.tie_break_var, value="学业成绩", 
                      bg="white", font=("微软雅黑", 9)).pack(anchor="w", pady=2)
        
        # 门槛说明
        threshold_frame = tk.Frame(mid_box, bg="white", pady=10)
        threshold_frame.pack(fill=tk.X)
        tk.Label(threshold_frame, text="[门槛说明]", bg="white", fg="#7f8c8d", 
                font=("微软雅黑", 9, "bold")).pack(anchor="w", pady=5)
        
        threshold_text = "• 劳动实践项评分须 > 2.0\n• 其他荣誉项评分须 > 0\n• 身心素质评分须 ≥ 27.0（否则取消评奖评优资格）"
        tk.Label(threshold_frame, text=threshold_text, bg="white", fg="#e67e22", 
                justify=tk.LEFT, font=("微软雅黑", 9), padx=20).pack(anchor="w")

        # 开始评定按钮 - 居中显示
        button_frame = tk.Frame(mid_box, bg="white", pady=20)
        button_frame.pack(fill=tk.X, anchor="center")
        self.calc_btn = tk.Button(button_frame, text="🚀 开始自动评定", command=self.evaluate, 
                                  bg=self.accent_color, fg="white", font=("微软雅黑", 12, "bold"), 
                                  width=20, pady=12)
        self.calc_btn.pack(anchor="center", pady=10)

        # --- C. 结果产出区 (右侧) ---
        right_box = tk.LabelFrame(main_container, text=" 3. 结果产出 ", font=("微软雅黑", 11, "bold"), bg="white", padx=20, pady=20)
        right_box.grid(row=0, column=2, sticky="nsew", padx=10, pady=5)

        self.save_honor_btn = tk.Button(right_box, text="📜 导出荣誉名单", command=self.save_honor, 
                                      state=tk.DISABLED, **btn_opts)
        self.save_honor_btn.pack(pady=8)

        self.save_scholar_btn = tk.Button(right_box, text="💰 导出奖学金名单", command=self.save_scholarship, 
                                        state=tk.DISABLED, **btn_opts)
        self.save_scholar_btn.pack(pady=8)

        self.save_full_btn = tk.Button(right_box, text="📊 导出明细汇总全表", command=self.save_full_results, 
                                     state=tk.DISABLED, **btn_opts)
        self.save_full_btn.pack(pady=8)

        # 优化列宽比例，使布局更加平衡
        main_container.columnconfigure(0, weight=1, minsize=250)
        main_container.columnconfigure(1, weight=2, minsize=400)
        main_container.columnconfigure(2, weight=1, minsize=250)
        # 确保行高自适应内容
        main_container.rowconfigure(0, weight=1)

        # 底部状态栏
        self.status_var = tk.StringVar(value="系统就绪，请先导入数据文件...")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor="w", padx=10, bg="#ecf0f1")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # ================= 业务逻辑层 =================
    
    def _load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv")])
        if not path: return None
        if path.lower().endswith('.csv'):
            # 使用 sep=None + engine=python 自动识别常见分隔符（逗号/分号/制表符等），
            # 同时处理 UTF-8 BOM，避免列名被读成一个整体导致“缺少必要列”。
            try:
                return pd.read_csv(path, sep=None, engine="python", encoding="utf-8-sig")
            except UnicodeDecodeError:
                # 兜底：兼容部分中文系统常见编码
                return pd.read_csv(path, sep=None, engine="python", encoding="gb18030")
        # Excel：有些表是“多行表头/表头不在第1行”，这里做一个简单的自动表头识别
        try:
            return self._read_excel_with_best_header(path)
        except Exception:
            # 兜底：退回默认解析
            return pd.read_excel(path)

    def _normalize_str(self, s):
        s = str(s) if s is not None else ""
        # 删除 BOM / NBSP / 全角空格 / 零宽空格等不可见字符
        s = s.replace('\ufeff', '').replace('\u00a0', '').replace('\u3000', '').replace('\u200b', '')
        # 删除所有空白（含中英文空格、换行、制表等）
        s = re.sub(r'\s+', '', s)
        return s

    def _read_excel_with_best_header(self, path):
        required = {"学号", "姓名", "学业成绩", "身心素质"}
        # 只读前 N 行用于猜测表头行，避免大表卡顿
        raw = pd.read_excel(path, header=None, nrows=30)

        best_row = None
        best_score = -1
        # 尝试将每一行当表头：统计该行里有多少“必需关键字”
        for i in range(min(len(raw), 30)):
            cells = raw.iloc[i].tolist()
            normalized_cells = [self._normalize_str(c) for c in cells]
            score = 0
            for k in required:
                # 允许列名包含关键字（比如“学号（必填）”）
                if any(k in c for c in normalized_cells):
                    score += 1
            if score > best_score:
                best_score = score
                best_row = i
            # 直接命中全部关键字就不用再找了
            if score == len(required):
                best_row = i
                break

        # 如果猜不到，使用默认 header=0
        if best_score <= 0 or best_row is None:
            return pd.read_excel(path)

        df = pd.read_excel(path, header=best_row)
        return df

    def _normalize_df_columns(self, df):
        # 清理列名里的奇怪空格/BOM，确保后续列名匹配稳定。
        # 同时保证列名永远是字符串。
        df.columns = df.columns.astype(str)
        df.columns = (
            df.columns
            .str.replace('\ufeff', '', regex=False)   # BOM
            .str.replace('\u00a0', '', regex=False)   # NBSP
            .str.replace('\u3000', '', regex=False)   # 全角空格
            .str.replace('\u200b', '', regex=False)   # 零宽空格
            .map(self._normalize_str)
        )
        return df

    def _align_zongce_item_columns(self, df):
        """
        将综测表中“列名含标准关键字”的分项对齐为标准列名，避免导出汇总表时出现空列
        （例如模板列名为“文化艺术活动”“文体得分”“美育”等）。
        长名称优先匹配；各校模板差异通过 EXTRA_KEYWORDS 补充。
        """
        # 标准列名（越长越优先，减少误匹配）
        standards = [
            "创新创业", "对外交流", "公益服务", "劳动实践",
            "社会工作", "身心素质", "体育活动", "文化艺术", "学术研究",
            "学业成绩", "自立自强", "综测成绩",
        ]
        # 标准列仍缺失时，允许的「列名须包含」的别名（normalize 后子串）
        extra_keywords = {
            "创新创业": ["创新创业大赛", "科创"],
            "对外交流": ["境外交流", "交流实践"],
            "公益服务": ["志愿服务", "义工"],
            "劳动实践": ["劳动教育", "社会实践"],
            "社会工作": ["学生工作", "社团工作"],
            "身心素质": ["心理素质", "体质", "体育素质"],
            "体育活动": ["体育类", "体育竞赛"],
            # 各校常见「文化艺术」异名列（不含「文化艺术」四字时仍能匹配）
            "文化艺术": [
                "文化艺术", "文体艺术", "文体活动", "文体类", "美育",
                "文艺", "艺术实践", "艺术文化", "校园文化",
            ],
            "学术研究": ["科研", "论文", "课题"],
            "学业成绩": ["学习成绩", "智育成绩", "学业测评分"],
            "自立自强": ["自强自立", "自强", "困难生"],
            "综测成绩": ["综合测评", "测评总分", "综测分"],
        }
        for std in standards:
            if std in df.columns:
                continue
            n_std = self._normalize_str(std)
            n_std_flat = n_std.replace("-", "")
            hints = [self._normalize_str(h) for h in extra_keywords.get(std, [])]
            found = False
            for c in list(df.columns):
                if c in standards:
                    continue
                nc = self._normalize_str(str(c))
                if n_std in nc or (n_std_flat and n_std_flat in nc):
                    df.rename(columns={c: std}, inplace=True)
                    found = True
                    break
                for h in hints:
                    if len(h) >= 2 and h in nc:
                        df.rename(columns={c: std}, inplace=True)
                        found = True
                        break
                if found:
                    break
        return df

    @staticmethod
    def _scholarship_level_order():
        """奖学金等级导出/展示顺序：一等 → 二等 → 三等 → 未获奖"""
        return ["一等奖学金", "二等奖学金", "三等奖学金", "未获奖"]

    def load_zongce(self):
        df = self._load_file()
        if df is not None:
            df = self._normalize_df_columns(df)
            
            # 只检查必要列是否存在
            required_columns = {"学号", "姓名", "学业成绩", "身心素质"}
            cols_set = set(df.columns)
            missing_columns = required_columns - cols_set

            # 如果严格匹配失败：允许“列名包含关键字”这种模板差异
            if missing_columns:
                col_map = {}
                for need in list(required_columns):
                    # 1) 完全匹配
                    exact = [c for c in df.columns if c == need]
                    if exact:
                        col_map[need] = exact[0]
                        continue
                    # 2) 包含匹配（如“学号（必填）”）
                    contains = [c for c in df.columns if need in c]
                    if contains:
                        col_map[need] = contains[0]

                # 若已能把全部关键列映射出来，则重命名并继续
                if len(col_map) == len(required_columns):
                    df = df.rename(columns={v: k for k, v in col_map.items()})
                    missing_columns = set()
                else:
                    cols_preview = "、".join(list(df.columns)[:10])
                    if len(df.columns) > 10:
                        cols_preview += "、..."
                    messagebox.showerror(
                        "格式错误",
                        f"综测表格缺少必要列：{', '.join(sorted(list(required_columns - set(df.columns))))}\n"
                        f"当前列名预览：{cols_preview}"
                    )
                    return
            
            self.zongce_df = df.copy()
            if "总分" in self.zongce_df.columns:
                self.zongce_df.rename(columns={"总分": "综测成绩"}, inplace=True)
            # 导入后即对齐分项列名（与评定引擎一致），导出汇总时「文化艺术」等不会因列名差异为空
            self._align_zongce_item_columns(self.zongce_df)
            self.status_var.set(f"✅ 已载入 {len(df)} 条学生数据")
            messagebox.showinfo("成功", f"综测数据导入成功，共 {len(df)} 条学生记录")

    def load_honor(self):
        df = self._load_file()
        if df is not None:
            df = self._normalize_df_columns(df)
            awards = ["学业优秀", "社会工作", "学术研究", "创新创业", "公益服务", "对外交流", "体育活动", "文化艺术", "劳动实践", "自强自立"]
            
            # 允许列名带额外文字：只要包含关键字就算匹配
            col_map = {}
            for award in awards:
                if award in df.columns:
                    col_map[award] = award
                    continue
                matches = [c for c in df.columns if award in c]
                if matches:
                    col_map[award] = matches[0]

            has_award_columns = len(col_map) > 0
            if not has_award_columns:
                messagebox.showerror(
                    "格式错误",
                    "荣誉名额表缺少必要的奖项列，请确保包含以下至少一个奖项：\n" + "、".join(awards)
                )
                return

            # 直接加载数据，不验证数值类型
            self.honor_limits = {}
            for award, col in col_map.items():
                if pd.notna(df[col].iloc[0]):
                    self.honor_limits[award] = self._safe_int_from_cell(df[col].iloc[0], default=0)
            self.status_var.set(f"✅ 荣誉名额配置完成，共加载 {len(self.honor_limits)} 个奖项名额")
            messagebox.showinfo("成功", f"荣誉名额导入成功，共加载 {len(self.honor_limits)} 个奖项名额")

    def load_scholar(self):
        df = self._load_file()
        if df is not None:
            df = self._normalize_df_columns(df)
            levels = ["一等奖学金", "二等奖学金", "三等奖学金"]
            
            # 允许列名带额外文字：只要包含关键字就算匹配
            col_map = {}
            for level in levels:
                if level in df.columns:
                    col_map[level] = level
                    continue
                matches = [c for c in df.columns if level in c]
                if matches:
                    col_map[level] = matches[0]

            has_level_columns = len(col_map) > 0
            if not has_level_columns:
                messagebox.showerror("格式错误", "奖学金名额表缺少必要的等级列，请确保包含以下至少一个等级：\n" + "、".join(levels))
                return

            # 直接加载数据，不验证数值类型
            self.scholar_limits = {}
            for level, col in col_map.items():
                if pd.notna(df[col].iloc[0]):
                    self.scholar_limits[level] = self._safe_int_from_cell(df[col].iloc[0], default=0)
            self.status_var.set(f"✅ 奖学金名额配置完成，共加载 {len(self.scholar_limits)} 个等级名额")
            messagebox.showinfo("成功", f"奖学金名额导入成功，共加载 {len(self.scholar_limits)} 个等级名额")

    def evaluate(self):
        if self.zongce_df is None:
            messagebox.showwarning("提醒", "请先导入综测成绩表")
            return
        
        if not self.honor_limits:
            messagebox.showwarning("提醒", "请先导入荣誉名额表")
            return
        
        if not self.scholar_limits:
            messagebox.showwarning("提醒", "请先导入奖学金名额表")
            return

        try:
            res = self.zongce_df.copy()
            self._align_zongce_item_columns(res)
            res["学院"] = "计算机科学与技术学院（大数据学院）"
            res["所获荣誉"] = ""
            res["荣誉奖总数"] = 0
            res["奖学金等级"] = "未获奖"
            
            # 添加评奖评优资格判断：身心素质低于27分取消资格
            if "身心素质" in res.columns:
                res["身心素质"] = pd.to_numeric(res["身心素质"], errors="coerce")
            res["有评奖资格"] = res["身心素质"] >= 27
            
            tie_breaker = self.tie_break_var.get()
            
            # 荣誉奖项映射表
            honor_map = {
                "学业优秀": "学业成绩", "社会工作": "社会工作", "学术研究": "学术研究",
                "创新创业": "创新创业", "公益服务": "公益服务", "对外交流": "对外交流",
                "体育活动": "体育活动", "文化艺术": "文化艺术", "劳动实践": "劳动实践",
                "自强自立": "自立自强"
            }

            # ==== 关键修复：把参与比较/排序的列统一转成数值 ====
            # 否则当 Excel 把数字列读成了字符串（如带空格/文本格式），会触发：
            # '>' / '>=' 中 str 与 float 不能比较的异常。
            numeric_cols = {"学业成绩", "身心素质", "综测成绩"}
            numeric_cols.update(set(honor_map.values()))
            # tie_breaker 也可能是这两列之一
            if tie_breaker in ("综测成绩", "学业成绩"):
                numeric_cols.add(tie_breaker)

            for col in numeric_cols:
                if col in res.columns:
                    res[col] = pd.to_numeric(res[col], errors="coerce")

            # 确保综测成绩列存在：如果没匹配到就尝试从其他“总分”列推导
            if "综测成绩" not in res.columns:
                fallback_total_cols = [c for c in res.columns if "总分" in c or "综测" in c and "成绩" in c]
                if fallback_total_cols:
                    res["综测成绩"] = pd.to_numeric(res[fallback_total_cols[0]], errors="coerce")

            # 关键列缺失直接提示（避免后续比较报更难理解的错）
            required_for_eval = ["学业成绩", "身心素质", "综测成绩"]
            missing_for_eval = [c for c in required_for_eval if c not in res.columns]
            if missing_for_eval:
                messagebox.showerror("格式错误", f"综测表格缺少关键列：{', '.join(missing_for_eval)}")
                return

            # 1. 评荣誉奖（应用劳动实践 > 2.0 的新规则）
            for h_name, q in self.honor_limits.items():
                score_col = honor_map.get(h_name)
                if score_col in res.columns and q > 0:
                    # 【逻辑修正】：劳动实践门槛为2，其余为0
                    threshold = 2.0 if h_name == "劳动实践" else 0.0
                    
                    # 只考虑有评奖资格的学生
                    candidates = res[(res[score_col] > threshold) & (res["有评奖资格"] == True)].copy()
                    if not candidates.empty:
                        top = candidates.sort_values(by=[score_col, tie_breaker], ascending=[False, False]).head(q)
                        res.loc[top.index, "所获荣誉"] = res.loc[top.index, "所获荣誉"].apply(
                            # 避免把 NaN 当作“已存在荣誉”
                            lambda x: (f"{x}, {h_name}" if (isinstance(x, str) and x.strip()) else h_name)
                        )
                        res.loc[top.index, "荣誉奖总数"] += 1

            # 2. 评奖学金（基于荣誉奖结果）
            # 学业成绩百分比：百分位名次×100（与综测表导出列名一致，便于核对门槛）
            if "学业成绩" in res.columns:
                # 百分位×100，保留小数便于核对（Excel 中可见差异）
                res["学业成绩百分比"] = (
                    res["学业成绩"].rank(pct=True, method="average") * 100
                ).round(4)
                thresholds = {
                    "一等奖学金": {"pct": 0.20, "honors": 3, "money": 1800},
                    "二等奖学金": {"pct": 0.50, "honors": 2, "money": 1200},
                    "三等奖学金": {"pct": 0.60, "honors": 1, "money": 600}
                }
                for lvl in ["一等奖学金", "二等奖学金", "三等奖学金"]:
                    q_limit = self.scholar_limits.get(lvl, 0)
                    t = thresholds[lvl]
                    # 只考虑有评奖资格的学生（学业百分位：数值越高名次越好，一等需前20%即>=80）
                    cond = (res["奖学金等级"] == "未获奖") & (res["学业成绩百分比"] >= (1 - t["pct"]) * 100) & (res["荣誉奖总数"] >= t["honors"]) & (res["有评奖资格"] == True)
                    winners = res[cond].sort_values(by=["综测成绩", "荣誉奖总数"], ascending=[False, False]).head(q_limit)
                    res.loc[winners.index, "奖学金等级"] = lvl
                    res.loc[winners.index, "金额"] = t["money"]

            res["金额"] = res.get("金额", 0).fillna(0)
            
            # 添加综测排名（按综测成绩降序排名，同分同排名，连续排名）
            res = res.sort_values(by="综测成绩", ascending=False)
            res["综测排名"] = res["综测成绩"].rank(method="min", ascending=False).astype(int)
            
            # 按奖学金等级：一等 → 二等 → 三等 → 未获奖，同等级按综测成绩降序
            lvl_order = self._scholarship_level_order()
            res["奖学金等级"] = pd.Categorical(res["奖学金等级"], categories=lvl_order, ordered=True)
            self.results_df = res.sort_values(by=["奖学金等级", "综测成绩"], ascending=[True, False])

            self.status_var.set("🎉 评定计算成功完成")
            # 启用所有导出按钮
            for btn in [self.save_honor_btn, self.save_scholar_btn, self.save_full_btn]:
                btn.config(state=tk.NORMAL)
            messagebox.showinfo("完成", "系统已完成所有评定逻辑计算")

        except Exception as e:
            messagebox.showerror("逻辑错误", f"计算过程中发生错误: {e}")

    # ================= 导出层 =================

    def save_honor(self):
        if self.results_df is None: return
        df = self.results_df[self.results_df["荣誉奖总数"] > 0].copy()
        df["荣誉称号"] = df["所获荣誉"].str.split(", ")
        df_exploded = df.explode("荣誉称号")
        
        # 只保留需要的列：班级、学号、姓名、荣誉奖类型
        # 确保班级列存在，如果不存在则使用空值填充
        if "班级" not in df_exploded.columns:
            df_exploded["班级"] = ""
        
        # 选择指定列并按荣誉称号排序，将学号作为第一列
        honor_cols = ["学号", "班级", "姓名", "荣誉称号"]
        df_result = df_exploded[honor_cols].sort_values(by=["荣誉称号", "姓名"])
        
        # 重命名列名以符合用户要求
        df_result.rename(columns={"荣誉称号": "荣誉奖类型"}, inplace=True)
        
        self._save_df(df_result, "荣誉奖获奖名单")

    def save_scholarship(self):
        if self.results_df is None: return
        sl = self.results_df["奖学金等级"].astype(str)
        df = self.results_df[sl != "未获奖"].copy()
        
        # 只保留需要的列：班级、学号、姓名、金额
        # 确保班级列存在，如果不存在则使用空值填充
        if "班级" not in df.columns:
            df["班级"] = ""
        
        # 顺序：一等奖 → 二等奖 → 三等奖，同级按综测成绩降序
        lvl_order = ["一等奖学金", "二等奖学金", "三等奖学金"]
        df["奖学金等级"] = pd.Categorical(df["奖学金等级"], categories=lvl_order, ordered=True)
        df = df.sort_values(by=["奖学金等级", "综测成绩"], ascending=[True, False])
        scholar_cols = ["学号", "班级", "姓名", "奖学金等级", "金额"]
        df_result = df[scholar_cols].copy()
        df_result["奖学金等级"] = df_result["奖学金等级"].astype(str)
        
        self._save_df(df_result, "奖学金获奖名单")

    def save_full_results(self):
        if self.results_df is None: return
        # 只保留用户指定的列
        df = self.results_df.copy()
        
        # 用户指定的保留列，添加综测排名，将学号作为第一列
        keep_columns = [
            "学号", "姓名", "班级", "创新创业", "对外交流", "公益服务", "劳动实践", 
            "社会工作", "身心素质", "体育活动", 
            "文化艺术", "学术研究", "学业成绩", "学业成绩百分比", 
            "自立自强", "综测成绩", "综测排名", "所获荣誉", "荣誉奖总数", 
            "奖学金等级", "金额"
        ]
        
        # 确保所有指定列存在，如果不存在则添加空列
        for col in keep_columns:
            if col not in df.columns:
                df[col] = ""
        
        # 分类列导出为普通文本，避免 Excel 里看起来「仍是旧表」
        if "奖学金等级" in df.columns:
            df["奖学金等级"] = df["奖学金等级"].astype(str)
        if "学业成绩百分比" in df.columns:
            df["学业成绩百分比"] = pd.to_numeric(df["学业成绩百分比"], errors="coerce").round(4)

        # 选择并只保留指定的列
        df_result = df[keep_columns]
        self._save_df(df_result, "综测评定汇总全表")

    def _save_df(self, df, name):
        """保存DataFrame到Excel文件"""
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"{name}.xlsx")
        if path:
            df.to_excel(path, index=False)
            messagebox.showinfo("导出成功", f"文件已保存至：\n{path}")
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    import traceback
    import sys
    try:
        ScholarshipSystem().run()
    except Exception as e:
        traceback.print_exc()
        try:
            _root = tk.Tk()
            _root.withdraw()
            messagebox.showerror(
                "启动失败",
                f"{e}\n\n若从命令行运行，上方应有详细报错。\n"
                "提示：文件名含 & 时请使用引号：python \"CCST&CDS.py\"",
            )
        except Exception:
            sys.exit(1)