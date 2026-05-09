---
name: "workhours-analysis"
description: "Analyzes work hours data, calculates headcount ratio and generates Excel report. Invoke when user provides work hours CSV and asks for headcount analysis or work distribution."
---

# 工时人头占比分析报告生成器

## 功能说明

此Skill用于分析员工工时数据，计算每位成员在各项目上的**人头占比**，并生成格式化的Excel报告。

## 触发条件

当用户提供工时数据文件（CSV格式）并要求：
- 分析人头占比
- 计算项目投入
- 生成工时报告
- 查看人员在不同项目上的投入分布

## 计算逻辑

### 人头占比定义
- **月度人头** = 该月项目工时 ÷ 当月标准工时
- **人头占比** = 某项目人头 ÷ 该成员当月总人头 × 100%

### 标准工时计算
- 每天标准工时：8小时
- 月度标准工时 = 工作日数 × 8小时
- 自动识别中国法定节假日（元旦、春节等）并调整工作日数

## 数据格式要求

输入CSV文件需包含以下列：
| 列名 | 说明 |
|------|------|
| 所属项目 | 项目名称 |
| 任务标题 | 任务描述（可选） |
| 成员 | 员工姓名 |
| 实际工时数 (小时) | 工时数值 |
| 工时执行时间 | 日期时间格式 |

## 报告输出格式

生成的Excel报告包含以下结构：

| 成员 | 项目 | 月份 | 总人头 | 人头分布 |
|------|------|------|--------|----------|
| 陈思翰 | UGG | 1月 | 1.00 | 工时 99h \| 人头 0.59 \| 占比 59% |
| | | 2月 | 0.94 | 工时 48h \| 人头 0.35 \| 占比 38% |
| | 养乐多 | 1月 | 1.00 | 工时 37h \| 人头 0.22 \| 占比 22% |

- 成员、项目使用合并单元格
- 月份、总人头使用红色字体
- 总人头保留2位小数

## 执行脚本

```python
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date
import calendar

# ========== 参数配置 ==========
# 输入文件路径（由用户提供）
INPUT_FILE = r'<用户提供的CSV文件路径>'
# 输出文件路径
OUTPUT_FILE = r'<输出目录>\人头占比分析报告.xlsx'

# ========== 读取数据 ==========
df = pd.read_csv(INPUT_FILE)
df['工时执行时间'] = pd.to_datetime(df['工时执行时间'])
df['月份'] = df['工时执行时间'].dt.to_period('M')

# 获取数据年份
data_year = df['工时执行时间'].dt.year.min()

# ========== 计算每月标准工时（含节假日） ==========
def get_holidays(year):
    """获取指定年份的法定节假日（可根据实际情况调整）"""
    holidays = set()
    # 元旦
    holidays.add(date(year, 1, 1))
    # 春节（示例：假设春节为2月中旬，需根据实际年份调整）
    # 实际使用时需要根据国务院发布的放假通知调整
    holidays.update([
        date(year, 2, 16), date(year, 2, 17), date(year, 2, 18),
        date(year, 2, 19), date(year, 2, 20), date(year, 2, 21), date(year, 2, 22),
    ])
    return holidays

def get_extra_workdays(year):
    """获取调休上班日"""
    extra = set()
    # 春节调休（需根据实际年份调整）
    extra.update([date(year, 2, 15), date(year, 2, 28)])
    return extra

def count_working_days(year, month, holidays, extra_workdays):
    """计算某月的工作日数"""
    _, days_in_month = calendar.monthrange(year, month)
    all_days = [date(year, month, d) for d in range(1, days_in_month + 1)]
    
    working_days = 0
    for d in all_days:
        weekday = d.weekday()  # 0=Mon, 6=Sun
        if weekday < 5:  # 周一到周五
            if d not in holidays:
                working_days += 1
        else:  # 周六周日
            if d in extra_workdays:
                working_days += 1
    return working_days

# 计算各月标准工时
holidays = get_holidays(data_year)
extra_workdays = get_extra_workdays(data_year)

months_info = {}
for month_num in df['月份'].unique():
    month_str = str(month_num)
    y, m = int(month_str[:4]), int(month_str[5:7])
    wd = count_working_days(y, m, holidays, extra_workdays)
    months_info[month_str] = {'working_days': wd, 'standard_hours': wd * 8}

# ========== 按成员、项目、月汇总工时 ==========
monthly_hours = df.groupby(['成员', '所属项目', '月份'])['实际工时数 (小时)'].sum().reset_index()
monthly_hours = monthly_hours[monthly_hours['实际工时数 (小时)'] > 0]

monthly_hours['标准工时'] = monthly_hours['月份'].astype(str).map(lambda x: months_info[x]['standard_hours'])
monthly_hours['月度人头'] = monthly_hours['实际工时数 (小时)'] / monthly_hours['标准工时']

member_monthly_total = monthly_hours.groupby(['成员', '月份'])['实际工时数 (小时)'].sum().reset_index()
member_monthly_total.columns = ['成员', '月份', '成员月度总工时']

monthly_hours = monthly_hours.merge(member_monthly_total, on=['成员', '月份'])
monthly_hours['人头占比'] = (monthly_hours['实际工时数 (小时)'] / monthly_hours['成员月度总工时'] * 100).round(1)

# ========== 生成Excel报告 ==========
wb = Workbook()

# 样式定义
header_font = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
header_fill = PatternFill(start_color='4A7C59', end_color='4A7C59', fill_type='solid')
member_font = Font(name='微软雅黑', bold=True, size=11)
project_font = Font(name='微软雅黑', size=11)
normal_font = Font(name='微软雅黑', size=10)
month_font = Font(name='微软雅黑', size=10, color='C00000')
head_font = Font(name='微软雅黑', size=10, color='C00000')
thin_border = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000')
)
center_align = Alignment(horizontal='center', vertical='center')
left_align = Alignment(horizontal='left', vertical='center')

# 创建Sheet
ws = wb.active
ws.title = '人头占比总览'

headers = ['成员', '项目', '月份', '总人头', '人头分布']
col_widths = [12, 16, 10, 10, 42]

for col_idx, (header, width) in enumerate(zip(headers, col_widths), 1):
    cell = ws.cell(row=1, column=col_idx, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align
    cell.border = thin_border
    ws.column_dimensions[get_column_letter(col_idx)].width = width

# 获取排序后的月份列表
sorted_months = sorted(monthly_hours['月份'].astype(str).unique())
month_labels = {m: m[5:7] + '月' for m in sorted_months}

all_members = sorted(monthly_hours['成员'].unique())

row = 2
for member in all_members:
    member_data = monthly_hours[monthly_hours['成员'] == member].copy()
    projects = sorted(member_data['所属项目'].unique())
    
    member_start_row = row
    
    for project in projects:
        proj_data = member_data[member_data['所属项目'] == project]
        project_start_row = row
        
        for month_str in sorted_months:
            m_data = proj_data[proj_data['月份'].astype(str) == month_str]
            
            if len(m_data) == 0:
                continue
            
            std_h = months_info[month_str]['standard_hours']
            r = m_data.iloc[0]
            total_hours = r['成员月度总工时']
            total_head = round(total_hours / std_h, 2)
            proj_head = round(r['实际工时数 (小时)'] / std_h, 2)
            
            # 月份
            cell_m = ws.cell(row=row, column=3, value=month_labels[month_str])
            cell_m.font = month_font
            cell_m.alignment = center_align
            cell_m.border = thin_border
            
            # 总人头
            cell_h = ws.cell(row=row, column=4, value=total_head)
            cell_h.font = head_font
            cell_h.alignment = center_align
            cell_h.border = thin_border
            cell_h.number_format = '0.00'
            
            # 人头分布
            dist_text = f"工时 {r['实际工时数 (小时)']:.0f}h | 人头 {proj_head:.2f} | 占比 {r['人头占比']:.0f}%"
            cell_d = ws.cell(row=row, column=5, value=dist_text)
            cell_d.font = normal_font
            cell_d.alignment = left_align
            cell_d.border = thin_border
            
            row += 1
        
        # 合并项目名
        n_proj_rows = row - project_start_row
        if n_proj_rows > 1:
            ws.merge_cells(start_row=project_start_row, start_column=2, end_row=project_start_row + n_proj_rows - 1, end_column=2)
        
        cell_p = ws.cell(row=project_start_row, column=2, value=project)
        cell_p.font = project_font
        cell_p.alignment = center_align
        cell_p.border = thin_border
        for r_idx in range(project_start_row, project_start_row + n_proj_rows):
            ws.cell(row=r_idx, column=2).border = thin_border
    
    # 合并成员名
    n_member_rows = row - member_start_row
    if n_member_rows > 1:
        ws.merge_cells(start_row=member_start_row, start_column=1, end_row=member_start_row + n_member_rows - 1, end_column=1)
    
    cell_mb = ws.cell(row=member_start_row, column=1, value=member)
    cell_mb.font = member_font
    cell_mb.alignment = center_align
    cell_mb.border = thin_border
    for r_idx in range(member_start_row, member_start_row + n_member_rows):
        ws.cell(row=r_idx, column=1).border = thin_border

wb.save(OUTPUT_FILE)
print(f"报告已保存: {OUTPUT_FILE}")
```

## 使用步骤

1. **接收用户上传的工时CSV文件**
2. **确认数据格式**：检查是否包含必需的列（所属项目、成员、实际工时数、工时执行时间）
3. **确认输出路径**：询问用户报告保存位置
4. **执行脚本**：替换INPUT_FILE和OUTPUT_FILE参数后运行
5. **返回结果**：向用户提供生成的Excel文件链接

## 注意事项

- 节假日计算基于中国法定节假日，如需调整请修改`get_holidays()`和`get_extra_workdays()`函数
- 如数据跨越多个年份，需要分别计算各年的节假日
- 输出文件格式为.xlsx，需要安装openpyxl库
