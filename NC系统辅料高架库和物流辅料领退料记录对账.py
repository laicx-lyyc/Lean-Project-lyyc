import os
import pandas as pd
import tkinter as tk

# 用于标记一区物流系统和二区物流系统的对账结果是否正常
is_normal_1 = True
is_normal_2 = True

# 作者信息
def show_personal_info_normal():
    info = '''
    作者: 卷包一乙赖灿兴
    版本: 1.0
    日期: 2025年2月20日
    电子邮件: laicx@stu.xmu.edu.cn
    联系方式: 18046307080
    代码用途: 物流系统与NC系统的辅料出入库对账
    描述: 该程序用于物流系统与NC系统的辅料出入库对账，并给出疑似异常数据和自动判断是否真实存在异常数据。
    
    本次对账没有异常数据，您可以打开输出的对账表查看~ 
    
    如有问题和需要变更需求请及时与我取得联系QAQ~ 
    
    点关闭后程序执行完毕...
    请于'./NC_Logistics_Reconciliation/_internal'中提取输出的表格
    '''
    
    # 创建一个 Tkinter 窗口
    window = tk.Tk()
    window.title("个人信息")
    
    # 创建一个 Label 显示个人信息
    label = tk.Label(window, text=info, padx=10, pady=10)
    label.pack()
    
    # 添加一个按钮来关闭窗口
    button = tk.Button(window, text="关闭", command=window.quit)
    button.pack(pady=10)
    
    # 运行窗口
    window.mainloop()

# 作者信息
def show_personal_info_abnormal():
    info = '''
    作者: 卷包一乙赖灿兴
    版本: 1.0
    日期: 2025年2月20日
    电子邮件: laicx@stu.xmu.edu.cn
    联系方式: 18046307080
    代码用途: 物流系统与NC系统的辅料出入库对账
    描述: 该程序用于物流系统与NC系统的辅料出入库对账，并给出疑似异常数据和自动判断是否真实存在异常数据。
    
    本次对账存在疑似异常数据，请打开输出的对账表确认~ 
    
    如有问题和需要变更需求请及时与我取得联系QAQ~ 
    
    点关闭后程序执行完毕...
    请于'./NC_Logistics_Reconciliation/_internal'中提取输出的表格
    '''
    
    # 创建一个 Tkinter 窗口
    window = tk.Tk()
    window.title("个人信息")
    
    # 创建一个 Label 显示个人信息
    label = tk.Label(window, text=info, padx=10, pady=10)
    label.pack()
    
    # 添加一个按钮来关闭窗口
    button = tk.Button(window, text="关闭", command=window.quit)
    button.pack(pady=10)
    
    # 运行窗口
    window.mainloop()

# 获取当前脚本所在的文件夹路径
folder_path = os.path.dirname(os.path.abspath(__file__))

# ------------------- 合并高架库中间表 -------------------

# 获取文件夹中所有符合条件的文件（以“高架库中间表”开头的文件）
files = [f for f in os.listdir(folder_path) if '高架库中间表' in f and f.endswith('.xlsx')]

# 记录每个文件的数据量
data_info = {}

# 创建一个空的 DataFrame 用于存储合并的数据
merged_data = pd.DataFrame()

# 遍历每个文件，进行处理
for i, file in enumerate(sorted(files, key=lambda x: int(x.split('高架库中间表')[1].split('.xlsx')[0]))):
    file_path = os.path.join(folder_path, file)
    
    # 读取数据时不加载表头（header=None）
    df = pd.read_excel(file_path, header=None)
    
    # 提取最后一行合计行中的数据量
    last_row = df.iloc[-1]
    total_count = int(str(last_row[0]).split('共')[1].strip('[]').strip())
    
    # 记录该表的数据量
    data_info[file] = total_count
    
    # 删除前两行和最后一行（合计行）
    df = df.drop([0, 1, df.index[-1]])
    
    # 如果不是第一个文件，删除表头（跳过第一行）
    if i > 0:
        df.columns = merged_data.columns  # 保持合并表头一致
    else:
        # 将第一个文件的表头设置为实际的列名
        df.columns = df.iloc[0]  # 将第一行作为列名

    # 将数据合并到一个大的 DataFrame 中
    merged_data = pd.concat([merged_data, df], ignore_index=True)

# 删除合并后表格的第一行
merged_data = merged_data.drop([0])

# 保留表头，并删除含有“单据日期”的行
merged_data = merged_data[~merged_data.iloc[:, 0].str.contains('单据日期', na=False)]

# 保存合并后的数据为新的 Excel 文件
output_file_merged = os.path.join(folder_path, '高架库中间表_总.xlsx')
merged_data.to_excel(output_file_merged, index=False)

# 输出统计信息
output = f"共检测{len(files)}个中间表，以下是各个表的数据量：\n"
for file in sorted(files, key=lambda x: int(x.split('高架库中间表')[1].split('.xlsx')[0])):
    output += f"{file}：{data_info[file]}\n"
output += f"合并的高架库中间表：{sum(data_info.values())}"

# 输出结果
print(output)
# 输出合并完成的提示信息
print(f"各高架库中间表已合并，结果已保存为 {output_file_merged}")

# ------------------- 处理一区物流领退料记录 -------------------

# 输入文件路径
file_1 = os.path.join(folder_path, '一区物流领退料记录.xlsx')
file_2 = output_file_merged  # 高架库中间表_总.xlsx

# 读取数据，指定第二行作为表头
df = pd.read_excel(file_1, header=1)
df_high_rack = pd.read_excel(file_2)

# 处理列名：去除两端空格和可能的特殊字符
df.columns = df.columns.str.strip().str.replace(r'\s+', '', regex=True)

# 1. 删除第一行（由于已经指定header=1，这一行应已经被忽略）
df = df.drop(index=0)

# 2. 删除“物料编号”为空值的行，并记录删除的条数
initial_length = len(df)
df = df.dropna(subset=['物料编号'])
deleted_rows = initial_length - len(df)

# 3. 填充单据号列的空值，使其等于上个数据
df['单据号'] = df['单据号'].ffill()

# 4. 判断“物料编号”的前四位是“0301”时修改“单据号”
df.loc[df['物料编号'].astype(str).str[:4] == '0301', '单据号'] = df['单据号'] + 'ss'

# 5. 在“单据号”右侧新建“NC单据编号”列，并根据“高架库中间表_总.xlsx”的“单据编号”进行匹配
df['NC单据编号'] = df['单据号'].apply(lambda x: x if x in df_high_rack['单据编号'].values else '异常')

# 获取“单据号”列的位置
single_code_index = df.columns.get_loc('单据号')

# 将“NC单据编号”插入到“单据号”右侧
df.insert(single_code_index + 1, 'NC单据编号', df.pop('NC单据编号'))

# 筛选出疑似异常的数据，并确保是副本
suspected_df = df[(df['类型'].isin(['02', '313'])) & (df['NC单据编号'] == '异常')].copy()

# 在疑似异常数据中插入“是否确认是异常数据”列，初始值为NaN
suspected_df.loc[:, '是否确认是异常数据'] = None  # 使用.loc避免警告

# 为“是否确认是异常数据”列赋值
for index, row in suspected_df.iterrows():
    if row['状态'] in ['新增', '取消'] or pd.isna(row['统计数量']) or row['统计数量'] <= 0:
        suspected_df.at[index, '是否确认是异常数据'] = '否'
    else:
        suspected_df.at[index, '是否确认是异常数据'] = '是'

# 在“搬运类型”列的右边插入“是否确认是异常数据”列，并填充黄色
搬运类型_index = suspected_df.columns.get_loc('搬运类型')
suspected_df.insert(搬运类型_index + 1, '是否确认是异常数据', suspected_df.pop('是否确认是异常数据'))

# 创建一个ExcelWriter对象
output_file_1 = os.path.join(folder_path, '一区物流领退料记录对账表.xlsx')
with pd.ExcelWriter(output_file_1, engine='xlsxwriter') as writer:
    # 将主表保存
    df.to_excel(writer, sheet_name='主表', index=False)

    # 将疑似异常数据表保存
    suspected_df.to_excel(writer, sheet_name='疑似异常的数据', index=False)

    # 获取工作簿和工作表
    workbook = writer.book
    worksheet = writer.sheets['疑似异常的数据']

    # 获取“是否确认是异常数据”列的索引位置
    column_index = suspected_df.columns.get_loc('是否确认是异常数据')

    # 填充黄色背景
    yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
    worksheet.set_column(column_index, column_index, None, yellow_format)

# 判断是否有“是否确认是异常数据”为“是”的情况
if (suspected_df['是否确认是异常数据'] == '是').any():
    is_normal_1 = False
    print("本次一区对账存在疑似异常数据，请打开输出的对账表查看")
    # show_personal_info_abnormal()
else:
    # show_personal_info_normal()
    print("本次一区对账没有异常数据，您可以打开输出的对账表查看")

print(f"处理完成，已保存至 {output_file_1}")

# ------------------- 处理二区物流领退料记录 -------------------

# 输入文件路径
file_2 = os.path.join(folder_path, '二区物流领退料记录.xlsx')
file_3 = output_file_merged  # 高架库中间表_总.xlsx

# 读取数据，指定第二行作为表头
df = pd.read_excel(file_2, header=1)
df_high_rack = pd.read_excel(file_3)

# 处理列名：去除两端空格和可能的特殊字符
df.columns = df.columns.str.strip().str.replace(r'\s+', '', regex=True)

# 1. 删除第一行（由于已经指定header=1，这一行应已经被忽略）
df = df.drop(index=0)

# 2. 删除第一列
df = df.drop(df.columns[0], axis=1)

# 3. 将“流水ID”列和“高架库中间表_总.xlsx”的“单据编号”列转换为字符串类型，确保格式一致
df['流水ID'] = df['流水ID'].astype(str)
df_high_rack['单据编号'] = df_high_rack['单据编号'].astype(str)

# 4. 在“流水ID”右侧新建“NC单据编号”列，根据规则填充
df['NC单据编号'] = df['流水ID'].apply(lambda x: x if x in df_high_rack['单据编号'].values else '异常')

# 获取“流水ID”列的位置
流水ID_index = df.columns.get_loc('流水ID')

# 将“NC单据编号”插入到“流水ID”右侧
df.insert(流水ID_index + 1, 'NC单据编号', df.pop('NC单据编号'))

# 5. 创建“疑似异常的数据”子表，筛选出“NC单据编号”=“异常”的数据
suspected_df = df[df['NC单据编号'] == '异常'].copy()

# 6. 在子表中插入“是否确认是异常数据”列，初始值为NaN
suspected_df.loc[:, '是否确认是异常数据'] = None

# 7. 根据规则为“是否确认是异常数据”列赋值
for index, row in suspected_df.iterrows():
    if row['类型'] in ['车间暂存区发料到机台', '机台人工退车间暂存区', '机台自动叫料（平衡区）', '交班', '接班']:
        suspected_df.at[index, '是否确认是异常数据'] = '否'
    else:
        suspected_df.at[index, '是否确认是异常数据'] = '是'

# 8. 在“NC单据编号”右边插入“是否确认是异常数据”列，并填充黄色
NC单据编号_index = suspected_df.columns.get_loc('NC单据编号')
suspected_df.insert(NC单据编号_index + 1, '是否确认是异常数据', suspected_df.pop('是否确认是异常数据'))

# 创建一个ExcelWriter对象
output_file_2 = os.path.join(folder_path, '二区物流领退料记录对账表.xlsx')
with pd.ExcelWriter(output_file_2, engine='xlsxwriter') as writer:
    # 将主表保存
    df.to_excel(writer, sheet_name='主表', index=False)

    # 将疑似异常数据表保存
    suspected_df.to_excel(writer, sheet_name='疑似异常的数据', index=False)

    # 获取工作簿和工作表
    workbook = writer.book
    worksheet = writer.sheets['疑似异常的数据']

    # 获取“是否确认是异常数据”列的索引位置
    column_index = suspected_df.columns.get_loc('是否确认是异常数据')

    # 填充黄色背景
    yellow_format = workbook.add_format({'bg_color': '#FFFF00'})
    worksheet.set_column(column_index, column_index, None, yellow_format)

# 判断是否有“是否确认是异常数据”为“是”的情况
if (suspected_df['是否确认是异常数据'] == '是').any():
    is_normal_2 = False
    print("本次二区对账存在疑似异常数据，请打开输出的对账表查看")
    # show_personal_info_abnormal()
else:
    # show_personal_info_normal()
    print("本次二区对账没有异常数据，您可以打开输出的对账表查看")

print(f"处理完成，已保存至 {output_file_2}")

# 弹出窗口展示结果
if is_normal_1 and is_normal_2:
    show_personal_info_normal()
else:
    show_personal_info_abnormal()


