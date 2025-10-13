import pandas as pd
import os
from tkinter import filedialog
from tkinter import Tk
def split_csv_by_secondary_organization(input_file_path, output_dir):
    """
    按二级组织拆分CSV考勤文件
    
    :param input_file_path: 输入的CSV文件路径
    :param output_dir: 拆分后文件的存储目录
    """
    # 读取CSV文件，添加 encoding='gbk' 以处理中文编码问题
    df = pd.read_csv(input_file_path, encoding='gbk')
    
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 从所属组织列中提取二级组织
    df['二级组织'] = df['所属组织'].str.split('/').str[1]
    
    # 按二级组织分组
    grouped = df.groupby('二级组织')
    
    # 为每个二级组织保存单独的文件
    for org_name, group in grouped:
        # 移除临时添加的二级组织列
        group = group.drop(columns=['二级组织'])
        output_file_path = os.path.join(output_dir, f'{org_name}_考勤记录.csv')
        group.to_csv(output_file_path, index=False, encoding='utf-8-sig')

if __name__ == "__main__":
    # 创建一个隐藏的Tkinter窗口
    root = Tk()
    root.withdraw()
    
    # 使用文件对话框选择输入文件
    input_file = filedialog.askopenfilename(
        title="请选择打卡CSV文件",
        filetypes=[("CSV文件", "*.csv")]
    )
    
    # 使用目录对话框选择输出目录
    output_directory = filedialog.askdirectory(
        title="请选择文件存储位置"
    )
    
    if input_file and output_directory:
        split_csv_by_secondary_organization(input_file, output_directory)
