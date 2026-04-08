import os
import shutil
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def select_excel_file():
    """选择Excel文件"""
    file_path = filedialog.askopenfilename(
        title="选择包含音频信息的Excel文件",
        filetypes=[("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*")]
    )
    return file_path

def select_target_folder():
    """选择目标文件夹"""
    folder_path = filedialog.askdirectory(title="选择目标文件夹")
    return folder_path

def process_and_move_files():
    """处理Excel文件并移动音频和LRC文件"""
    # 1. 选择Excel文件
    excel_path = select_excel_file()
    if not excel_path:
        return
    
    # 2. 选择目标文件夹
    target_folder = select_target_folder()
    if not target_folder:
        return
    
    # 3. 创建目标文件夹（如果不存在）
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    
    try:
        # 4. 读取Excel文件
        df = pd.read_excel(excel_path)
        
        # 5. 获取"完整路径"列
        if "完整路径" not in df.columns:
            # 尝试其他可能的列名
            possible_columns = ["完整路径", "文件路径", "path", "file_path", "Path"]
            path_column = None
            for col in possible_columns:
                if col in df.columns:
                    path_column = col
                    break
            
            if not path_column:
                raise ValueError("Excel文件中未找到文件路径列")
        else:
            path_column = "完整路径"
        
        # 6. 处理每一行
        moved_count = 0
        audio_extensions = {'.mp3', '.flac', '.m4a', '.wav', '.aac', '.ogg'}
        
        for index, row in df.iterrows():
            try:
                audio_path = str(row[path_column])
                
                # 检查路径是否存在
                if not os.path.exists(audio_path):
                    print(f"文件不存在: {audio_path}")
                    continue
                
                # 获取文件名和扩展名
                filename = os.path.basename(audio_path)
                name_without_ext, ext = os.path.splitext(filename)
                ext = ext.lower()
                
                # 检查是否是音频文件
                if ext in audio_extensions:
                    # 移动音频文件
                    target_audio_path = os.path.join(target_folder, filename)
                    shutil.move(audio_path, target_audio_path)
                    print(f"已移动音频文件: {filename}")
                    
                    # 查找并移动同名的LRC文件
                    audio_dir = os.path.dirname(audio_path)
                    lrc_filename = f"{name_without_ext}.lrc"
                    lrc_path = os.path.join(audio_dir, lrc_filename)
                    
                    if os.path.exists(lrc_path):
                        target_lrc_path = os.path.join(target_folder, lrc_filename)
                        shutil.move(lrc_path, target_lrc_path)
                        print(f"已移动LRC文件: {lrc_filename}")
                    
                    moved_count += 1
                else:
                    print(f"非音频文件跳过: {filename}")
                    
            except Exception as e:
                print(f"处理文件时出错: {e}")
                continue
        
        # 7. 显示结果
        if moved_count > 0:
            messagebox.showinfo("完成", f"成功移动 {moved_count} 个音频文件及其LRC文件到:\n{target_folder}")
        else:
            messagebox.showwarning("无文件移动", "未找到可移动的音频文件")
    
    except Exception as e:
        messagebox.showerror("错误", f"处理Excel文件时出错:\n{e}")
        print(f"详细错误: {e}")

def main():
    """主函数"""
    # 创建简单的GUI
    root = tk.Tk()
    root.title("音频文件整理工具")
    root.geometry("400x200")
    
    # 添加说明标签
    label = tk.Label(root, text="音频文件整理工具", font=("Arial", 16))
    label.pack(pady=20)
    
    instruction = tk.Label(root, 
                          text="1. 选择包含音频信息的Excel文件\n2. 选择目标文件夹\n3. 程序将自动移动音频文件及同名LRC文件",
                          justify="left")
    instruction.pack(pady=10)
    
    # 添加开始按钮
    start_btn = tk.Button(root, text="开始整理", command=process_and_move_files, 
                         font=("Arial", 12), bg="#4CAF50", fg="white", 
                         padx=20, pady=10)
    start_btn.pack(pady=20)
    
    root.mainloop()

if __name__ == "__main__":
    main()