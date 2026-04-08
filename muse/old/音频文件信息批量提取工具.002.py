import os
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD  # 需要安装tkinterdnd2: pip install tkinterdnd2
import pandas as pd
from mutagen.easyid3 import EasyID3
from mutagen.id3 import ID3, error as ID3error
from mutagen.flac import FLAC
from mutagen.mp3 import MP3
from mutagen.mp4 import MP4
import traceback
from datetime import timedelta

class AudioInfoExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("音频文件信息批量提取器")
        self.root.geometry("900x650")
        
        # 存储所有文件信息的列表
        self.all_file_info = []
        
        self.setup_ui()
        
    def setup_ui(self):
        # 顶部框架：标题和说明
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        title_label = ttk.Label(top_frame, text="音频文件信息批量提取工具", font=("Arial", 16, "bold"))
        title_label.pack()
        
        instruction_label = ttk.Label(top_frame, 
                                     text="将包含音频文件（MP3/FLAC）的文件夹拖放到下方区域，或使用“选择文件夹”按钮。\n程序将自动读取所有子文件夹中的音频文件，并提取详细信息。",
                                     wraplength=800,
                                     justify="center")
        instruction_label.pack(pady=5)
        
        # 控制按钮框架
        button_frame = ttk.Frame(self.root, padding="5")
        button_frame.pack(fill=tk.X)
        
        self.select_btn = ttk.Button(button_frame, text="选择文件夹", command=self.select_folder)
        self.select_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_btn = ttk.Button(button_frame, text="清空列表", command=self.clear_list)
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_btn = ttk.Button(button_frame, text="导出为Excel", command=self.export_to_excel, state=tk.DISABLED)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        # 拖放区域
        drop_frame = ttk.LabelFrame(self.root, text="拖放区域", padding="20")
        drop_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 使用 tk.Label 替代 ttk.Label，因为 ttk.Label 不支持 height 参数
        self.drop_label = tk.Label(drop_frame, 
                                   text="拖放文件夹至此",
                                   font=("Arial", 12),
                                   relief="sunken",
                                   anchor="center",
                                   width=50,
                                   height=8)
        self.drop_label.pack(fill=tk.BOTH, expand=True)
        
        # 注册拖放事件
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.on_drop)
        
        # 信息显示区域
        info_frame = ttk.LabelFrame(self.root, text="文件信息列表", padding="10")
        info_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        # 创建树形视图
        columns = ("序号", "文件名", "路径", "标题", "艺术家", "专辑", "时长", "比特率", "采样率", "文件大小")
        self.tree = ttk.Treeview(info_frame, columns=columns, show="headings", height=12)
        
        # 设置列标题和宽度
        col_widths = [50, 120, 200, 120, 100, 120, 80, 80, 80, 80]
        for idx, col in enumerate(columns):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths[idx])
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(info_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def on_drop(self, event):
        """处理拖放事件"""
        # 获取拖放的文件/文件夹路径
        path = event.data.strip('{}')
        if os.path.isdir(path):
            self.process_folder(path)
        else:
            messagebox.showwarning("警告", "请拖放文件夹，而不是单个文件。")
    
    def select_folder(self):
        """选择文件夹"""
        folder_path = filedialog.askdirectory(title="选择包含音频文件的文件夹")
        if folder_path:
            self.process_folder(folder_path)
    
    def process_folder(self, folder_path):
        """处理文件夹，递归查找音频文件"""
        self.status_var.set(f"正在扫描文件夹: {folder_path}")
        self.root.update()
        
        # 清空现有列表
        self.clear_list()
        
        # 支持的音频文件扩展名
        audio_extensions = {'.mp3', '.flac', '.m4a'}
        file_count = 0
        
        # 递归遍历文件夹
        for root_dir, dirs, files in os.walk(folder_path):
            for file in files:
                if os.path.splitext(file)[1].lower() in audio_extensions:
                    file_path = os.path.join(root_dir, file)
                    self.process_audio_file(file_path, file_count + 1)
                    file_count += 1
        
        if file_count > 0:
            self.status_var.set(f"完成！共找到 {file_count} 个音频文件。")
            self.export_btn.config(state=tk.NORMAL)
        else:
            self.status_var.set("未在指定文件夹中找到MP3或FLAC文件。")
            messagebox.showinfo("提示", "未在指定文件夹中找到MP3或FLAC文件。")
    
    def process_audio_file(self, file_path, index):
        """处理单个音频文件，提取信息"""
        try:
            file_name = os.path.basename(file_path)
            file_ext = os.path.splitext(file_name)[1].lower()
            
            # 初始化信息字典
            info = {
                "文件名": file_name,
                "完整路径": file_path,
                "标题": "",
                "艺术家": "",
                "专辑": "",
                "专辑艺术家": "",
                "作曲者": "",
                "年代": "",
                "曲目编号": "",
                "光盘编号": "",
                "流派": "",
                "时长": "",
                "比特率": "",
                "采样率": "",
                "位深": "",
                "文件大小": "",
                "编码类型": "",
                "文件格式": file_ext[1:].upper() if file_ext else "",
            }
            
            # 获取文件大小
            file_size = os.path.getsize(file_path)
            info["文件大小"] = f"{file_size / 1024 / 1024:.2f} MB"
            
            # 根据文件类型使用不同的方法读取标签
            if file_ext == '.mp3':
                self.read_mp3_info(file_path, info)
            elif file_ext == '.flac':
                self.read_flac_info(file_path, info)
            elif file_ext == '.m4a':
                self.read_m4a_info(file_path, info)
            
            # 将信息添加到列表
            self.all_file_info.append(info)
            
            # 添加到树形视图
            self.tree.insert("", tk.END, values=(
                index,
                info["文件名"],
                os.path.dirname(info["完整路径"]),
                info["标题"],
                info["艺术家"],
                info["专辑"],
                info["时长"],
                info["比特率"],
                info["采样率"],
                info["文件大小"]
            ))
            
        except Exception as e:
            print(f"处理文件 {file_path} 时出错: {e}")
            traceback.print_exc()
    
    def read_mp3_info(self, file_path, info):
        """读取MP3文件信息"""
        try:
            # 使用EasyID3读取标准ID3标签
            audio = EasyID3(file_path)
            for key in audio:
                if key in info:
                    info[key] = ", ".join(audio[key]) if audio[key] else ""
        except (ID3error, Exception):
            # 如果EasyID3失败，尝试使用ID3
            try:
                audio_tags = ID3(file_path)
                # 手动映射一些常见标签
                tag_mapping = {
                    'TIT2': '标题',
                    'TPE1': '艺术家',
                    'TALB': '专辑',
                    'TPE2': '专辑艺术家',
                    'TCOM': '作曲者',
                    'TYER': '年代',
                    'TRCK': '曲目编号',
                    'TPOS': '光盘编号',
                    'TCON': '流派',
                }
                
                for tag_key, info_key in tag_mapping.items():
                    if tag_key in audio_tags:
                        info[info_key] = str(audio_tags[tag_key])
            except:
                pass
        
        # 使用mutagen获取音频属性
        try:
            audio_props = MP3(file_path)
            if audio_props.info.length:
                seconds = int(audio_props.info.length)
                info["时长"] = str(timedelta(seconds=seconds))
            
            if audio_props.info.bitrate:
                info["比特率"] = f"{audio_props.info.bitrate / 1000:.1f} kbps"
            
            if audio_props.info.sample_rate:
                info["采样率"] = f"{audio_props.info.sample_rate / 1000:.1f} kHz"
        except:
            pass
        
        info["编码类型"] = "MPEG Audio Layer III"
    
    def read_flac_info(self, file_path, info):
        """读取FLAC文件信息"""
        try:
            audio = FLAC(file_path)
            
            # FLAC使用Vorbis注释，标签名可能小写
            tag_mapping = {
                'title': '标题',
                'artist': '艺术家',
                'album': '专辑',
                'albumartist': '专辑艺术家',
                'composer': '作曲者',
                'date': '年代',
                'tracknumber': '曲目编号',
                'discnumber': '光盘编号',
                'genre': '流派',
            }
            
            for tag_key, info_key in tag_mapping.items():
                if tag_key in audio:
                    info[info_key] = ", ".join(audio[tag_key]) if audio[tag_key] else ""
            
            # 音频属性
            if audio.info.length:
                seconds = int(audio.info.length)
                info["时长"] = str(timedelta(seconds=seconds))
            
            if audio.info.bitrate:
                info["比特率"] = f"{audio.info.bitrate / 1000:.1f} kbps"
            
            if audio.info.sample_rate:
                info["采样率"] = f"{audio.info.sample_rate / 1000:.1f} kHz"
            
            info["位深"] = f"{audio.info.bits_per_sample} bit" if audio.info.bits_per_sample else ""
            info["编码类型"] = "Free Lossless Audio Codec"
            
        except Exception as e:
            print(f"读取FLAC文件 {file_path} 信息失败: {e}")
    
    def read_m4a_info(self, file_path, info):
        """读取M4A文件信息"""
        try:
            audio = MP4(file_path)
            
            # M4A标签映射
            tag_mapping = {
                '\xa9nam': '标题',
                '\xa9ART': '艺术家',
                '\xa9alb': '专辑',
                'aART': '专辑艺术家',
                '\xa9wrt': '作曲者',
                '\xa9day': '年代',
                'trkn': '曲目编号',
                'disk': '光盘编号',
                '\xa9gen': '流派',
            }
            
            for tag_key, info_key in tag_mapping.items():
                if tag_key in audio:
                    if tag_key in ['trkn', 'disk'] and audio[tag_key]:
                        # 处理曲目编号和光盘编号（格式为[(当前, 总数)]）
                        info[info_key] = f"{audio[tag_key][0][0]}/{audio[tag_key][0][1]}" if len(audio[tag_key][0]) > 1 else str(audio[tag_key][0][0])
                    else:
                        info[info_key] = ", ".join(audio[tag_key]) if audio[tag_key] else ""
            
            # 音频属性
            if audio.info.length:
                seconds = int(audio.info.length)
                info["时长"] = str(timedelta(seconds=seconds))
            
            if audio.info.bitrate:
                info["比特率"] = f"{audio.info.bitrate / 1000:.1f} kbps"
            
            if audio.info.sample_rate:
                info["采样率"] = f"{audio.info.sample_rate / 1000:.1f} kHz"
            
            info["编码类型"] = "MPEG-4 Audio"
            
        except Exception as e:
            print(f"读取M4A文件 {file_path} 信息失败: {e}")
    
    def clear_list(self):
        """清空文件列表"""
        self.all_file_info.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.export_btn.config(state=tk.DISABLED)
        self.status_var.set("就绪")
    
    def export_to_excel(self):
        """导出为Excel文件"""
        if not self.all_file_info:
            messagebox.showwarning("警告", "没有可导出的数据。")
            return
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            title="保存Excel文件"
        )
        
        if save_path:
            try:
                # 创建DataFrame
                df = pd.DataFrame(self.all_file_info)
                
                # 重新排列列顺序，使重要信息靠前
                column_order = [
                    "文件名", "完整路径", "标题", "艺术家", "专辑", "专辑艺术家", 
                    "作曲者", "年代", "曲目编号", "光盘编号", "流派", 
                    "时长", "比特率", "采样率", "位深", "文件大小", 
                    "编码类型", "文件格式"
                ]
                
                # 只保留DataFrame中实际存在的列
                existing_columns = [col for col in column_order if col in df.columns]
                df = df[existing_columns]
                
                # 导出到Excel
                df.to_excel(save_path, index=False)
                self.status_var.set(f"数据已成功导出到: {save_path}")
                messagebox.showinfo("成功", f"数据已成功导出到:\n{save_path}")
                
            except Exception as e:
                messagebox.showerror("错误", f"导出文件时出错:\n{e}")
                self.status_var.set("导出失败")

def main():
    # 创建支持拖放的Tkinter窗口
    root = TkinterDnD.Tk()
    app = AudioInfoExtractorApp(root)
    root.mainloop()

if __name__ == "__main__":
    # 检查必要库是否已安装
    try:
        import pandas
        import mutagen
        from tkinterdnd2 import TkinterDnD
    except ImportError as e:
        print("缺少必要的库，请使用以下命令安装：")
        print("pip install pandas mutagen tkinterdnd2 openpyxl")
        sys.exit(1)
    
    main()
