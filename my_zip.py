import os
import zipfile

def extract_zip_skip_transpdf(zip_path, extract_to_dir):
    """
    解压zip文件，但跳过名为 'trans.pdf' 的文件。
    
    :param zip_path: 要解压的zip文件路径
    :param extract_to_dir: 解压目标目录
    """
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            # 获取压缩包内所有文件信息
            for file_info in zip_ref.infolist():
                # 如果文件是trans.pdf，则跳过
                if os.path.basename(file_info.filename) == 'trans.pdf':
                    print(f"跳过文件: {file_info.filename}")
                    continue
                # 解压非trans.pdf的文件
                zip_ref.extract(file_info, extract_to_dir)
        print(f"已解压（跳过trans.pdf）: {zip_path} -> {extract_to_dir}")
    except zipfile.BadZipFile:
        print(f"错误：文件不是有效的ZIP文件或已损坏 - {zip_path}")
    except Exception as e:
        print(f"解压 {zip_path} 时发生错误: {e}")

def process_inner_zips(main_folder_path):
    """
    处理主文件夹（即外层zip解压得到的同名文件夹）内的所有zip文件。
    解压这些内部zip文件中除'trans.pdf'之外的所有文件到主文件夹内。
    
    :param main_folder_path: 外层zip解压得到的同名文件夹路径
    """
    # 遍历主文件夹下的所有项
    for item in os.listdir(main_folder_path):
        item_full_path = os.path.join(main_folder_path, item)
        # 检查是否是文件且以.zip结尾
        if os.path.isfile(item_full_path) and item.lower().endswith('.zip'):
            print(f"处理内部压缩文件: {item_full_path}")
            # 解压这个内部zip文件（跳过trans.pdf）到主文件夹
            extract_zip_skip_transpdf(item_full_path, main_folder_path)

def main(target_directory):
    """
    主处理函数：遍历目标目录的子目录，处理每个子目录中的zip文件。
    
    :param target_directory: 需要处理的目录路径
    """
    # 检查目标目录是否存在
    if not os.path.exists(target_directory):
        print(f"错误：目录 '{target_directory}' 不存在。")
        return

    print(f"开始处理目录: {target_directory}")
    
    # 遍历目标目录下的所有子目录
    for entry in os.listdir(target_directory):
        entry_path = os.path.join(target_directory, entry)
        
        # 只处理子目录，跳过文件
        if not os.path.isdir(entry_path):
            continue
            
        print(f"处理子目录: {entry}")
        
        # 遍历子目录下的所有项
        for file_entry in os.listdir(entry_path):
            file_entry_path = os.path.join(entry_path, file_entry)
            
            # 检查是否是文件且是zip文件
            if os.path.isfile(file_entry_path) and file_entry.lower().endswith('.zip'):
                # 为zip文件创建同名文件夹路径（去掉.zip扩展名）
                folder_name = file_entry[:-4]  # 移除 .zip 后缀
                folder_path = os.path.join(entry_path, folder_name)

                # 如果同名文件夹已存在，提示并跳过（避免覆盖）
                if os.path.exists(folder_path):
                    print(f"注意：文件夹已存在，跳过处理 '{file_entry}' -> '{folder_name}'")
                    continue

                # 创建同名文件夹
                os.makedirs(folder_path)
                print(f"创建文件夹: {folder_path}")

                # 首先，将外层zip文件解压到刚创建的同名文件夹
                try:
                    with zipfile.ZipFile(file_entry_path, 'r') as outer_zip:
                        outer_zip.extractall(folder_path)
                    print(f"已解压外层压缩包: {file_entry} -> {folder_name}")
                except Exception as e:
                    print(f"解压外层压缩包 {file_entry} 时出错: {e}")
                    continue  # 如果外层解压失败，跳过该zip的后续处理

                # 然后，处理这个刚解压出来的同名文件夹里的内部zip文件
                process_inner_zips(folder_path)

    print("处理完成。")

# 使用示例
if __name__ == "__main__":
    # 替换为你的目标目录路径
    path_to_scan = input("请输入要处理的目录路径（留空则为当前目录）: ").strip()
    if not path_to_scan:
        path_to_scan = os.getcwd()
    main(path_to_scan)