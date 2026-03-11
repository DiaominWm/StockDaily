import os
import sys
import shutil

def remove_all_trailing_spaces(file_path, backup_suffix='.bak'):
    """
    彻底删除文件中所有行尾的空格/制表符（包括空行中的),兼容Windows(\r\n)和Linux(\n)换行
    """
    # 1. 备份原文件
    if not os.path.isfile(file_path):
        print(f"❌ 错误：{file_path} 不是有效文件！")
        return False

    backup_path = file_path + backup_suffix
    shutil.copy2(file_path, backup_path)
    print(f"✅ 已备份原文件到：{backup_path}")

    # 2. 读取文件（二进制模式避免编码问题，统一处理换行）
    with open(file_path, 'rb') as f:
        content = f.read().decode('utf-8', errors='ignore')

    # 3. 核心处理：拆分所有行 → 清理每行尾部空格 → 重新合并
    # 兼容 \r\n（Windows）和 \n（Linux/Mac）换行
    lines = content.splitlines()  # 自动识别所有换行符，拆分后无换行符
    cleaned_lines = []
    for line in lines:
        # 移除行尾所有空格/制表符（无论多少个）
        cleaned_line = line.rstrip(' \t')
        cleaned_lines.append(cleaned_line)

    # 4. 写回文件（用\n统一换行，也可改为\r\n适配Windows）
    with open(file_path, 'w', encoding='utf-8', newline='\n') as f:
        # 合并时添加换行符，空行会变成纯\n（无任何空格）
        f.write('\n'.join(cleaned_lines))

    print(f"✅ 处理完成！{file_path} 中所有行尾空格已删除")
    return True

def main():
    if len(sys.argv) != 2:
        print("📚 使用说明：")
        print("  python clean_spaces.py 你的代码文件.py")
        sys.exit(1)

    target_file = sys.argv[1]
    remove_all_trailing_spaces(target_file)

if __name__ == '__main__':
    main()
