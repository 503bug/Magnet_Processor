# magnet_processor.py
import requests
import openpyxl
import xlsxwriter
from openpyxl import load_workbook
from PIL import Image
import io
import tempfile
import os
import time
import sys
import glob
import json
import hashlib
from datetime import datetime
from pathlib import Path

# ========== 配置参数（可通过环境变量覆盖） ==========
REQUEST_INTERVAL = float(os.getenv('REQUEST_INTERVAL', 1))        # API 请求间隔（秒）
BATCH_SIZE = int(os.getenv('BATCH_SIZE', 2))                      # 每批处理行数
MAX_RETRIES = int(os.getenv('MAX_RETRIES', 5))                    # 下载图片重试次数
TIMEOUT = int(os.getenv('TIMEOUT', 20))                           # 请求超时（秒）
DEFAULT_COL_WIDTH = int(os.getenv('DEFAULT_COL_WIDTH', 80))       # 截图列宽（字符数）
DEFAULT_ROW_HEIGHT = int(os.getenv('DEFAULT_ROW_HEIGHT', 274))    # 截图行高（磅）
JPEG_QUALITY = int(os.getenv('JPEG_QUALITY', 85))                 # JPEG 压缩质量
MAX_RUN_TIME = int(os.getenv('MAX_RUN_TIME', 50))                 # 单次最大运行时间（分钟）
MAX_ROWS_PER_RUN = int(os.getenv('MAX_ROWS_PER_RUN', 200))        # 单次运行最大处理行数
# ====================================================

# 尝试导入 PIL
try:
    from PIL import Image as PILImage
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False
    print("错误: 未安装 Pillow，请安装: pip install Pillow")
    sys.exit(1)


def format_size(size_bytes):
    """字节转可读格式"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 ** 2:
        return f"{size_bytes / 1024:.2f} KB"
    elif size_bytes < 1024 ** 3:
        return f"{size_bytes / 1024 ** 2:.2f} MB"
    elif size_bytes < 1024 ** 4:
        return f"{size_bytes / 1024 ** 3:.2f} GB"
    else:
        return f"{size_bytes / 1024 ** 4:.2f} TB"


def get_magnet_info(magnet_link):
    """调用 whatslink.info API 获取磁力信息"""
    url = "https://whatslink.info/api/v1/link"
    params = {"url": magnet_link}
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=TIMEOUT)
        resp.raise_for_status()
        data = resp.json()
        if data.get("error"):
            print(f"API返回错误: {data['error']}")
            return None

        name = data.get("name", "").strip()
        count = data.get("count", 0)
        size_bytes = data.get("size", 0)
        size_str = format_size(size_bytes)

        screenshots = []
        for item in data.get("screenshots", []):
            if isinstance(item, dict) and item.get("screenshot"):
                screenshots.append({
                    "time": item.get("time", 0),
                    "screenshot": item.get("screenshot", "")
                })

        return {
            "name": name,
            "count": count,
            "size": size_str,
            "screenshots": screenshots
        }
    except Exception as e:
        print(f"API请求失败: {e}")
        return None


def download_and_convert_to_jpeg(url, quality=JPEG_QUALITY, max_retries=MAX_RETRIES, timeout=TIMEOUT):
    """下载图片并转换为 JPEG"""
    for attempt in range(max_retries):
        try:
            resp = requests.get(url, timeout=timeout)
            resp.raise_for_status()
            img = PILImage.open(io.BytesIO(resp.content))
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            fd, tmp_path = tempfile.mkstemp(suffix='.jpg')
            os.close(fd)
            img.save(tmp_path, format='JPEG', quality=quality)
            return tmp_path
        except Exception as e:
            print(f"下载/转换图片失败 (尝试 {attempt+1}/{max_retries}): {url} - {e}")
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                return None
    return None


def get_file_hash(file_path):
    """获取文件的 MD5 哈希值"""
    hasher = hashlib.md5()
    with open(file_path, 'rb') as f:
        buf = f.read(65536)
        while len(buf) > 0:
            hasher.update(buf)
            buf = f.read(65536)
    return hasher.hexdigest()


class ProgressManager:
    """管理处理进度，支持断点续传"""
    
    def __init__(self, input_dir="Input", output_dir="Output", state_file=".progress.json"):
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.state_file = self.output_dir / state_file
        self.state = {}
        self.load_state()
    
    def load_state(self):
        """加载状态文件"""
        if self.state_file.exists():
            try:
                with open(self.state_file, 'r', encoding='utf-8') as f:
                    self.state = json.load(f)
                print(f"加载进度状态: 已处理 {self.get_total_processed()} 行")
            except Exception as e:
                print(f"加载状态文件失败: {e}")
                self.state = {}
        else:
            self.state = {}
    
    def save_state(self):
        """保存状态文件"""
        self.state['_updated_at'] = datetime.now().isoformat()
        try:
            with open(self.state_file, 'w', encoding='utf-8') as f:
                json.dump(self.state, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"保存状态文件失败: {e}")
    
    def get_file_state(self, file_path):
        """获取单个文件的处理状态"""
        file_key = str(file_path)
        if file_key not in self.state:
            self.state[file_key] = {
                'processed_rows': 0,      # 已处理的数据行数（不含标题行）
                'total_rows': 0,           # 总数据行数（不含标题行）
                'file_hash': '',
                'completed': False,
                'last_modified': '',
                'batches_created': []
            }
        return self.state[file_key]
    
    def update_file_state(self, file_path, processed_rows, total_rows=None, completed=False, batch_num=None):
        """更新文件处理状态"""
        state = self.get_file_state(file_path)
        state['processed_rows'] = processed_rows
        if total_rows is not None:
            state['total_rows'] = total_rows
        if completed:
            state['completed'] = True
        state['last_modified'] = datetime.now().isoformat()
        if batch_num is not None:
            if 'batches_created' not in state:
                state['batches_created'] = []
            if batch_num not in state['batches_created']:
                state['batches_created'].append(batch_num)
        self.save_state()
    
    def mark_file_processed(self, file_path):
        """标记文件已完全处理"""
        state = self.get_file_state(file_path)
        state['completed'] = True
        state['last_modified'] = datetime.now().isoformat()
        self.save_state()
    
    def is_file_completed(self, file_path):
        """检查文件是否已完全处理"""
        state = self.get_file_state(file_path)
        return state.get('completed', False)
    
    def get_processed_rows(self, file_path):
        """获取文件已处理的数据行数（不含标题行）"""
        state = self.get_file_state(file_path)
        return state.get('processed_rows', 0)
    
    def get_total_processed(self):
        """获取所有文件已处理的总行数"""
        total = 0
        for key, state in self.state.items():
            if key.startswith('_'):
                continue
            total += state.get('processed_rows', 0)
        return total
    
    def reset_file_state(self, file_path):
        """重置文件状态"""
        file_key = str(file_path)
        if file_key in self.state:
            del self.state[file_key]
            self.save_state()


def get_excel_files(input_dir):
    """从 Input 目录获取所有 Excel 文件"""
    if not input_dir.exists():
        print(f"Input 目录不存在: {input_dir}")
        return []
    
    excel_files = []
    extensions = ['*.xlsx', '*.xlsm', '*.xls']
    for ext in extensions:
        excel_files.extend(glob.glob(str(input_dir / ext)))
    
    return [Path(f) for f in excel_files]


def get_unprocessed_files(input_dir, progress_manager):
    """获取未处理的文件列表"""
    all_files = get_excel_files(input_dir)
    unprocessed = []
    
    for file_path in all_files:
        if progress_manager.is_file_completed(file_path):
            print(f"  跳过已完成文件: {file_path.name}")
            continue
        
        file_hash = get_file_hash(file_path)
        state = progress_manager.get_file_state(file_path)
        if state.get('file_hash') and state.get('file_hash') != file_hash:
            print(f"  文件 {file_path.name} 已变更，重置进度")
            progress_manager.reset_file_state(file_path)
        
        state['file_hash'] = file_hash
        unprocessed.append(file_path)
    
    unprocessed.sort(key=lambda f: progress_manager.get_processed_rows(f))
    return unprocessed


def count_data_rows_in_file(file_path):
    """
    统计 Excel 文件中的数据行数（不含标题行）
    第一行为标题行，从第二行开始计数
    """
    try:
        wb = load_workbook(file_path, read_only=True)
        ws = wb.active
        # 从第2行开始计数（跳过标题行）
        row_count = sum(1 for _ in ws.iter_rows(min_row=2))
        wb.close()
        return row_count
    except Exception as e:
        print(f"统计行数失败 {file_path.name}: {e}")
        return 0


def get_headers_from_file(file_path):
    """
    从 Excel 文件读取标题行
    如果第一行为空，使用默认标题
    """
    try:
        wb = load_workbook(file_path, read_only=True)
        ws = wb.active
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        wb.close()
        
        if first_row:
            # 取前4列作为标题
            headers = [str(h) if h else "" for h in first_row[:4]]
            # 如果标题为空，使用默认值
            if not any(headers):
                return ["链接", "Resource Name", "Number of Files", "Total File Size"]
            return headers
        else:
            return ["链接", "Resource Name", "Number of Files", "Total File Size"]
    except Exception as e:
        print(f"读取标题行失败 {file_path.name}: {e}")
        return ["链接", "Resource Name", "Number of Files", "Total File Size"]


def save_batch_to_file(batch_data, headers, batch_num, output_dir, base_name):
    """
    保存批次数据为 Excel 文件，包含标题行
    """
    output_filename = f"{batch_num:03d}_{base_name}.xlsx"
    output_path = output_dir / output_filename
    
    print(f"  保存批次 {batch_num} 到: {output_filename}")
    
    workbook = xlsxwriter.Workbook(str(output_path))
    worksheet = workbook.add_worksheet()
    
    # 设置列宽
    worksheet.set_column(0, 3, 20)  # 前4列
    
    # ===== 写入标题行（第1行） =====
    for col_idx, header in enumerate(headers):
        worksheet.write(0, col_idx, header)
    
    # ===== 写入数据行（从第2行开始） =====
    for row_idx, row_data in enumerate(batch_data, start=1):
        # row_idx 从1开始，对应Excel的第2行（因为第1行是标题）
        excel_row = row_idx + 1
        
        worksheet.write(excel_row, 0, row_data['magnet'])
        worksheet.write(excel_row, 1, row_data['name'])
        worksheet.write(excel_row, 2, row_data['count'])
        worksheet.write(excel_row, 3, row_data['size'])
        
        # 设置行高
        worksheet.set_row(excel_row, DEFAULT_ROW_HEIGHT)
        
        # 插入截图（从第5列开始）
        if row_data['temp_files']:
            col = 4  # 第5列（E列）
            for tmp_path in row_data['temp_files']:
                if os.path.exists(tmp_path):
                    worksheet.set_column(col, col, DEFAULT_COL_WIDTH)
                    worksheet.insert_image(excel_row, col, tmp_path, {'x_scale': 1, 'y_scale': 1})
                    col += 1
    
    workbook.close()
    print(f"  批次 {batch_num} 保存完成")


def process_single_file(file_path, output_dir, progress_manager):
    """
    处理单个 Excel 文件，支持断点续传
    第一行为标题行，从第二行开始处理数据
    """
    print(f"\n处理文件: {file_path.name}")
    
    # 获取已处理的数据行数（不含标题行）
    start_row = progress_manager.get_processed_rows(file_path)
    if start_row > 0:
        print(f"  从第 {start_row + 2} 行继续（已处理 {start_row} 行数据）")
    
    try:
        # 打开源文件
        wb_in = load_workbook(file_path, read_only=True)
        ws_in = wb_in.active
        
        # ===== 读取标题行 =====
        headers = get_headers_from_file(file_path)
        print(f"  标题行: {headers}")
        
        # ===== 统计总数据行数（不含标题行） =====
        total_rows = count_data_rows_in_file(file_path)
        print(f"  总数据行数: {total_rows}")
        
        # 如果已经全部处理完成
        if start_row >= total_rows:
            progress_manager.mark_file_processed(file_path)
            print(f"  ✓ 文件 {file_path.name} 已全部处理完成！")
            wb_in.close()
            return 0, False
        
        # 初始化批次
        batch_num = 1
        if start_row > 0:
            batch_num = (start_row // BATCH_SIZE) + 1
        
        # Excel 行号：从第2行开始（第1行是标题）
        # 已处理 start_row 行，所以当前从 start_row + 2 行开始
        current_excel_row = start_row + 2
        
        processed_in_this_run = 0
        max_rows_to_process = min(MAX_ROWS_PER_RUN, total_rows - start_row)
        batch_data = []
        temp_files = []
        start_time = time.time()
        
        print(f"  本次计划处理: {max_rows_to_process} 行数据")
        
        # ===== 跳过已处理的行 =====
        # 从第2行开始，跳过 start_row 行
        row_counter = 0
        for row in ws_in.iter_rows(min_row=2, values_only=True):
            if row_counter >= start_row:
                break
            row_counter += 1
            current_excel_row += 1
        
        # ===== 处理新数据 =====
        for row in ws_in.iter_rows(min_row=current_excel_row, values_only=True):
            # 检查是否达到限制
            if processed_in_this_run >= max_rows_to_process:
                print(f"  已达到本次运行行数限制: {max_rows_to_process}")
                break
            
            # 检查运行时间
            elapsed = time.time() - start_time
            if elapsed > MAX_RUN_TIME * 60:
                print(f"  已达到运行时间限制: {MAX_RUN_TIME} 分钟")
                break
            
            # 检查是否为空行
            if not row or not row[0]:
                print(f"  第 {current_excel_row} 行为空，停止处理")
                break
            
            magnet = str(row[0]).strip()
            
            print(f"  处理第 {current_excel_row} 行 (批次 {batch_num}, 本次 {processed_in_this_run+1}/{max_rows_to_process})...")
            
            row_data = {
                'magnet': magnet,
                'name': '',
                'count': '',
                'size': '',
                'screenshots': [],
                'temp_files': []
            }
            
            # 处理磁力链接
            if magnet.startswith("magnet:"):
                info = get_magnet_info(magnet)
                if info:
                    row_data['name'] = info["name"]
                    row_data['count'] = info["count"]
                    row_data['size'] = info["size"]
                    row_data['screenshots'] = info["screenshots"]
                else:
                    row_data['name'] = "Error"
            else:
                print(f"    警告: 不是有效磁力链接，跳过")
                row_data['name'] = "Invalid Link"
            
            # 下载截图
            if row_data['screenshots']:
                for s in row_data['screenshots']:
                    img_url = s.get("screenshot")
                    if img_url:
                        tmp_path = download_and_convert_to_jpeg(img_url)
                        if tmp_path:
                            row_data['temp_files'].append(tmp_path)
                            temp_files.append(tmp_path)
            
            batch_data.append(row_data)
            processed_in_this_run += 1
            current_excel_row += 1
            
            # 控制 API 请求频率
            time.sleep(REQUEST_INTERVAL)
            
            # 当批次达到 BATCH_SIZE 时，保存批次
            if len(batch_data) >= BATCH_SIZE:
                file_base = file_path.stem
                save_batch_to_file(batch_data, headers, batch_num, output_dir, file_base)
                
                # 更新进度（已处理的数据行数）
                progress_manager.update_file_state(
                    file_path, 
                    start_row + processed_in_this_run,
                    total_rows,
                    completed=False,
                    batch_num=batch_num
                )
                
                batch_data = []
                batch_num += 1
                
                elapsed = time.time() - start_time
                print(f"  已处理 {start_row + processed_in_this_run}/{total_rows} 行数据，耗时 {elapsed/60:.1f} 分钟")
                print()
        
        # ===== 保存最后一个批次 =====
        if batch_data:
            file_base = file_path.stem
            save_batch_to_file(batch_data, headers, batch_num, output_dir, file_base)
            progress_manager.update_file_state(
                file_path,
                start_row + processed_in_this_run,
                total_rows,
                completed=False,
                batch_num=batch_num
            )
            batch_data = []
        
        # ===== 检查是否全部完成 =====
        total_processed = start_row + processed_in_this_run
        if total_processed >= total_rows:
            progress_manager.mark_file_processed(file_path)
            print(f"  ✓ 文件 {file_path.name} 处理完成！")
        else:
            print(f"  ⏳ 文件 {file_path.name} 部分完成: {total_processed}/{total_rows}")
        
        # ===== 清理临时图片文件 =====
        for tmp in temp_files:
            try:
                if os.path.exists(tmp):
                    os.unlink(tmp)
            except Exception as e:
                print(f"  删除临时文件失败 {tmp}: {e}")
        
        wb_in.close()
        
        # 判断是否应继续处理
        elapsed = time.time() - start_time
        should_continue = (
            processed_in_this_run < max_rows_to_process and
            elapsed < MAX_RUN_TIME * 60 and
            total_processed < total_rows
        )
        
        return processed_in_this_run, should_continue
        
    except Exception as e:
        print(f"  处理文件 {file_path.name} 时出错: {e}")
        import traceback
        traceback.print_exc()
        return 0, False


def main():
    """主函数"""
    print("=" * 60)
    print("磁力链接批量处理器 - GitHub Actions 版本")
    print("=" * 60)
    print(f"运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    print("当前配置:")
    print(f"  REQUEST_INTERVAL: {REQUEST_INTERVAL} 秒")
    print(f"  BATCH_SIZE: {BATCH_SIZE} 行/批次")
    print(f"  MAX_RUN_TIME: {MAX_RUN_TIME} 分钟")
    print(f"  MAX_ROWS_PER_RUN: {MAX_ROWS_PER_RUN} 行")
    print(f"  MAX_RETRIES: {MAX_RETRIES} 次")
    print(f"  TIMEOUT: {TIMEOUT} 秒")
    print("-" * 60)
    
    # 设置目录
    input_dir = Path("Input")
    output_dir = Path("Output")
    
    # 创建目录
    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)
    
    # 初始化进度管理器
    progress_manager = ProgressManager(input_dir, output_dir)
    
    # 获取未处理的文件
    unprocessed_files = get_unprocessed_files(input_dir, progress_manager)
    
    if not unprocessed_files:
        print("所有文件已处理完成！")
        return
    
    print(f"找到 {len(unprocessed_files)} 个未处理的文件:")
    for f in unprocessed_files:
        processed = progress_manager.get_processed_rows(f)
        total = count_data_rows_in_file(f)
        print(f"  - {f.name} (已处理 {processed}/{total} 行数据)")
    print()
    
    # 处理文件
    total_processed = 0
    files_processed = 0
    
    for file_path in unprocessed_files:
        processed, should_continue = process_single_file(
            file_path, output_dir, progress_manager
        )
        total_processed += processed
        files_processed += 1
        
        if should_continue:
            print("\n" + "=" * 60)
            print("⚠️ 已达到运行限制，将在下次运行继续")
            print(f"  本次处理: {processed} 行")
            print(f"  累计处理: {progress_manager.get_total_processed()} 行")
            print("=" * 60)
            return
    
    # 全部完成
    print("\n" + "=" * 60)
    print("🎉 所有文件处理完成！")
    print(f"  处理文件数: {files_processed}")
    print(f"  总处理行数: {total_processed}")
    print(f"  输出目录: {output_dir}")
    print("=" * 60)


if __name__ == "__main__":
    # 支持从环境变量读取配置（GitHub Actions）
    if os.getenv('GITHUB_ACTIONS') == 'true':
        MAX_ROWS_PER_RUN = int(os.getenv('MAX_ROWS_PER_RUN', 200))
        MAX_RUN_TIME = int(os.getenv('MAX_RUN_TIME', 50))
        REQUEST_INTERVAL = float(os.getenv('REQUEST_INTERVAL', 1))
        BATCH_SIZE = int(os.getenv('BATCH_SIZE', 50))
        print(f"GitHub Actions 环境配置:")
        print(f"  MAX_ROWS_PER_RUN: {MAX_ROWS_PER_RUN}")
        print(f"  MAX_RUN_TIME: {MAX_RUN_TIME}")
        print(f"  REQUEST_INTERVAL: {REQUEST_INTERVAL}")
        print(f"  BATCH_SIZE: {BATCH_SIZE}")
        print()
    
    main()
