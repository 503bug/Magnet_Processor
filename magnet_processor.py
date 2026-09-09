#!/usr/bin/env python3
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
import subprocess
import argparse
from datetime import datetime
from pathlib import Path

# ========== 用户可调配置（直接修改此处） ==========
REQUEST_INTERVAL = 1.5          # API 请求间隔（秒）
BATCH_SIZE = 50                 # 每批处理行数
MAX_RETRIES = 5                 # 下载图片重试次数
TIMEOUT = 20                    # 请求超时（秒）
DEFAULT_COL_WIDTH = 80          # 截图列宽（字符数）
DEFAULT_ROW_HEIGHT = 274        # 截图行高（磅）
JPEG_QUALITY = 85               # JPEG 压缩质量
MAX_RUN_TIME = 330              # 单次最大运行时间（分钟）
MAX_ROWS_PER_RUN = 999999       # 单次运行最大处理行数
# ===================================================

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


def git_commit_files(output_dir, batch_num, input_file_base, file_path=None):
    """提交该文件相关的批次和进度（避免并发冲突）"""
    try:
        subprocess.run(['git', 'config', '--global', 'user.name', 'github-actions[bot]'], check=False)
        subprocess.run(['git', 'config', '--global', 'user.email', 'github-actions[bot]@users.noreply.github.com'], check=False)

        # 添加该文件的进度文件
        progress_file = output_dir / f'.progress_{input_file_base}.json'
        if progress_file.exists():
            subprocess.run(['git', 'add', '-f', str(progress_file)], check=False, capture_output=True)

        # 添加当前批次的 Excel 文件（匹配该文件名的所有批次）
        batch_pattern = f'*_{input_file_base}.xlsx'
        for xlsx_file in glob.glob(str(output_dir / batch_pattern)):
            subprocess.run(['git', 'add', '-f', xlsx_file], check=False, capture_output=True)

        # 检查是否有变更
        result = subprocess.run(['git', 'diff', '--staged', '--quiet'], capture_output=True)

        if result.returncode != 0:
            commit_msg = f"更新进度和批次 {batch_num:03d} - {input_file_base} [skip ci]"
            # 多次尝试提交（应对并发推送冲突）
            for attempt in range(5):
                try:
                    subprocess.run(['git', 'commit', '-m', commit_msg, '--no-edit'], check=True, capture_output=True)
                    subprocess.run(['git', 'pull', '--rebase'], check=True, capture_output=True)
                    subprocess.run(['git', 'push'], check=True, capture_output=True)
                    print(f"  ✅ Git 提交成功: 批次 {batch_num:03d} - {input_file_base}")
                    return True
                except subprocess.CalledProcessError as e:
                    print(f"  ⚠️ Git 操作失败 (尝试 {attempt+1}/5): {e}")
                    time.sleep(2)
            else:
                print(f"  ❌ Git 提交最终失败: {input_file_base}")
                return False
        else:
            print(f"  ⏭️  无变更，跳过提交")
            return False

    except Exception as e:
        print(f"  ⚠️ Git 提交异常: {e}")
        return False


class ProgressManager:
    """管理单个文件的处理进度（独立状态文件）"""

    def __init__(self, input_file_path, output_dir="Output"):
        self.input_file = Path(input_file_path)
        self.output_dir = Path(output_dir)
        self.base_name = self.input_file.stem
        self.state_file = self.output_dir / f".progress_{self.base_name}.json"
        self.state = {}
        self.load_state()

    def load_state(self):
        """加载该文件的状态文件"""
        if self.state_file.exists():
            try:
                with open(self.state_file, 'r', encoding='utf-8') as f:
                    self.state = json.load(f)
                print(f"加载进度状态: 已处理 {self.state.get('processed_rows', 0)} 行")
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

    def get_state(self):
        return self.state

    def update_state(self, processed_rows, total_rows=None, completed=False, batch_num=None):
        """更新处理状态"""
        self.state['processed_rows'] = processed_rows
        if total_rows is not None:
            self.state['total_rows'] = total_rows
        self.state['completed'] = completed
        self.state['last_modified'] = datetime.now().isoformat()
        if batch_num is not None:
            if 'batches_created' not in self.state:
                self.state['batches_created'] = []
            if batch_num not in self.state['batches_created']:
                self.state['batches_created'].append(batch_num)
        self.save_state()

    def mark_completed(self):
        self.state['completed'] = True
        self.state['last_modified'] = datetime.now().isoformat()
        self.save_state()


def count_data_rows_in_file(file_path):
    """统计数据行数（不含标题行）"""
    try:
        wb = load_workbook(file_path, read_only=True)
        ws = wb.active
        row_count = sum(1 for _ in ws.iter_rows(min_row=2))
        wb.close()
        return row_count
    except Exception as e:
        print(f"统计行数失败 {file_path.name}: {e}")
        return 0


def get_headers_from_file(file_path):
    """读取标题行"""
    try:
        wb = load_workbook(file_path, read_only=True)
        ws = wb.active
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        wb.close()
        if first_row:
            headers = [str(h) if h else "" for h in first_row[:4]]
            if not any(headers):
                return ["链接", "Resource Name", "Number of Files", "Total File Size"]
            return headers
        else:
            return ["链接", "Resource Name", "Number of Files", "Total File Size"]
    except Exception as e:
        print(f"读取标题行失败 {file_path.name}: {e}")
        return ["链接", "Resource Name", "Number of Files", "Total File Size"]


def save_batch_to_file(batch_data, headers, batch_num, output_dir, base_name):
    """保存批次 Excel"""
    output_filename = f"{batch_num:03d}_{base_name}.xlsx"
    output_path = output_dir / output_filename

    print(f"  📦 保存批次 {batch_num} 到: {output_filename}")
    print(f"  📊 数据行数: {len(batch_data)}")

    workbook = xlsxwriter.Workbook(str(output_path))
    worksheet = workbook.add_worksheet()

    worksheet.set_column(0, 3, 20)

    # 标题
    for col_idx, header in enumerate(headers):
        worksheet.write(0, col_idx, header)

    # 数据
    for row_idx, row_data in enumerate(batch_data):
        excel_row = row_idx + 1
        worksheet.write(excel_row, 0, row_data['magnet'])
        worksheet.write(excel_row, 1, row_data['name'])
        worksheet.write(excel_row, 2, row_data['count'])
        worksheet.write(excel_row, 3, row_data['size'])
        worksheet.set_row(excel_row, DEFAULT_ROW_HEIGHT)

        if row_data['temp_files']:
            col = 4
            for tmp_path in row_data['temp_files']:
                if os.path.exists(tmp_path):
                    worksheet.set_column(col, col, DEFAULT_COL_WIDTH)
                    worksheet.insert_image(excel_row, col, tmp_path, {'x_scale': 1, 'y_scale': 1})
                    col += 1

    workbook.close()
    print(f"  ✅ 批次 {batch_num} 保存完成")


def process_single_file(input_file, output_dir):
    """
    处理单个文件，支持断点续传
    """
    file_path = Path(input_file)
    base_name = file_path.stem
    print(f"开始处理文件: {file_path.name}")

    progress = ProgressManager(file_path, output_dir)
    state = progress.get_state()

    # 如果已完成，直接返回
    if state.get('completed', False):
        print(f"  ✓ 文件 {file_path.name} 已全部处理完成！")
        return

    # 获取总行数
    total_rows = count_data_rows_in_file(file_path)
    start_row = state.get('processed_rows', 0)

    if start_row >= total_rows:
        progress.mark_completed()
        print(f"  ✓ 文件 {file_path.name} 已处理完所有行！")
        return

    print(f"  从第 {start_row + 2} 行继续（已处理 {start_row} 行数据）")
    print(f"  总数据行数: {total_rows}")

    # 打开输入文件
    try:
        wb_in = load_workbook(file_path, read_only=True)
        ws_in = wb_in.active
        headers = get_headers_from_file(file_path)
        print(f"  标题行: {headers}")

        # 计算要处理的行数
        max_rows_to_process = min(MAX_ROWS_PER_RUN, total_rows - start_row)
        batch_num = (start_row // BATCH_SIZE) + 1 if start_row > 0 else 1
        processed_in_this_run = 0
        batch_data = []
        temp_files = []
        start_time = time.time()

        print(f"  本次计划处理: {max_rows_to_process} 行")
        print(f"  起始批次号: {batch_num}")

        # 跳过已处理行
        current_excel_row = start_row + 2
        row_counter = 0
        for row in ws_in.iter_rows(min_row=2, values_only=True):
            if row_counter >= start_row:
                break
            row_counter += 1
            current_excel_row += 1

        # 处理新行
        for row in ws_in.iter_rows(min_row=current_excel_row, values_only=True):
            if processed_in_this_run >= max_rows_to_process:
                print(f"  已达到本次运行行数限制: {max_rows_to_process}")
                break

            elapsed = time.time() - start_time
            if elapsed > MAX_RUN_TIME * 60:
                print(f"  已达到运行时间限制: {MAX_RUN_TIME} 分钟")
                break

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

            time.sleep(REQUEST_INTERVAL)

            # 批量保存
            if len(batch_data) >= BATCH_SIZE:
                save_batch_to_file(batch_data, headers, batch_num, output_dir, base_name)
                # 更新进度
                new_processed = start_row + processed_in_this_run
                progress.update_state(new_processed, total_rows, completed=False, batch_num=batch_num)
                print(f"  💾 进度已保存: {new_processed}/{total_rows}")

                # 提交到 Git（使用独立提交，带重试）
                git_commit_files(output_dir, batch_num, base_name, file_path)

                # 清理临时文件
                for row in batch_data:
                    for tmp in row.get('temp_files', []):
                        try:
                            if os.path.exists(tmp):
                                os.unlink(tmp)
                        except:
                            pass
                batch_data = []
                batch_num += 1

                elapsed = time.time() - start_time
                print(f"  📈 已处理 {new_processed}/{total_rows} 行数据，耗时 {elapsed/60:.1f} 分钟")
                print()

        # 剩余不足一批的数据
        if batch_data:
            save_batch_to_file(batch_data, headers, batch_num, output_dir, base_name)
            new_processed = start_row + processed_in_this_run
            progress.update_state(new_processed, total_rows, completed=False, batch_num=batch_num)
            print(f"  💾 进度已保存: {new_processed}/{total_rows}")
            git_commit_files(output_dir, batch_num, base_name, file_path)
            for row in batch_data:
                for tmp in row.get('temp_files', []):
                    try:
                        if os.path.exists(tmp):
                            os.unlink(tmp)
                    except:
                        pass
            batch_data = []

        # 检查是否全部完成
        final_processed = start_row + processed_in_this_run
        if final_processed >= total_rows:
            progress.mark_completed()
            print(f"  ✓ 文件 {file_path.name} 处理完成！")
            # 最终提交一次（标记完成）
            git_commit_files(output_dir, batch_num, base_name, file_path)
        else:
            print(f"  ⏳ 文件 {file_path.name} 部分完成: {final_processed}/{total_rows}")

        wb_in.close()

    except Exception as e:
        print(f"  处理文件 {file_path.name} 时出错: {e}")
        import traceback
        traceback.print_exc()
        raise


def main():
    parser = argparse.ArgumentParser(description="磁力链接批量处理器 - 单文件模式")
    parser.add_argument('--input-file', required=True, help='要处理的输入Excel文件路径')
    parser.add_argument('--output-dir', default='Output', help='输出目录')
    args = parser.parse_args()

    print("=" * 60)
    print("磁力链接批量处理器 - 并发作业模式")
    print("=" * 60)
    print(f"运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"处理文件: {args.input_file}")
    print()

    input_dir = Path("Input")
    output_dir = Path(args.output_dir)
    input_dir.mkdir(exist_ok=True)
    output_dir.mkdir(exist_ok=True)

    process_single_file(args.input_file, output_dir)


if __name__ == "__main__":
    main()
