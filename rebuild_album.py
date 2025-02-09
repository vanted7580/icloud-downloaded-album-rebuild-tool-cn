import os
import csv
import datetime
import logging
import shutil
import hashlib

import win32file, pywintypes, win32con


def set_file_creation_time(file_path, creation_time):
    """
    pip install pywin32
    creation_time: 时间戳（秒）
    """
    try:
        win_time = pywintypes.Time(creation_time)
        handle = win32file.CreateFile(
            file_path,
            win32con.GENERIC_WRITE,
            0,
            None,
            win32con.OPEN_EXISTING,
            0,
            None
        )
        win32file.SetFileTime(handle, win_time, None, None)
        handle.close()
        return True
    except Exception as e:
        logging.error(f"设置创建时间失败 {file_path}: {e}")
        return False


def update_file_times(file_path, creation_time, modification_time):
    """
    更新文件的创建时间和修改时间
    creation_time: 原始创建时间（时间戳，秒）
    modification_time: 导入时间（时间戳，秒）
    """
    try:
        os.utime(file_path, (modification_time, modification_time))
        set_file_creation_time(file_path, creation_time)
        logging.debug(f"更新时间成功: {file_path}")
    except Exception as e:
        logging.error(f"更新文件时间失败 {file_path}: {e}")


def compute_md5(file_path, chunk_size=8192):
    """
    计算文件的 MD5 值
    """
    md5 = hashlib.md5()
    try:
        with open(file_path, 'rb') as f:
            while True:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                md5.update(chunk)
    except Exception as e:
        logging.error(f"计算 MD5 出错 {file_path}: {e}")
        raise
    return md5.hexdigest()


def generate_new_filename(dest_dir, file_name):
    """
    根据目标目录中已存在的同名文件生成新的文件名（只允许重命名一次，即生成 {basename}_1{ext}）
    如果目标中已存在 {basename}_1{ext}，则返回 None 表示冲突（需要人工检查）
    """
    base, ext = os.path.splitext(file_name)
    new_name = f"{base}_1{ext}"
    if os.path.exists(os.path.join(dest_dir, new_name)):
        return None
    return new_name


def parse_date(date_str):
    """
    解析 CSV 中的日期字符串，返回时间戳（秒）
    日期格式示例: "Wednesday October 16,2024 4:32 AM GMT"
    """
    if not date_str:
        return None
    date_str = date_str.strip().strip('"')
    try:
        dt = datetime.datetime.strptime(date_str, "%A %B %d,%Y %I:%M %p GMT")
        return dt.timestamp()
    except Exception as e:
        logging.error(f"解析日期失败 {date_str}: {e}")
        return None


def load_photo_details(details_csv_path):
    """
    加载 Photo Details.csv，返回字典：键为图片文件名，值为记录字典
    """
    photo_details = {}
    try:
        with open(details_csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            for row in reader:
                img_name = row.get("imgName", "").strip()
                if img_name:
                    photo_details[img_name] = row
    except Exception as e:
        logging.error(f"读取 {details_csv_path} 失败: {e}")
    return photo_details


def copy_file_with_md5(src_file, dest_dir):
    """
    将 src_file 从源目录复制到 dest_dir。如果目标中已有同名文件，则：
      - 如果文件大小相同且 MD5 相同，则直接返回目标文件路径；
      - 如果文件内容不同，则尝试生成新的文件名（仅允许重命名一次，生成文件名末尾带 "_1"），
        如果该重命名文件已存在，则跳过并记录警告，返回 None。
    成功时返回目标中的文件路径，否则返回 None。
    注意：此处为复制操作，不删除源文件。
    """
    file_name = os.path.basename(src_file)
    dest_file = os.path.join(dest_dir, file_name)

    if os.path.exists(dest_file):
        try:
            src_size = os.path.getsize(src_file)
            dest_size = os.path.getsize(dest_file)
        except Exception as e:
            logging.error(f"获取文件大小失败: {src_file} 或 {dest_file}: {e}")
            return None

        if src_size == dest_size:
            try:
                src_md5 = compute_md5(src_file)
                dest_md5 = compute_md5(dest_file)
            except Exception as e:
                logging.error(f"计算 MD5 出错 {src_file} 或 {dest_file}: {e}")
                return None

            if src_md5 == dest_md5:
                logging.info(f"文件已存在且内容相同: {dest_file}")
                return dest_file
            else:
                new_name = generate_new_filename(dest_dir, file_name)
                if new_name is None:
                    logging.warning(
                        f"重命名冲突：目标目录中已有 {file_name} 与 {os.path.splitext(file_name)[0]}_1{os.path.splitext(file_name)[1]}，跳过 {src_file}")
                    return None
                new_dest_file = os.path.join(dest_dir, new_name)
                try:
                    shutil.copy2(src_file, new_dest_file)
                    logging.info(f"文件内容不同，将 {src_file} 复制并重命名为 {new_dest_file}")
                    return new_dest_file
                except Exception as e:
                    logging.error(f"复制文件失败 {src_file} 到 {new_dest_file}: {e}")
                    return None
        else:
            new_name = generate_new_filename(dest_dir, file_name)
            if new_name is None:
                logging.warning(f"大小不同但重命名冲突，跳过 {src_file}")
                return None
            new_dest_file = os.path.join(dest_dir, new_name)
            try:
                shutil.copy2(src_file, new_dest_file)
                logging.info(f"大小不同，将 {src_file} 复制并重命名为 {new_dest_file}")
                return new_dest_file
            except Exception as e:
                logging.error(f"复制文件失败 {src_file} 到 {new_dest_file}: {e}")
                return None
    else:
        try:
            shutil.copy2(src_file, dest_file)
            # 验证目标文件是否存在
            if os.path.exists(dest_file):
                logging.debug(f"目标文件 {dest_file} 校验成功！")
            else:
                logging.error(f"目标文件 {dest_file} 校验失败！")
            return dest_file
        except Exception as e:
            logging.error(f"复制失败: {src_file} -> {dest_file}: {e}")
            return None


def process_part_phase1(part_number, source_root, target_root):
    """
    阶段1：处理单个部分，将 Photos 目录中的照片（更新时间后）从源文件夹复制到目标根目录，
    源文件夹（"iCloud 照片 第 {n} 部分（共 101 部分）/Photos"）内的照片保持不变。
    """
    part_folder = os.path.join(source_root, f"iCloud 照片 第 {part_number} 部分（共 101 部分）")
    photos_dir = os.path.join(part_folder, "Photos")
    if not os.path.exists(photos_dir):
        logging.warning(f"未找到 Photos 文件夹: {photos_dir}")
        return

    details_csv_path = os.path.join(photos_dir, "Photo Details.csv")
    photo_details = {}
    if os.path.exists(details_csv_path):
        photo_details = load_photo_details(details_csv_path)
    else:
        logging.info(f"未找到 Photo Details.csv: {details_csv_path}")

    for entry in os.listdir(photos_dir):
        src_file = os.path.join(photos_dir, entry)
        # 跳过 CSV 文件和子目录
        if os.path.isdir(src_file) or entry.lower().endswith(".csv"):
            continue
        if not os.path.isfile(src_file):
            continue

        # 若在 Photo Details 中有记录，则更新文件时间
        details = photo_details.get(entry)
        if details:
            orig_date_str = details.get("originalCreationDate", "")
            import_date_str = details.get("importDate", "")
            creation_time = parse_date(orig_date_str)
            import_time = parse_date(import_date_str)
            if creation_time and import_time:
                update_file_times(src_file, creation_time, import_time)

        copied_path = copy_file_with_md5(src_file, target_root)
        if copied_path is None:
            logging.warning(f"文件复制失败或跳过: {src_file}")


def build_global_album_mapping(source_root):
    """
    遍历所有部分的 Albums 文件夹，读取各个 CSV 文件，
    返回字典：键为照片名（例如 "IMG_8372.JPG"），值为所属相册名称的集合。
    """
    albums_mapping = {}
    for part in range(1, 102):
        part_folder = os.path.join(source_root, f"iCloud 照片 第 {part} 部分（共 101 部分）")
        albums_dir = os.path.join(part_folder, "Albums")
        if not os.path.exists(albums_dir) or not os.path.isdir(albums_dir):
            continue
        for file in os.listdir(albums_dir):
            if not file.lower().endswith(".csv"):
                continue
            album_name = os.path.splitext(file)[0]
            album_csv_path = os.path.join(albums_dir, file)
            try:
                with open(album_csv_path, 'r', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    header = next(reader, None)  # 跳过表头（如果存在）
                    for row in reader:
                        if not row:
                            continue
                        image_name = row[0].strip().strip('"')
                        if image_name:
                            albums_mapping.setdefault(image_name, set()).add(album_name)
            except Exception as e:
                logging.error(f"读取相册文件 {album_csv_path} 失败: {e}")
    return albums_mapping


def copy_file_to_album(src_file, album_folder):
    """
    将 src_file 从目标根目录复制到 album_folder。
    复制时若目标中已有同名文件，则：
      - 如果内容相同，则认为已存在，不再复制；
      - 如果内容不同，则尝试以重命名方式（末尾加 "_1"）复制，但只允许一次重命名，
        如果目标中已存在重命名文件，则跳过并记录警告。
    返回 True 表示复制成功或已存在，False 表示跳过或复制失败。
    """
    file_name = os.path.basename(src_file)
    base, ext = os.path.splitext(file_name)
    candidate_renamed = f"{base}_1{ext}"
    if os.path.exists(os.path.join(album_folder, candidate_renamed)):
        logging.warning(f"在相册 {album_folder} 中已存在重命名文件 {candidate_renamed}，跳过复制 {src_file}")
        return False

    dest_file = os.path.join(album_folder, file_name)
    if os.path.exists(dest_file):
        try:
            src_size = os.path.getsize(src_file)
            dest_size = os.path.getsize(dest_file)
        except Exception as e:
            logging.error(f"获取文件大小失败 {src_file} 或 {dest_file}: {e}")
            return False

        if src_size == dest_size:
            try:
                src_md5 = compute_md5(src_file)
                dest_md5 = compute_md5(dest_file)
            except Exception as e:
                logging.error(f"计算 MD5 失败 {src_file} 或 {dest_file}: {e}")
                return False
            if src_md5 == dest_md5:
                logging.info(f"相册中已存在相同文件: {dest_file}")
                return True
            else:
                renamed_dest = os.path.join(album_folder, candidate_renamed)
                if os.path.exists(renamed_dest):
                    logging.warning(f"在相册 {album_folder} 中重命名文件 {candidate_renamed} 已存在，跳过 {src_file}")
                    return False
                try:
                    shutil.copy2(src_file, renamed_dest)
                    logging.info(f"文件 {src_file} 以重命名形式复制到 {renamed_dest} (内容不同)")
                    return True
                except Exception as e:
                    logging.error(f"复制 {src_file} 到 {renamed_dest} 失败: {e}")
                    return False
        else:
            renamed_dest = os.path.join(album_folder, candidate_renamed)
            if os.path.exists(renamed_dest):
                logging.warning(
                    f"在相册 {album_folder} 中重命名文件 {candidate_renamed} 已存在（大小不同），跳过 {src_file}")
                return False
            try:
                shutil.copy2(src_file, renamed_dest)
                if os.path.exists(renamed_dest):
                    logging.debug(f"目标文件 {renamed_dest} 校验成功！")
                else:
                    logging.error(f"目标文件 {renamed_dest} 校验失败！")
                return True
            except Exception as e:
                logging.error(f"复制 {src_file} 到 {renamed_dest} 失败: {e}")
                return False
    else:
        try:
            shutil.copy2(src_file, dest_file)
            if os.path.exists(dest_file):
                logging.debug(f"目标文件 {dest_file} 校验成功！")
            else:
                logging.error(f"目标文件 {dest_file} 校验失败！")
            return True
        except Exception as e:
            logging.error(f"复制 {src_file} 到 {dest_file} 失败: {e}")
            return False


def process_album_image(image_name, album_set, target_root, allowed_albums=None):
    """
    根据传入的 image_name 和该照片所属的 album_set，
    将照片复制到各个相册文件夹中。如果参数 allowed_albums 指定了要复制的相册，则只复制这些相册。
    同时检测是否存在与该图片同名的 .MOV 文件（即 Live Photo）。
    如果发现主图或 Live Photo 已重命名（带有 "_1" 后缀），则直接跳过该照片。
    只有在至少复制成功到一个相册后，才删除目标根目录中的主图和 Live Photo文件。
    """
    orig_path = os.path.join(target_root, image_name)
    base, ext = os.path.splitext(image_name)
    renamed_name = f"{base}_1{ext}"
    renamed_path = os.path.join(target_root, renamed_name)

    # 如果目标中存在已重命名的主图，则跳过
    if os.path.exists(renamed_path):
        logging.warning(f"文件 {image_name} 已经被重命名为 {renamed_name}，跳过 album 复制。")
        return
    if not os.path.exists(orig_path):
        logging.warning(f"目标根目录中未找到照片 {image_name}，跳过。")
        return
    file_to_copy = orig_path

    # 检查对应的 Live Photo (.MOV 文件)
    live_photo_orig = os.path.join(target_root, base + ".MOV")
    live_photo_renamed = os.path.join(target_root, base + "_1.MOV")
    if os.path.exists(live_photo_renamed):
        logging.warning(f"Live photo for {image_name} 已重命名为 {base}_1.MOV，跳过 album 复制。")
        return
    live_photo_to_copy = live_photo_orig if os.path.exists(live_photo_orig) else None

    copied_to_any_album = False  # 记录是否至少复制到一个相册

    # 遍历 album_set，根据 allowed_albums 过滤
    for album in album_set:
        if allowed_albums is not None and album not in allowed_albums:
            logging.debug(f"跳过不在允许列表中的相册 {album}")
            continue
        album_folder = os.path.join(target_root, album)
        if not os.path.exists(album_folder):
            try:
                os.makedirs(album_folder, exist_ok=True)
                logging.info(f"创建相册文件夹: {album_folder}")
            except Exception as e:
                logging.error(f"创建相册文件夹失败 {album_folder}: {e}")
                continue
        success_image = copy_file_to_album(file_to_copy, album_folder)
        if success_image:
            copied_to_any_album = True
        else:
            logging.warning(f"复制 {file_to_copy} 到相册 {album_folder} 失败或被跳过。")
        if live_photo_to_copy:
            success_live = copy_file_to_album(live_photo_to_copy, album_folder)
            if success_live:
                copied_to_any_album = True
            else:
                logging.warning(f"复制 live photo {live_photo_to_copy} 到相册 {album_folder} 失败或被跳过。")
    # 只有在至少复制到一个相册后，才删除原文件
    if copied_to_any_album:
        try:
            os.remove(file_to_copy)
            logging.debug(f"删除目标根目录中的文件 {file_to_copy}")
        except Exception as e:
            logging.error(f"删除目标根目录文件失败 {file_to_copy}: {e}")
        if live_photo_to_copy:
            try:
                os.remove(live_photo_to_copy)
                logging.debug(f"删除目标根目录中的 live photo 文件 {live_photo_to_copy}")
            except Exception as e:
                logging.error(f"删除 live photo 文件 {live_photo_to_copy} 失败: {e}")


def main():
    # 请根据实际情况修改以下路径
    source_root = r"D:\Download\数据和隐私"
    target_root = r"D:\Download\Photos"

    # 仅复制指定相册功能：
    # 若只希望复制特定相册，请在此处定义 allowed_albums，例如：
    # allowed_albums = {"AlbumA", "AlbumB"}
    # 若不限制复制所有相册，则设置为 None
    allowed_albums = None  # 例如：allowed_albums = {"AlbumA"}

    os.makedirs(target_root, exist_ok=True)

    logging.info("========== 阶段1：将所有照片从各部分复制到目标根目录 ==========")
    for part in range(1, 102):
        logging.info(f"处理第 {part} 部分")
        process_part_phase1(part, source_root, target_root)
    logging.info("阶段1完成。")

    logging.info("========== 阶段2：根据 Albums 信息整理相册 ==========")
    albums_mapping = build_global_album_mapping(source_root)
    for image_name, album_set in albums_mapping.items():
        process_album_image(image_name, album_set, target_root, allowed_albums)
    logging.info("阶段2完成。")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[logging.StreamHandler()]
    )
    main()
