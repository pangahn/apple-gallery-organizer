# -*- coding: utf-8 -*-
from pathlib import Path
from tqdm import tqdm

from PIL import Image
from pillow_heif import register_heif_opener

register_heif_opener()


def heic_to_jpg(input_file: str, output_file: str, quality: int = 100):
    try:
        with Image.open(input_file) as img:
            rgb_image = img.convert("RGB")
            rgb_image.save(output_file, "JPEG", quality=quality, exif=img.info.get("exif", b""))
    except Exception as e:
        raise ValueError(f"Conversion failed: {str(e)}")


def get_subdirectories(directory: str):
    path = Path(directory)
    return [entry for entry in path.iterdir() if entry.is_dir()]


def process_directory(subdir: Path, _del: bool = False):
    heic_files = list(subdir.glob("*.heic"))
    with tqdm(total=len(heic_files), desc="Converting files") as pbar:
        for heic_file in heic_files:
            jpg_file = heic_file.with_suffix(".jpg")
            heic_to_jpg(heic_file, jpg_file)
            if _del:
                heic_file.unlink()
            pbar.update(1)


def main(base_dir: str):
    subdirectories = get_subdirectories(base_dir)
    for subdir in subdirectories:
        print(subdir)
        process_directory(subdir)


if __name__ == "__main__":
    source_folder = "C:\\Users\\yourname\\Desktop\\apple"
    main(source_folder)
