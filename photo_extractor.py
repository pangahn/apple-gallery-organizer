# -*- coding: utf-8 -*-
from typing import List
from dataclasses import dataclass
from pathlib import Path
from datetime import datetime

from win32comext.shell import shell, shellcon
from win32com.propsys import propsys

import pythoncom
import pytz


def get_child_shell_folder(parent_shell_folder, child_folder_name: str):
    for child_pidl in parent_shell_folder:
        child_display_name = parent_shell_folder.GetDisplayNameOf(child_pidl, shellcon.SHGDN_NORMAL)
        if child_display_name == child_folder_name:
            return parent_shell_folder.BindToObject(child_pidl, None, shell.IID_IShellFolder)
    raise Exception(f"Cannot find {child_folder_name}")


def get_device_shell_folder(device_display_name):
    current_shell_folder = shell.SHGetDesktopFolder()
    folders = device_display_name.split("\\")
    for folder in folders:
        try:
            current_shell_folder = get_child_shell_folder(current_shell_folder, folder)
        except BaseException as exception:
            raise Exception(f"Cannot get shell folder for {device_display_name} (at {folder})") from exception
    return current_shell_folder


def get_absolute_name(shell_item):
    return shell_item.GetDisplayName(shellcon.SIGDN_DESKTOPABSOLUTEEDITING)


def walk_folder(shell_folder, only_contain: List[str] = None):
    result = {}

    for folder_pidl in shell_folder.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
        child_shell_folder = shell_folder.BindToObject(folder_pidl, None, shell.IID_IShellFolder)
        child_folder_name = shell_folder.GetDisplayNameOf(folder_pidl, shellcon.SHGDN_FORADDRESSBAR)
        if only_contain is None or len(only_contain) == 0 or any([n in child_folder_name for n in only_contain]):
            print(f"Listing folder {child_folder_name}")
            result.update(walk_folder(child_shell_folder, only_contain))
        else:
            print(f"Ignore folder {child_folder_name}")

    for file_pidl in shell_folder.EnumObjects(0, shellcon.SHCONTF_NONFOLDERS):
        folder_pidl = shell.SHGetIDListFromObject(shell_folder)
        shell_item = shell.SHCreateShellItem(folder_pidl, None, file_pidl)
        file_absolute_path = get_absolute_name(shell_item)
        result[file_absolute_path] = shell_item

    return result


@dataclass
class CopyParams:
    sourcefile_shell_item: object
    dest_folder_shell_item: object
    copy_name: str


def remove_prefix(str, prefix):
    if not str.startswith(prefix):
        raise Exception(f"'{str}' should start with '{prefix}")
    return str[len(prefix) :]


def get_shell_item_from_path(path):
    try:
        return shell.SHCreateItemFromParsingName(path, None, shell.IID_IShellItem)
    except BaseException as exception:
        raise Exception(f"Cannot get shell item for {path}") from exception


def convert_to_beijing_time(input_time_str):
    original_time = datetime.strptime(input_time_str, "%Y/%m/%d:%H:%M:%S.%f")
    local_timezone = pytz.timezone("UTC")
    original_time_with_tz = local_timezone.localize(original_time)
    target_timezone = pytz.timezone("Asia/Shanghai")
    beijing_time = original_time_with_tz.astimezone(target_timezone)
    return beijing_time


def get_date(si, property_name: str, output_format: str = "%Y%m%d_%H%M%S"):
    DATE_PROP_KEY = propsys.PSGetPropertyKeyFromName(property_name)
    ps = propsys.PSGetItemPropertyHandler(si)
    date_str = ps.GetValue(DATE_PROP_KEY).ToString()
    if date_str:
        date_time = convert_to_beijing_time(date_str)
        date_str = date_time.strftime(output_format)
    return date_str


def get_copy_params(shell_items: dict, dest_path: str, suffixes: List[str], display_name: str):
    dest_file_names = []
    for file_absolute_path in sorted(shell_items.keys()):
        file_absolute_path = Path(file_absolute_path)
        if file_absolute_path.suffix.lower() not in suffixes:
            continue

        name = file_absolute_path.name
        if name.startswith("IMG_E"):
            if name not in dest_file_names:
                dest_file_names.append(name)

            _name = name.replace("_E", "_")
            if _name in dest_file_names:
                dest_file_names.remove(_name)
                print(f"-- {_name}, ++ {name}")
        else:
            if name not in dest_file_names and name.replace("IMG_", "IMG_E") not in dest_file_names:
                dest_file_names.append(name)

    dest_shell_items = {}
    copy_params_list = []
    copy_names = {}

    for file_absolute_path, file_shell_item in sorted(shell_items.items()):
        file_relative_path = remove_prefix(file_absolute_path, display_name)
        file_relative_path = remove_prefix(file_relative_path, "\\")
        dest_absolute_path = Path(dest_path) / file_relative_path
        name = dest_absolute_path.name
        suffix = dest_absolute_path.suffix.lower()

        if name not in dest_file_names and suffix not in suffixes:
            continue

        dest_folder = str(dest_absolute_path.parent)
        if dest_folder not in dest_shell_items:
            Path(dest_folder).mkdir(parents=True, exist_ok=True)
            target_folder_shell_item = get_shell_item_from_path(dest_folder)
            dest_shell_items[dest_folder] = target_folder_shell_item

        photo_taken_date = get_date(file_shell_item, "System.Photo.DateTaken")
        if photo_taken_date == "":
            copy_name = dest_absolute_path.name
        else:
            copy_name = photo_taken_date
            if copy_name not in copy_names:
                copy_names[copy_name] = 1
            else:
                copy_names[copy_name] += 1
                copy_name = copy_name + f"_{copy_names[copy_name]}"
            copy_name = f"IMG_{copy_name}{dest_absolute_path.suffix.lower()}"

        copy_params = CopyParams(file_shell_item, dest_shell_items[dest_folder], copy_name)
        copy_params_list.append(copy_params)

    return copy_params_list


def copy_files(copy_params_list: list[CopyParams]):
    if copy_params_list:
        fileOperationObject = pythoncom.CoCreateInstance(
            shell.CLSID_FileOperation,
            None,
            pythoncom.CLSCTX_ALL,
            shell.IID_IFileOperation,
        )
        for copy_params in copy_params_list:
            src_str = get_absolute_name(copy_params.sourcefile_shell_item)
            dst_str = get_absolute_name(copy_params.dest_folder_shell_item)

            print(f"Queuing copying {src_str} to {dst_str}\{copy_params.copy_name}")
            fileOperationObject.CopyItem(
                copy_params.sourcefile_shell_item,
                copy_params.dest_folder_shell_item,
                copy_params.copy_name,
            )

        print(f"Running copy operations...")
        fileOperationObject.PerformOperations()

    else:
        print("nothing to do")
    print("done")


if __name__ == "__main__":
    device_display_name = "此电脑\\Apple iPhone\\Internal Storage"
    dest_path = "C:\\Users\\yourname\\Desktop\\apple"

    device_shell_folder = get_device_shell_folder(device_display_name)
    device_shell_items = walk_folder(device_shell_folder, only_contain=["202301"])
    copy_params_list = get_copy_params(device_shell_items, dest_path, [".heic", ".jpg"], device_display_name)
    copy_files(copy_params_list)
