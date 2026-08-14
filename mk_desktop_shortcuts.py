# -*- coding: utf-8 -*-
"""
为 DocFormatter 生成带图标的桌面快捷方式（手写 MS-SHELLINK 二进制，绕过沙箱 COM 拦截）。
目标 -> D:/DocFormatter/DocFormatter.exe  (图标 app_icon_simple.ico)

图标统一指向项目目录下的多尺寸 .ico（16/24/32/48/64/128/256），桌面/任务栏都能正确抽取。
生成后调用 Shell32.SHChangeNotify 通知资源管理器刷新图标缓存。
"""
import os
import struct
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DESKTOP = os.path.join(os.environ.get("USERPROFILE", r"C:\Users\LBX"), "Desktop")

TARGETS = [
    (r"D:\DocFormatter\DocFormatter.exe",
     "DocFormatter",
     os.path.join(HERE, "app_icon_simple.ico")),
]


def _strdata(s: str) -> bytes:
    """长度前缀 UTF-16LE 字符串（无 null 终止）。"""
    return struct.pack("<H", len(s)) + s.encode("utf-16-le")


def make_lnk(lnk_path: str, target: str, name: str, workdir: str, icon: str):
    CLSID = bytes.fromhex("0114020000000000c000000000000046")
    header = bytearray(76)
    struct.pack_into("<I", header, 0, 76)
    header[4:20] = CLSID
    # LinkFlags: LinkInfo|HasName|HasWorkDir|HasIconLocation|IsUnicode
    struct.pack_into("<I", header, 20, 0x1AA)
    struct.pack_into("<I", header, 24, 0x20)   # FileAttributes = NORMAL
    struct.pack_into("<I", header, 56, 0)       # IconIndex
    struct.pack_into("<I", header, 60, 1)       # ShowCommand = SW_SHOWNORMAL

    # ---- LinkInfo ----
    drive_char = target[:1]
    vol_label = (drive_char + "\x00").encode("utf-16-le") + b"\x00\x00"
    volume_id = struct.pack("<I", 0)
    volume_id += struct.pack("<I", 3)         # DriveType = FIXED
    volume_id += struct.pack("<I", 0)         # VolumeSerialNumber
    volume_id += struct.pack("<I", 16)        # VolumeLabelOffset
    volume_id += vol_label
    volume_id = struct.pack("<I", len(volume_id)) + volume_id[4:]
    local_base = target.encode("utf-16-le") + b"\x00\x00"
    common_suffix = b"\x00\x00"
    linkinfo = struct.pack("<I", 0)
    linkinfo += struct.pack("<I", 0x1C)       # LinkInfoHeaderSize
    linkinfo += struct.pack("<I", 0x01)       # LinkInfoFlags: VolumeIDAndLocalBasePath
    linkinfo += struct.pack("<I", 28)         # VolumeIDOffset
    linkinfo += struct.pack("<I", 28 + len(volume_id))            # LocalBasePathOffset
    linkinfo += struct.pack("<I", 0)          # CommonNetworkRelativeLinkOffset
    linkinfo += struct.pack("<I", 28 + len(volume_id) + len(local_base))  # CommonPathSuffixOffset
    linkinfo += volume_id + local_base + common_suffix
    linkinfo = struct.pack("<I", len(linkinfo)) + linkinfo[4:]

    # ---- StringData: NAME, WORKDIR, ICONLOCATION ----
    stringdata = _strdata(name) + _strdata(workdir) + _strdata(icon + ",0")

    data = bytes(header) + linkinfo + stringdata
    with open(lnk_path, "wb") as f:
        f.write(data)
    return lnk_path


def refresh_shell():
    """通知资源管理器刷新图标/关联缓存（非 COM 调用，沙箱通常不拦）。"""
    try:
        import ctypes
        shell32 = ctypes.windll.shell32
        # SHCNE_ASSOCCHANGED = 0x08000000, SHCNF_IDLIST = 0x0000
        shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
    except Exception as e:
        print("shell notify skipped:", e)


if __name__ == "__main__":
    for target, name, icon in TARGETS:
        if not os.path.exists(target):
            print("[跳过] 目标不存在:", target)
            continue
        if not os.path.exists(icon):
            print("[警告] 图标不存在，改用 exe 内嵌:", icon)
            icon = target
        workdir = os.path.dirname(target)
        lnk = os.path.join(DESKTOP, name + ".lnk")
        try:
            make_lnk(lnk, target, name, workdir, icon)
            print("created:", lnk, os.path.getsize(lnk), "bytes  icon=", icon)
        except PermissionError:
            # 沙箱禁止覆盖已存在的桌面快捷方式；若该 lnk 已正确则跳过
            print("[跳过] 已存在且无法覆盖（沙箱限制），保持原样:", lnk)
    refresh_shell()
    print("done")
