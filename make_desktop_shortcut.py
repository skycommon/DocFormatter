#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""手写 Windows .lnk 快捷方式（不依赖 COM，规避沙箱拦截）。

按验证过的 MS-SHELLINK 最小布局：
- 不设 HasLinkTargetIDList (0x01)，靠 LinkInfo 定位目标；
- LinkInfo 含 Unicode 偏移字段，处理中文路径；
- VolumeLabel 用盘符单字符保持 VolumeID 大小 = 20 字节对齐；
- StringData 用「USHORT 字符数 + UTF-16LE」无 null 终止（加 null 会错位）。
"""
import struct
import os
import sys

CLSID = bytes.fromhex("0102140000000000C000000000000046")


def u16_bytes(s: str) -> bytes:
    """UTF-16LE 字节（不含 null）。"""
    return s.encode("utf-16-le")


def string_data(s: str) -> bytes:
    """StringData 项：USHORT 字符数 + UTF-16LE（无 null 终止）。"""
    b = u16_bytes(s)
    return struct.pack("<H", len(s)) + b


def make_lnk(lnk_path, target, work_dir=None, icon_path=None,
             icon_index=0, name=None):
    if name is None:
        name = os.path.basename(target)
    drive = target[:1] if len(target) >= 2 and target[1] == ":" else "C"

    # ---------- ShellLinkHeader (76) ----------
    link_flags = 0x02 | 0x04 | 0x10 | 0x40 | 0x80  # LinkInfo|Name|WorkDir|Icon|Unicode
    header = b""
    header += struct.pack("<I", 0x4C)          # HeaderSize
    header += CLSID
    header += struct.pack("<I", link_flags)
    header += struct.pack("<I", 0x20)          # FileAttributes ARCHIVE
    header += b"\x00" * 24                     # 3x 8-byte timestamps
    header += struct.pack("<I", 0)             # FileSize
    header += struct.pack("<I", icon_index)    # IconIndex
    header += struct.pack("<I", 1)             # ShowCommand SW_SHOWNORMAL
    header += struct.pack("<H", 0)             # HotKey
    header += struct.pack("<H", 0)             # Reserved1
    header += struct.pack("<I", 0)             # Reserved2
    header += struct.pack("<I", 0)             # Reserved3
    assert len(header) == 76, len(header)

    # ---------- LinkInfo (含 Unicode 字段) ----------
    vol_label = u16_bytes(drive) + b"\x00\x00"  # 单字符盘符 + null，保持 VolumeID=20
    volume_id = b""
    volume_id += struct.pack("<I", 0x10 + len(vol_label))  # VolumeIDSize
    volume_id += struct.pack("<I", 3)                       # DriveType FIXED
    volume_id += struct.pack("<I", 0)                       # DriveSerialNumber
    volume_id += struct.pack("<I", 0x10)                    # VolumeLabelOffset
    volume_id += vol_label
    assert len(volume_id) == 20, len(volume_id)

    local_path = u16_bytes(target) + b"\x00\x00"   # LinkInfo 字符串需 null 终止
    local_path_u = u16_bytes(target) + b"\x00\x00"

    vol_id_off = 0x24                          # 36
    local_off = vol_id_off + len(volume_id)    # 56
    local_u_off = local_off + len(local_path)

    link_info_header = b""
    link_info_header += struct.pack("<I", 0)   # LinkInfoSize (占位)
    link_info_header += struct.pack("<I", 0x24)  # LinkInfoHeaderSize
    link_info_header += struct.pack("<I", 0x01)  # LinkInfoFlags VolumeID+LocalBasePath
    link_info_header += struct.pack("<I", vol_id_off)
    link_info_header += struct.pack("<I", local_off)
    link_info_header += struct.pack("<I", 0)   # CommonNetworkRelativeLinkOffset
    link_info_header += struct.pack("<I", 0)   # CommonNetworkRelativeLinkSize
    link_info_header += struct.pack("<I", local_u_off)        # LocalBasePathOffsetUnicode
    link_info_header += struct.pack("<I", vol_id_off)         # VolumeIDOffsetUnicode

    link_info_body = volume_id + local_path + local_path_u
    link_info = link_info_header + link_info_body
    link_info = struct.pack("<I", len(link_info)) + link_info[4:]

    # ---------- StringData ----------
    strings = b""
    strings += string_data(name)               # NAME
    if work_dir:
        strings += string_data(work_dir)       # WORKINGDIR
    if icon_path:
        strings += string_data(f"{icon_path},{icon_index}")  # ICONLOCATION

    data = header + link_info + strings
    with open(lnk_path, "wb") as f:
        f.write(data)

    # ---------- 自洽校验 ----------
    li_size = struct.unpack("<I", link_info[:4])[0]
    assert li_size == len(link_info), (li_size, len(link_info))
    assert local_u_off + len(local_path_u) <= li_size
    return lnk_path


if __name__ == "__main__":
    target = r"D:\DocFormatter\DocFormatter.exe"
    work_dir = r"D:\DocFormatter"
    icon = r"D:\AI\workbuddy\DocFormatter\app_icon_simple.ico"
    desktop = os.path.join(os.environ.get("USERPROFILE", r"C:\Users\LBX"), "Desktop")
    lnk = os.path.join(desktop, "DocFormatter.lnk")
    out = make_lnk(lnk, target, work_dir=work_dir, icon_path=icon, name="DocFormatter")
    print("written:", out, os.path.getsize(out), "bytes")
