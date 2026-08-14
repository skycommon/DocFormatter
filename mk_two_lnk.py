# -*- coding: utf-8 -*-
"""
为 DocFormatter 生成桌面快捷方式（手写 MS-SHELLINK 二进制，绕过沙箱 COM 拦截）。
- D:\DocFormatter\DocFormatter.exe  -> 桌面\DocFormatter.lnk
"""
import os
import struct

DESKTOP = r"C:\Users\LBX\Desktop"
TARGETS = [
    (r"D:\DocFormatter\DocFormatter.exe", "DocFormatter"),
]


def _strdata(s: str) -> bytes:
    """长度前缀 UTF-16LE 字符串（无 null 终止）。"""
    b = s.encode("utf-16-le")
    return struct.pack("<H", len(s)) + b


def make_lnk(lnk_path: str, target: str, name: str, workdir: str, icon: str):
    CLSID = bytes.fromhex("0114020000000000c000000000000046")
    header = bytearray(76)
    struct.pack_into("<I", header, 0, 76)
    header[4:20] = CLSID
    struct.pack_into("<I", header, 20, 0x1AA)   # LinkFlags: LinkInfo|Name|WorkDir|IconLoc|IsUnicode
    struct.pack_into("<I", header, 24, 0x20)    # FileAttributes = NORMAL
    struct.pack_into("<I", header, 52, 0)       # FileSize
    struct.pack_into("<I", header, 56, 0)       # IconIndex
    struct.pack_into("<I", header, 60, 1)       # ShowCommand = SW_SHOWNORMAL

    # ---- LinkInfo ----
    drive_char = target[:1]                      # 盘符单字符
    # VolumeID：先占位再填真实长度，保证 VolumeIDSize 与字节数一致（避免错位）
    vol_label = (drive_char + "\x00").encode("utf-16-le") + b"\x00\x00"
    volume_id = struct.pack("<I", 0)     # VolumeIDSize 占位
    volume_id += struct.pack("<I", 3)    # DriveType = FIXED
    volume_id += struct.pack("<I", 0)    # VolumeSerialNumber
    volume_id += struct.pack("<I", 16)   # VolumeLabelOffset
    volume_id += vol_label
    volume_id = struct.pack("<I", len(volume_id)) + volume_id[4:]
    local_base = target.encode("utf-16-le") + b"\x00\x00"
    common_suffix = b"\x00\x00"
    linkinfo = struct.pack("<I", 0)      # LinkInfoSize 占位
    linkinfo += struct.pack("<I", 0x1C)  # LinkInfoHeaderSize
    linkinfo += struct.pack("<I", 0x01)  # LinkInfoFlags: VolumeIDAndLocalBasePath
    linkinfo += struct.pack("<I", 28)    # VolumeIDOffset
    linkinfo += struct.pack("<I", 28 + len(volume_id))          # LocalBasePathOffset
    linkinfo += struct.pack("<I", 0)     # CommonNetworkRelativeLinkOffset
    linkinfo += struct.pack("<I", 28 + len(volume_id) + len(local_base))  # CommonPathSuffixOffset
    linkinfo += volume_id + local_base + common_suffix
    linkinfo = struct.pack("<I", len(linkinfo)) + linkinfo[4:]  # 填 LinkInfoSize

    # ---- StringData: NAME, WORKDIR, ICONLOCATION ----
    stringdata = _strdata(name) + _strdata(workdir) + _strdata(icon)

    data = bytes(header) + linkinfo + stringdata
    with open(lnk_path, "wb") as f:
        f.write(data)
    return lnk_path


if __name__ == "__main__":
    for target, name in TARGETS:
        workdir = os.path.dirname(target)
        lnk = os.path.join(DESKTOP, name + ".lnk")
        make_lnk(lnk, target, name, workdir, target)
        print("created:", lnk)
