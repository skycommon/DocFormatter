# Changelog

All notable changes to DocFormatter are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-08-15

### Added · 首发版本

- **多套排版预设**：默认 / 毕业论文 / 通用论文 / 课程论文 / 简洁风 / 阅读风
- **中英混排字体**：西文 Times New Roman + 中文 宋体 / 等线 / 楷体 / 微软雅黑 等自动分离
- **实时预览面板**：右侧所见即所得，编辑即预览
- **数学公式开关**：OMML 格式（Word 原生）
- **引用与脚注**：上标编号、自动编号
- **页眉 / 页脚 / 页码**：可开关
- **参考文献样式**：GB/T 7714 顺序编码 + 数字序号两种
- **批量处理**：多文件导入、统一应用预设、统一导出
- **导出 PDF**：借助本机 Word / WPS
- **中英双语界面**：运行时实时切换
- **配置导入 / 导出**：JSON 跨电脑迁移
- **常用导出位置**：可勾选的快速目录
- **多尺寸应用图标**：16/24/32/48/64/128/256 全分辨率，桌面/任务栏清晰可见
- **单文件 exe**：PyInstaller 6.22 打包，~18 MB，启动 < 2 秒
- **桌面快捷方式生成器**：手写 MS-SHELLINK 二进制，绕过沙箱 COM 限制
- **主题**：跟随系统 / 浅色 / 深色 三档

### Fixed
- 中英混排字体错位（西文字符不再用中文字体内置拉丁字形）
- `ref_style` 下拉显示 bug（`ref_style_gb7714` → `ref_style_gb`）
- onefile 模式下 `__file__` 指向临时目录导致配置丢失（已用 `APP_DIR` 修正）

### Notes
- 完全离线，模型 / 资源全部内置，无需联网
- 仅支持 Windows 10/11