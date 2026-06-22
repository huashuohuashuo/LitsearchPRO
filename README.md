# LitSearchPro 通用机构版完整源码

版本：22.1.21-generic

本源码包包含：

- `LitSearchPro_Generic_Client.py`：桌面客户端完整源码。
- `LitSearchPro_Generic_Server.py`：协作服务器、API 与轻量网页完整源码。
- `LitSearchPro_Generic_Setup.py`：客户端安装器源码。
- `LitSearchPro_Generic_Uninstall.py`：客户端卸载器源码。
- `LitSearchPro_Generic_Server_Setup.py`：服务器安装器源码。
- `LitSearchPro_Generic_Server_Uninstall.py`：服务器卸载器源码。
- 六个 `.spec` 文件：主程序、服务器、安装器和卸载器的 PyInstaller 构建配置。
- `generate_generic_logo.py`：中性 LSP 图标生成源码。
- `build_all.py`：完整构建脚本。
- `build_source_archive.py`：仅生成完整源码 ZIP，不生成 EXE。

## 运行源码

```powershell
python LitSearchPro_Generic_Client.py
python LitSearchPro_Generic_Server.py
```

主要依赖包括 `Pillow`、`openpyxl`、`PyMuPDF`、`pandas`、`python-docx` 和 `matplotlib`。打包时还需要 `PyInstaller`。

## 构建全部程序

```powershell
python build_all.py
```

构建脚本依次生成客户端、服务器、两个卸载器和两个安装器。安装器构建前会把卸载器复制到 `installer_payload`。

## 只生成源码包

```powershell
python build_source_archive.py
```

该命令只生成 ZIP 源码包，不生成 EXE。

## 品牌定制

通用品牌名称集中在客户端的 `DISPLAY_NAME`、服务器的 `SERVER_NAME` 以及四个安装/卸载源码顶部常量中。替换名称和 `generic_logo.png/.ico` 后即可制作其它机构版本。
