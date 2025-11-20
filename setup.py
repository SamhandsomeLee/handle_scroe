"""
打包配置文件
支持 Windows 和 macOS
"""
from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="score-import-tool",
    version="1.0.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="成绩导入转换工具 - 基于 PySide6 的现代化 GUI",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/handle_score",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Operating System :: Microsoft :: Windows",
        "Operating System :: MacOS",
    ],
    python_requires=">=3.8",
    install_requires=[
        "PySide6>=6.5.0",
        "xlrd>=2.0.1",
        "xlwt>=1.3.0",
        "openpyxl>=3.1.0",
    ],
    entry_points={
        "console_scripts": [
            "score-import-tool=score_import_gui:main",
        ],
    },
)
