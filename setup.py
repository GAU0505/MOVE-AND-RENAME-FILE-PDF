from setuptools import setup, find_packages

setup(
    name="pdf-management-app",
    version="0.1",
    author="Hồ Xuân Ánh",                        # Tác giả
    author_email="hoxuananh134@gmail.com",       # Email của tác giả
    description="hỗ trợ công việc",              # Mô tả ngắn
    long_description=open("README.md", encoding="utf-8").read(),   # Mô tả đầy đủ với mã hóa UTF-8
    long_description_content_type="text/markdown",
    url="https://github.com/GAU0505/MOVE-AND-RENAME-FILE-PDF",  # URL repo
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    classifiers=[                           # Phân loại gói
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',                # Phiên bản Python yêu cầu
    install_requires=[
        # Đảm bảo rằng tkinter đã bị loại bỏ
    ],
    entry_points={
        "console_scripts": [
            "pdf-management-app=main:MainApp",
        ],
    },
)