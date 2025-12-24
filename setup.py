"""
Setup script for ppt-to-images package.
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read README
readme_path = Path(__file__).parent / "README.md"
long_description = ""
if readme_path.exists():
    long_description = readme_path.read_text(encoding="utf-8")

setup(
    name="ppt-to-images",
    version="0.1.0",
    description="Convert PPT/PPTX/PDF files to image sequences",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="HQIT",
    author_email="",
    url="https://github.com/HQIT/ppt-to-images",
    packages=find_packages(),
    install_requires=[
        "pdf2image>=1.16.0",
        "Pillow>=9.0.0",
        "python-pptx>=0.6.21",
    ],
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "ppt-to-images=ppt_to_images.cli:main",
        ],
    },
    python_requires=">=3.7",
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Multimedia :: Graphics :: Graphics Conversion",
        "Topic :: Office/Business :: Office Suites",
    ],
    keywords="ppt pptx pdf image conversion slides presentation",
)

