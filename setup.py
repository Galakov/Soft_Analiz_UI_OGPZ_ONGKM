from setuptools import setup, find_packages

setup(
    name="Agent",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        'PyQt5>=5.15.0',
        'pytesseract>=0.3.8',
        'pdf2image>=1.16.0',
        'numpy>=1.21.0',
        'pandas>=1.3.0',
        'opencv-python>=4.5.0',
        'Pillow>=8.3.0',
        'matplotlib>=3.4.0',
        'pdfplumber>=0.7.0',
        'paddleocr>=2.6.0',
        'paddlepaddle>=2.4.0'
    ],
) 