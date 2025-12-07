from setuptools import setup, find_packages

setup(
    name="analytics_ui_ogpz",
    version="1.2.3",
    packages=find_packages(),
    install_requires=[
        'pandas',
        'openpyxl',
        'xlrd',
        'numpy',
    ],
    entry_points={
        'console_scripts': [
            'analytics-ui=analytics_ui.excel_merger:main',
            'analytics-ui-setup=analytics_ui.post_install:create_shortcuts',
        ],
    },
    package_data={
        'analytics_ui': ['*.xlsx'],
    },
    include_package_data=True,
    author="User",
    description="Tool for merging and analyzing Excel files",
)