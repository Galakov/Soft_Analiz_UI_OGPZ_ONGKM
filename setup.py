from setuptools import setup, find_packages

setup(
    name="analytics_ui_ogpz",
    version="1.2.5",
    packages=find_packages(),
    install_requires=[
        'pandas>=1.0',
        'openpyxl>=3.0',
        'xlrd>=2.0',
        'numpy>=1.18',
    ],
    entry_points={
        'console_scripts': [
            'analytics-ui=analytics_ui.excel_merger:main',
            'analytics-ui-setup=analytics_ui.post_install:create_shortcuts',
            'analytics-ui-uninstall=analytics_ui.post_install:remove_shortcuts',
        ],
    },
    package_data={
        'analytics_ui': ['*.xlsx', '*.png'],
    },
    include_package_data=True,
    author="User",
    description="Tool for merging and analyzing Excel files",
)