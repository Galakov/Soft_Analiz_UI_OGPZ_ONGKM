@echo off
pip install build wheel > build_log.txt 2>&1
python -m build >> build_log.txt 2>&1
