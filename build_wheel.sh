#!/bin/bash

# Установка инструментов для сборки
pip install build wheel

# Очистка предыдущих сборок
rm -rf dist build *.egg-info

# Сборка пакета
python3 -m build

echo "Сборка завершена. Файл .whl находится в папке dist/"
