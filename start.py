import os
import sys

# Добавляем путь к корневой директории проекта в PYTHONPATH
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_root)

from Agent.Main import main

if __name__ == '__main__':
    sys.exit(main()) 