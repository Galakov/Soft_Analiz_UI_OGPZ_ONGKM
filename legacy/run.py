#!/usr/bin/env python3
import sys
import os

# Добавляем текущую директорию в PYTHONPATH
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

from Agent.Main import main

if __name__ == '__main__':
    sys.exit(main()) 