#!/bin/bash

# Проверка прав суперпользователя
if [ "$EUID" -ne 0 ]; then 
  echo "Пожалуйста, запустите этот скрипт с правами root (sudo ./install_redos.sh)"
  exit 1
fi

echo "Начинаем установку Аналитика УИ ОГПЗ..."

# Установка необходимых зависимостей
echo "Установка зависимостей..."
dnf install -y python3 python3-pip python3-tkinter mesa-libGL

# Установка PyInstaller
echo "Установка PyInstaller..."
pip3 install pyinstaller

# Сборка приложения
echo "Сборка приложения..."
# Удаляем старые сборки если есть
rm -rf build dist
pyinstaller excel_merger.spec

# Проверка успешности сборки
if [ ! -f "dist/Аналитика УИ ОГПЗ v1.2" ]; then
    echo "Ошибка: Файл приложения не был создан. Проверьте вывод PyInstaller."
    exit 1
fi

# Копирование исполняемого файла
echo "Копирование файлов..."
cp "dist/Аналитика УИ ОГПЗ v1.2" /usr/local/bin/analytics_ui
chmod +x /usr/local/bin/analytics_ui

# Создание .desktop файла
echo "Создание ярлыка..."
cat > /usr/share/applications/analytics_ui.desktop << EOL
[Desktop Entry]
Version=1.0
Type=Application
Name=Аналитика УИ ОГПЗ
Comment=Программа для объединения и анализа Excel файлов
Exec=/usr/local/bin/analytics_ui
Icon=utilities-terminal
Terminal=false
Categories=Utility;Office;
EOL

# Обновление базы данных desktop-файлов
update-desktop-database /usr/share/applications/

echo "Установка завершена успешно!"
echo "Теперь вы можете найти программу 'Аналитика УИ ОГПЗ' в меню приложений."
