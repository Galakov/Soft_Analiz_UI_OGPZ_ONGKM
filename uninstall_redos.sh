#!/bin/bash

# Проверка прав суперпользователя
if [ "$EUID" -ne 0 ]; then 
  echo "Пожалуйста, запустите этот скрипт с правами root (sudo ./uninstall_redos.sh)"
  exit 1
fi

echo "Удаление Аналитика УИ ОГПЗ..."

# Удаление исполняемого файла
if [ -f "/usr/local/bin/analytics_ui" ]; then
    rm /usr/local/bin/analytics_ui
    echo "Исполняемый файл удален."
else
    echo "Исполняемый файл не найден."
fi

# Удаление ярлыка
if [ -f "/usr/share/applications/analytics_ui.desktop" ]; then
    rm /usr/share/applications/analytics_ui.desktop
    echo "Ярлык удален."
else
    echo "Ярлык не найден."
fi

# Обновление базы данных desktop-файлов
update-desktop-database /usr/share/applications/

echo "Удаление завершено."
