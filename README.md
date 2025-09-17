# 🐔 Chik-chik

Портативный Excel-процессор с GUI на PyQt5. Просто запусти — и обрабатывай Excel без формул и макросов!

## ✅ Функции

- Иерархическая нумерация
- Группировка строк
- Автоматическое форматирование
- Поддержка больших файлов (режим “Большой файл”)
- Темная тема
- Портативный .exe — без установки!

## 🚀 Как использовать

1. Скачай `Chik-chik.exe` из раздела [Releases](https://github.com/TimchenkoMI/chik-chik/releases)
2. Запусти
3. Выбери Excel-файл
4. Нажми “▶️ Запустить обработку”

## 🛠️ Как собрать из исходников

```bash
pip install PyQt5 openpyxl qdarkstyle psutil pyinstaller
pyinstaller --name="Chik-chik" --onefile --windowed --icon="chikchik.ico" main.py