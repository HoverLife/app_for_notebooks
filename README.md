# Парсер лотов PDF

Компактное настольное Windows-приложение (PyQt5) для парсинга аукционных PDF и выгрузки в Excel.

## Что улучшено в текущей версии

- Усилен фильтр мусорных/служебных строк (табличные шапки не попадают в итог).
- Все основные столбцы в Excel — **на русском языке**.
- Добавлена пост-обработка качества строк (`ПолнотаЗаполнения`, `ФлагиКачества`) и отсев нечитабельных записей.
- Фото **не выгружаются** в Excel и не сохраняются в отдельные файлы (по запросу).
- Добавлены расширенные тех. столбцы: `Процессор`, `ЧастотаПроцессора`, `ОЗУ`, `Накопитель`, `Видеовыходы`, `Порты`.

## 1) Установка

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 2) Установка Tesseract OCR (Windows)

1. Скачайте Tesseract OCR (например, UB Mannheim):
   https://github.com/UB-Mannheim/tesseract/wiki
2. Установите, например, в:
   `C:\Program Files\Tesseract-OCR\tesseract.exe`
3. Добавьте путь в `PATH`.
4. Проверьте, что `tessdata\rus.traineddata` существует.

## 3) Запуск

### GUI

```bash
python app.py
```

### CLI

```bash
python app.py --no-gui --input "C:\path\to\pdf_folder" --output "output\lots_clean.xlsx"
```

Дополнительно:
- `--ocr-mode Auto|Always|Never`
- `--save-csv`
- `--quality-threshold 0.35`
- `--run-sample-tests`

## 4) Сборка `.exe` (PyInstaller)

Базовая команда:

```bash
pyinstaller --onefile --windowed app.py
```

С встраиванием `tesseract.exe` и `tessdata`:

```bash
pyinstaller --onefile --windowed --add-binary "C:\Path\To\Tesseract-OCR\tesseract.exe;." --add-data "C:\Path\To\Tesseract-OCR\tessdata;tessdata" app.py
```

## 5) Ярлык и установщик

### Ярлык вручную

1. ПКМ по `dist\app.exe` → **Отправить** → **Рабочий стол (создать ярлык)**.
2. Для Пуска перенесите ярлык в `%AppData%\Microsoft\Windows\Start Menu\Programs`.

### Минимальный Inno Setup

```ini
[Setup]
AppName=Парсер лотов PDF
AppVersion=1.1
DefaultDirName={autopf}\PDFLotParser
DefaultGroupName=Парсер лотов PDF
OutputDir=.
OutputBaseFilename=PDFLotParserSetup
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\app.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Парсер лотов PDF"; Filename: "{app}\app.exe"
Name: "{commondesktop}\Парсер лотов PDF"; Filename: "{app}\app.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Дополнительные ярлыки:";
```
