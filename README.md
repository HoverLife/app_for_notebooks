# PDF Lot Parser

Минималистичное настольное Windows-приложение (PyQt5) для парсинга аукционных PDF-карточек техники и экспорта в Excel/CSV.

## 1) Установка

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## 2) Установка Tesseract OCR на Windows

1. Скачайте установщик Tesseract OCR (например, UB Mannheim build):  
   https://github.com/UB-Mannheim/tesseract/wiki
2. Установите в стандартную папку, например:
   `C:\Program Files\Tesseract-OCR\tesseract.exe`
3. Добавьте путь в `PATH` (или оставьте стандартный путь — приложение попробует найти его автоматически).
4. Для русского OCR убедитесь, что есть `tessdata\rus.traineddata`.

## 3) Запуск

### GUI (по умолчанию)

```bash
python app.py
```

### CLI режим

```bash
python app.py --no-gui --input "C:\path\to\pdf_folder" --output "output\parsed_lots.xlsx"
```

Дополнительно:
- `--ocr-mode Auto|Always|Never`
- `--no-images`
- `--save-csv`
- `--run-sample-tests`

## 4) Сборка `.exe` через PyInstaller

### Базовая команда (оконный режим, без консоли)

```bash
pyinstaller --onefile --windowed app.py
```

### Пример со встраиванием `tesseract.exe` и `tessdata`

```bash
pyinstaller --onefile --windowed --add-binary "C:\Path\To\Tesseract-OCR\tesseract.exe;." --add-data "C:\Path\To\Tesseract-OCR\tessdata;tessdata" app.py
```

> После сборки запускайте `dist\app.exe` двойным кликом (через ярлык/Пуск/рабочий стол).

## 5) Ярлык и установщик

### Создать ярлык вручную

1. ПКМ по `dist\app.exe` → **Отправить** → **Рабочий стол (создать ярлык)**.
2. Для меню Пуск: скопируйте ярлык в `%AppData%\Microsoft\Windows\Start Menu\Programs`.

### Минимальный пример Inno Setup

```ini
[Setup]
AppName=PDF Lot Parser
AppVersion=1.0
DefaultDirName={autopf}\PDFLotParser
DefaultGroupName=PDF Lot Parser
OutputDir=.
OutputBaseFilename=PDFLotParserSetup
Compression=lzma
SolidCompression=yes

[Files]
Source: "dist\app.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\PDF Lot Parser"; Filename: "{app}\app.exe"
Name: "{commondesktop}\PDF Lot Parser"; Filename: "{app}\app.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Дополнительные ярлыки:";
```

## 6) Подпись `.exe` (рекомендация)

Для корпоративного распространения желательно подписать `app.exe` код-подписью (Authenticode), чтобы снизить предупреждения SmartScreen.
