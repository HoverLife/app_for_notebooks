#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Парсер лотов PDF: GUI + CLI (Windows-friendly, PyInstaller-ready)."""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image
from dateutil import parser as date_parser
from PyQt5 import QtCore, QtGui, QtWidgets

APP_TITLE = "Парсер лотов PDF"
DEFAULT_OUTPUT_DIR = "output"
DEFAULT_ERROR_FILE = "parsing_errors.csv"


BASE_COLUMNS_RU = [
    "ФайлИсточник",
    "СтраницаИсточник",
    "ДатаОбработки",
    "НомерЛота",
    "ГородАукциона",
    "ДатаОкончанияТоргов",
    "ТипТехники",
    "Бренд",
    "Модель",
    "Диагональ",
    "СерийныйНомер",
    "СтартоваяЦенаРуб",
    "АдресСамовывоза",
    "ГородСамовывоза",
    "Видеовходы",
    "Видеовыходы",
    "ТипПитания",
    "ПитаниеВКомплекте",
    "ПроблемыМатрицы",
    "Дефект1",
    "Дефект2",
    "Процессор",
    "ЧастотаПроцессора",
    "ОЗУ",
    "Накопитель",
    "Порты",
    "СостояниеОценка",
    "КлючевыеСлова",
    "ПолнотаЗаполнения",
    "ФлагиКачества",
    "ИзвлеченныеПараметрыJSON",
    "СыройТекст",
    "Примечание",
]

KNOWN_BRANDS = {
    "AOC",
    "SAMSUNG",
    "LG",
    "PHILIPS",
    "DELL",
    "HP",
    "LENOVO",
    "ACER",
    "ASUS",
    "BENQ",
    "IIYAMA",
    "VIEWSONIC",
    "SONY",
    "APPLE",
}


@dataclass
class ParserOptions:
    ocr_mode: str = "Auto"  # Auto | Always | Never
    output_folder: Path = Path(DEFAULT_OUTPUT_DIR)
    auto_start_drop: bool = False
    log_level: str = "INFO"
    save_csv: bool = False
    quality_threshold: float = 0.35


@dataclass
class ParseError:
    source_file: str
    source_page: int
    lot_position: Optional[int]
    message: str
    raw_text: str


class TesseractResolver:
    @staticmethod
    def resolve(logger: logging.Logger) -> Optional[str]:
        candidates: List[Path] = []
        if hasattr(sys, "_MEIPASS"):
            meipass = Path(getattr(sys, "_MEIPASS"))
            candidates += [meipass / "tesseract.exe", meipass / "Tesseract-OCR" / "tesseract.exe"]
        candidates += [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ]
        for c in candidates:
            if c.exists():
                pytesseract.pytesseract.tesseract_cmd = str(c)
                logger.info("Найден Tesseract: %s", c)
                return str(c)

        from_path = shutil.which("tesseract")
        if from_path:
            pytesseract.pytesseract.tesseract_cmd = from_path
            logger.info("Найден Tesseract в PATH: %s", from_path)
            return from_path

        logger.warning("Tesseract не найден. OCR fallback будет недоступен.")
        return None


class LotParser:
    def __init__(self, options: ParserOptions, logger: logging.Logger):
        self.options = options
        self.logger = logger
        self.errors: List[ParseError] = []
        self.tesseract_path = TesseractResolver.resolve(logger)

    @staticmethod
    def normalize_text(text: str) -> str:
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = text.replace("—", "-").replace("–", "-")
        text = re.sub(r"[\t\f\v]+", " ", text)
        text = re.sub(r"[ ]{2,}", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        # Склейка случаев "1\n---"
        text = re.sub(r"(?m)^(\d+)\s*\n\s*---", r"\1---", text)
        # Убираем служебные строки табличной шапки
        garbage = [
            r"^Позиция$",
            r"^лота$",
            r"^Наименование оборудования$",
            r"^Стартовая$",
            r"^цена$",
            r"^продажи,$",
            r"^руб\.?$",
            r"^Город и место$",
            r"^получения$",
            r"^\(самовывоз\)$",
            r"^Серийный номер$",
            r"^Фото\s*\d+$",
            r"^Дефект$",
            r"^[12]$",
        ]
        lines = [ln.strip() for ln in text.split("\n")]
        cleaned = []
        for ln in lines:
            if any(re.match(g, ln, flags=re.IGNORECASE) for g in garbage):
                continue
            cleaned.append(ln)
        text = "\n".join(cleaned)
        return text.strip()

    def _need_ocr(self, text: str) -> bool:
        return len(re.sub(r"\s+", "", text)) < 80

    def _ocr_page(self, page: fitz.Page) -> str:
        if not self.tesseract_path:
            return ""
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        mode = "RGBA" if pix.alpha else "RGB"
        img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
        try:
            return pytesseract.image_to_string(img, lang="rus+eng")
        except Exception as exc:
            self.logger.error("Ошибка OCR: %s", exc)
            return ""

    def _extract_auction_meta(self, text: str) -> Tuple[str, Optional[datetime]]:
        city = ""
        dt = None
        m_city = re.search(r"самовывозом из города\s*:\s*([^\n]+)", text, re.IGNORECASE)
        if m_city:
            city = m_city.group(1).strip()

        m_dt = re.search(r"Дата и время окончания торгов\s*:\s*([^\n]+)", text, re.IGNORECASE)
        if m_dt:
            try:
                dt = date_parser.parse(m_dt.group(1).strip(), dayfirst=True)
            except Exception:
                dt = None
        return city, dt

    def split_into_lots(self, text: str) -> List[str]:
        text = self.normalize_text(text)
        pattern = re.compile(r"(?m)(?=^\s*\d+\s*---)")
        blocks = [b.strip() for b in pattern.split(text) if b.strip()]
        good = []
        for b in blocks:
            if re.match(r"^\d+\s*---", b):
                good.append(b)
        return good

    def _extract_number(self, block: str) -> Optional[int]:
        m = re.match(r"^\s*(\d+)\s*---", block)
        if not m:
            return None
        return int(m.group(1))

    def _extract_brand(self, text: str) -> str:
        up = text.upper()
        up = up.replace("АРРLE", "APPLE")
        for b in KNOWN_BRANDS:
            if re.search(rf"\b{re.escape(b)}\b", up):
                return b
        if re.search(r"\bMACBOOK\b|\bIPHONE\b|\bIMAC\b", up):
            return "APPLE"
        for b in KNOWN_BRANDS:
            if re.search(rf"\b{re.escape(b)}\b", up):
                return b
        return ""

    def _extract_item_type(self, text: str) -> str:
        for t in ["Монитор", "Ноутбук", "Телевизор", "Принтер", "ПК", "Системный блок"]:
            if re.search(rf"\b{re.escape(t)}\b", text, re.IGNORECASE):
                return t
        return ""

    def _extract_model(self, text: str, brand: str) -> str:
        if brand:
            m = re.search(rf"\b{re.escape(brand)}\b\s+([A-Za-z0-9\-_/]+)", text, re.IGNORECASE)
            if m:
                return m.group(1)
        m2 = re.search(r"(?:Монитор|Ноутбук|Телевизор|Принтер)\s+[A-Za-zА-Яа-я]+\s*([A-Za-z0-9\-_/]{3,})", text)
        return m2.group(1) if m2 else ""

    def _extract_diagonal(self, text: str) -> str:
        m = re.search(r"(\d{1,2}(?:[\.,]\d)?)\s*(?:'|\")", text)
        if not m:
            m = re.search(r"Диагональ\s*-\s*(\d{1,2}(?:[\.,]\d)?)", text, re.IGNORECASE)
        if not m:
            return ""
        return f"{m.group(1).replace(',', '.')}\""

    def _extract_serial(self, text: str) -> str:
        m = re.search(r"s\s*/\s*n\s*:\s*([A-Za-z0-9\-]{8,})", text, re.IGNORECASE)
        if m:
            return m.group(1).upper()
        m2 = re.search(r"\b([A-Z0-9]{10,})\b", text)
        return m2.group(1).upper() if m2 else ""

    def _extract_price(self, text: str) -> Optional[int]:
        # Приоритет: отдельная строка с числом
        for ln in reversed([x.strip() for x in text.splitlines() if x.strip()]):
            if re.fullmatch(r"\d{2,8}", re.sub(r"\s+", "", ln)):
                return int(re.sub(r"\D", "", ln))
        m = re.search(r"\b(\d[\d\s]{1,10})\s*(?:руб\.?|р\.?|RUB)?\b", text, re.IGNORECASE)
        if m:
            return int(re.sub(r"\D", "", m.group(1)))
        return None

    def _extract_address(self, text: str) -> Tuple[str, str]:
        lines = [ln.strip(" ,") for ln in text.splitlines() if ln.strip()]
        addr = []
        start = None
        for i, ln in enumerate(lines):
            if re.search(r"\bг\.\s*[А-Яа-яA-Za-z\-]+", ln):
                start = i
                break
        if start is not None:
            for ln in lines[start : min(start + 6, len(lines))]:
                if re.search(r"^\d+\s*---", ln):
                    break
                if re.search(r"^[-]+$", ln):
                    break
                if re.fullmatch(r"[A-Z0-9\-]{8,}", ln):
                    break
                addr.append(ln)
        address = ", ".join(addr)
        m_city = re.search(r"\bг\.\s*([А-Яа-яA-Za-z\-]+)", address)
        city = m_city.group(1) if m_city else ""
        return address, city

    def _extract_attrs(self, text: str) -> Dict[str, str]:
        attrs: Dict[str, str] = {}
        work = text.replace("\\", "/")
        # Всё, что после второго --- обычно и есть блок характеристик
        parts = work.split("---", 2)
        if len(parts) == 3:
            work = parts[2]
        chunks = [c.strip() for c in re.split(r"\s*/\s*", work) if c.strip()]
        for ch in chunks:
            ch = re.sub(r"\s+", " ", ch).strip(" -")
            if not ch:
                continue
            # Явная форма Ключ-Значение
            if "-" in ch:
                key, val = ch.split("-", 1)
                key = key.strip(" -")
                val = val.strip(" -")
                # Если в ключе остался мусорный префикс до последнего пробела-серийника — чистим
                key = re.sub(r"^[A-Z0-9\-]{8,}\s+", "", key)
                key = re.sub(r"^\d+\s*", "", key)
                key = key.replace("---", "").strip()
                if len(key) >= 2 and len(val) >= 1:
                    attrs[key] = val
                continue
            # Отдельные порты без дефиса (например DisplayPort/Type-c)
            if re.search(r"displayport|type-c|usb|hdmi|dvi|vga", ch, re.IGNORECASE):
                if re.fullmatch(r"[A-Za-z0-9\-]{2,10}", ch) and "Видеовходы" in attrs:
                    attrs["Видеовходы"] = f"{attrs['Видеовходы']}/{ch}"
                    continue
                attrs.setdefault("Порты", "")
                attrs["Порты"] = ", ".join([x for x in [attrs["Порты"], ch] if x]).strip(", ")

        # Нормализация некоторых ключей
        renames = {
            "Сетевой шнур или блок питания в наличии": "ПитаниеВКомплекте",
            "Наличие сетевого шнура": "ПитаниеВКомплекте",
            "Затемнения и полосы на матрице": "ПроблемыМатрицы",
            "Видеовыходы": "Видеовыходы",
            "Видеовходы": "Видеовходы",
            "Тип питания": "ТипПитания",
            "Оперативная память": "ОЗУ",
            "Жесткий диск": "Накопитель",
            "Процессор": "Процессор",
            "Процессор частота": "ЧастотаПроцессора",
        }
        normalized: Dict[str, str] = {}
        for k, v in attrs.items():
            nk = renames.get(k, k)
            if nk in normalized and v not in normalized[nk]:
                normalized[nk] = f"{normalized[nk]}, {v}"
            else:
                normalized[nk] = v
        return normalized

    def _estimate_condition(self, text: str, matrix_issues: str = "", defect_1: str = "", defect_2: str = "") -> str:
        low = text.lower()
        issues_low = f"{matrix_issues} {defect_1} {defect_2}".lower()
        if any(x in low for x in ["не работает", "разбит", "не включается"]):
            return "Плохое"
        # Если явно написано что проблем с матрицей нет — не считаем это плохим состоянием
        if matrix_issues and re.search(r"\bнет\b", matrix_issues.lower()):
            pass
        elif any(x in issues_low for x in ["полос", "бит", "дефект", "трещ", "затемнен"]):
        for k, v in re.findall(r"([А-Яа-яA-Za-z0-9\-\s]+)-\s*([^/\n]+)", text):
            key = re.sub(r"\s+", " ", k).strip(" -")
            val = re.sub(r"\s+", " ", v).strip(" -")
            if len(key) < 2 or len(val) < 1:
                continue
            if re.search(r"^\d+\s*---", key):
                continue
            attrs[key] = val
        return attrs

    def _estimate_condition(self, text: str) -> str:
        low = text.lower()
        if any(x in low for x in ["не работает", "разбит", "полосы", "дефект", "не включается"]):
            return "Плохое"
        if any(x in low for x in ["следы эксплуатации", "потертости", "потёртости", "царапины"]):
            return "Удовлетворительное"
        return "Хорошее"

    def _quality_flags(self, row: Dict[str, object]) -> str:
        flags = []
        if not row["НомерЛота"]:
            flags.append("нет_номера")
        if not row["Бренд"]:
            flags.append("нет_бренда")
        if not row["Модель"]:
            flags.append("нет_модели")
        if not row["СерийныйНомер"]:
            flags.append("нет_serial")
        if not row["СтартоваяЦенаРуб"]:
            flags.append("нет_цены")
        return ",".join(flags)

    def _fill_ratio(self, row: Dict[str, object]) -> float:
        important = [
            "НомерЛота",
            "ГородАукциона",
            "ТипТехники",
            "Бренд",
            "Модель",
            "СерийныйНомер",
            "СтартоваяЦенаРуб",
            "АдресСамовывоза",
        ]
        present = sum(1 for k in important if str(row.get(k, "") or "").strip())
        return round(present / len(important), 2)

    def _is_meaningful(self, row: Dict[str, object]) -> bool:
        ratio = self._fill_ratio(row)
        row["ПолнотаЗаполнения"] = ratio
        row["ФлагиКачества"] = self._quality_flags(row)
        # фильтр мусора: номер + хотя бы 2 из [бренд/модель/цена/serial]
        score = 0
        for k in ["Бренд", "Модель", "СтартоваяЦенаРуб", "СерийныйНомер"]:
            if str(row.get(k, "") or "").strip():
                score += 1
        if not row.get("НомерЛота"):
            return False
        return score >= 2 and ratio >= self.options.quality_threshold

    def parse_lot_block(
        self,
        block: str,
        source_file: str,
        source_page: int,
        processed_at: datetime,
        header_city: str,
        header_dt: Optional[datetime],
    ) -> Optional[Dict[str, object]]:
        t = self.normalize_text(block)
        lot_no = self._extract_number(t)
        if lot_no is None:
            return None

        city_lot = ""
        m_city = re.search(r"\(([А-Яа-яA-Za-z\-\s]+)\)", t)
        if m_city:
            city_lot = m_city.group(1).strip()

        brand = self._extract_brand(t)
        model = self._extract_model(t, brand)
        attrs = self._extract_attrs(t)

        row: Dict[str, object] = {
            "ФайлИсточник": source_file,
            "СтраницаИсточник": source_page,
            "ДатаОбработки": processed_at,
            "НомерЛота": lot_no,
            "ГородАукциона": city_lot or header_city,
            "ДатаОкончанияТоргов": header_dt,
            "ТипТехники": self._extract_item_type(t),
            "Бренд": brand,
            "Модель": model,
            "Диагональ": self._extract_diagonal(t),
            "СерийныйНомер": self._extract_serial(t),
            "СтартоваяЦенаРуб": self._extract_price(t),
            "АдресСамовывоза": "",
            "ГородСамовывоза": "",
            "Видеовходы": attrs.get("Видеовходы", ""),
            "Видеовыходы": attrs.get("Видеовыходы", ""),
            "ТипПитания": attrs.get("ТипПитания", ""),
            "ПитаниеВКомплекте": attrs.get("ПитаниеВКомплекте", ""),
            "ПроблемыМатрицы": attrs.get("ПроблемыМатрицы", ""),
            "Дефект1": attrs.get("Дефект 1", ""),
            "Дефект2": attrs.get("Дефект 2", ""),
            "Процессор": attrs.get("Процессор", ""),
            "ЧастотаПроцессора": attrs.get("ЧастотаПроцессора", ""),
            "ОЗУ": attrs.get("ОЗУ", ""),
            "Накопитель": attrs.get("Накопитель", ""),
            "Порты": attrs.get("Порты", ""),
            "СостояниеОценка": "",
            "ТипПитания": attrs.get("Тип питания", ""),
            "ПитаниеВКомплекте": attrs.get("Сетевой шнур или блок питания в наличии", ""),
            "ПроблемыМатрицы": attrs.get("Затемнения и полосы на матрице", ""),
            "Дефект1": attrs.get("Дефект 1", ""),
            "Дефект2": attrs.get("Дефект 2", ""),
            "СостояниеОценка": self._estimate_condition(t),
            "КлючевыеСлова": "",
            "ПолнотаЗаполнения": 0.0,
            "ФлагиКачества": "",
            "ИзвлеченныеПараметрыJSON": "{}",
            "СыройТекст": t,
            "Примечание": "",
        }

        address, city_pickup = self._extract_address(t)
        row["АдресСамовывоза"] = address
        row["ГородСамовывоза"] = city_pickup

        extras = {}
        mapped = {
            "Видеовходы",
            "Видеовыходы",
            "ТипПитания",
            "ПитаниеВКомплекте",
            "ПроблемыМатрицы",
            "Дефект 1",
            "Дефект 2",
            "Диагональ",
            "Процессор",
            "ЧастотаПроцессора",
            "ОЗУ",
            "Накопитель",
            "Порты",
            "Тип питания",
            "Сетевой шнур или блок питания в наличии",
            "Затемнения и полосы на матрице",
            "Дефект 1",
            "Дефект 2",
            "Диагональ",
        }
        for k, v in attrs.items():
            if k not in mapped:
                extras[k] = v
        row["ИзвлеченныеПараметрыJSON"] = json.dumps(extras, ensure_ascii=False)

        row["СостояниеОценка"] = self._estimate_condition(
            t,
            matrix_issues=str(row.get("ПроблемыМатрицы", "")),
            defect_1=str(row.get("Дефект1", "")),
            defect_2=str(row.get("Дефект2", "")),
        )

        keys = [
            row["Бренд"],
            row["Модель"],
            row["ТипТехники"],
            row["Процессор"],
            row["ОЗУ"],
            row["Накопитель"],
            row["ПроблемыМатрицы"],
            row["Дефект1"],
            row["Дефект2"],
        ]
        keys = [row["Бренд"], row["Модель"], row["ТипТехники"], row["ПроблемыМатрицы"], row["Дефект1"], row["Дефект2"]]
        row["КлючевыеСлова"] = ", ".join([str(k) for k in keys if str(k).strip()])[:250]

        if not self._is_meaningful(row):
            self.errors.append(ParseError(source_file, source_page, lot_no, "Отфильтрована неинформативная строка", t))
            return None

        if row["СтартоваяЦенаРуб"] is None:
            self.errors.append(ParseError(source_file, source_page, lot_no, "Не найдена стартовая цена", t))

        return row

    def parse_pdf(self, pdf_path: Path) -> List[Dict[str, object]]:
        rows: List[Dict[str, object]] = []
        try:
            doc = fitz.open(pdf_path)
        except Exception as exc:
            self.errors.append(ParseError(pdf_path.name, 0, None, f"Ошибка открытия PDF: {exc}", ""))
            return rows

        processed_at = datetime.now()
        header_city = ""
        header_dt: Optional[datetime] = None

        for i in range(len(doc)):
            page = doc[i]
            text = page.get_text("text") or ""
            if self.options.ocr_mode == "Always" or (self.options.ocr_mode == "Auto" and self._need_ocr(text)):
                ocr = self._ocr_page(page)
                if ocr.strip():
                    text = text + "\n" + ocr

            if i == 0:
                header_city, header_dt = self._extract_auction_meta(text)

            blocks = self.split_into_lots(text)
            for b in blocks:
                row = self.parse_lot_block(b, pdf_path.name, i + 1, processed_at, header_city, header_dt)
                if row:
                    rows.append(row)

        doc.close()
        return rows

    def parse_inputs(self, input_paths: Sequence[Path]) -> List[Dict[str, object]]:
        pdfs: List[Path] = []
        for p in input_paths:
            if p.is_dir():
                pdfs.extend(sorted(p.glob("*.pdf")))
            elif p.is_file() and p.suffix.lower() == ".pdf":
                pdfs.append(p)

        rows: List[Dict[str, object]] = []
        for pdf in pdfs:
            self.logger.info("Обрабатывается: %s", pdf)
            rows.extend(self.parse_pdf(pdf))
        return rows

    def save_errors(self, out_dir: Path) -> Path:
        out_dir.mkdir(parents=True, exist_ok=True)
        err = out_dir / DEFAULT_ERROR_FILE
        with open(err, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["ФайлИсточник", "СтраницаИсточник", "НомерЛота", "Сообщение", "СыройТекст"])
            for e in self.errors:
                w.writerow([e.source_file, e.source_page, e.lot_position, e.message, e.raw_text])
        return err


def export_excel(rows: List[Dict[str, object]], out_xlsx: Path, save_csv: bool = False) -> Path:
    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)
    for c in BASE_COLUMNS_RU:
        if c not in df.columns:
            df[c] = ""
    df = df[BASE_COLUMNS_RU]

    df["СтартоваяЦенаРуб"] = pd.to_numeric(df["СтартоваяЦенаРуб"], errors="coerce")
    df["ПолнотаЗаполнения"] = pd.to_numeric(df["ПолнотаЗаполнения"], errors="coerce")

    from openpyxl.styles import Font, PatternFill
    with pd.ExcelWriter(out_xlsx, engine="openpyxl", datetime_format="YYYY-MM-DD HH:MM:SS") as writer:
        df.to_excel(writer, index=False, sheet_name="Лоты")
        ws = writer.sheets["Лоты"]
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill(start_color="FFEAF2FF", end_color="FFEAF2FF", fill_type="solid")
        for col in ws.columns:
            width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(12, width + 2), 70)

    if save_csv:
        df.to_csv(out_xlsx.with_suffix(".csv"), index=False, encoding="utf-8-sig")

    return out_xlsx


def build_logger(name: str, level: str = "INFO", extra_handler: Optional[logging.Handler] = None) -> logging.Logger:
    logger = logging.getLogger(name)
    logger.handlers.clear()
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

    if extra_handler:
        extra_handler.setFormatter(fmt)
        logger.addHandler(extra_handler)

    return logger


class QtLogHandler(logging.Handler):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def emit(self, record: logging.LogRecord) -> None:
        self.callback(self.format(record))


class DropArea(QtWidgets.QFrame):
    filesDropped = QtCore.pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self._style(False)
        lay = QtWidgets.QVBoxLayout(self)
        lbl = QtWidgets.QLabel("Перетащите PDF-файлы сюда")
        lbl.setAlignment(QtCore.Qt.AlignCenter)
        lbl.setStyleSheet("font-size: 16px; color: #555;")
        lay.addWidget(lbl)

    def _style(self, active: bool):
        if active:
            self.setStyleSheet("QFrame {border: 2px dashed #2e7d32; border-radius: 8px; background: #e8f5e9;}")
        else:
            self.setStyleSheet("QFrame {border: 2px dashed #9aa0a6; border-radius: 8px; background: #f7f7f7;}")

    def dragEnterEvent(self, event: QtGui.QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self._style(True)

    def dragLeaveEvent(self, event: QtGui.QDragLeaveEvent) -> None:
        self._style(False)

    def dropEvent(self, event: QtGui.QDropEvent) -> None:
        self._style(False)
        paths = []
        for url in event.mimeData().urls():
            p = Path(url.toLocalFile())
            if p.exists():
                paths.append(str(p))
        self.filesDropped.emit(paths)


class OptionsDialog(QtWidgets.QDialog):
    def __init__(self, options: ParserOptions, parent=None):
        super().__init__(parent)
        self.options = options
        self.setWindowTitle("Опции")
        form = QtWidgets.QFormLayout(self)

        self.ocr_mode = QtWidgets.QComboBox()
        self.ocr_mode.addItems(["Auto", "Always", "Never"])
        self.ocr_mode.setCurrentText(options.ocr_mode)

        self.out_folder = QtWidgets.QLineEdit(str(options.output_folder))
        self.auto_drop = QtWidgets.QCheckBox("Авто-старт при drop")
        self.auto_drop.setChecked(options.auto_start_drop)

        self.log_level = QtWidgets.QComboBox()
        self.log_level.addItems(["DEBUG", "INFO", "WARNING", "ERROR"])
        self.log_level.setCurrentText(options.log_level)

        self.save_csv = QtWidgets.QCheckBox("Сохранять дополнительный CSV")
        self.save_csv.setChecked(options.save_csv)

        self.quality = QtWidgets.QDoubleSpinBox()
        self.quality.setRange(0.1, 1.0)
        self.quality.setSingleStep(0.05)
        self.quality.setValue(options.quality_threshold)

        form.addRow("OCR mode", self.ocr_mode)
        form.addRow("Папка вывода", self.out_folder)
        form.addRow("", self.auto_drop)
        form.addRow("Уровень логирования", self.log_level)
        form.addRow("Минимальная полнота строки", self.quality)
        form.addRow("", self.save_csv)

        bb = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        form.addWidget(bb)

    def accept(self) -> None:
        self.options.ocr_mode = self.ocr_mode.currentText()
        self.options.output_folder = Path(self.out_folder.text().strip() or DEFAULT_OUTPUT_DIR)
        self.options.auto_start_drop = self.auto_drop.isChecked()
        self.options.log_level = self.log_level.currentText()
        self.options.save_csv = self.save_csv.isChecked()
        self.options.quality_threshold = self.quality.value()
        super().accept()


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(700, 450)

        self.options = ParserOptions()
        self.selected_paths: List[Path] = []
        self.rows: List[Dict[str, object]] = []
        self.error_file: Optional[Path] = None

        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        v = QtWidgets.QVBoxLayout(central)

        self.drop = DropArea()
        self.drop.filesDropped.connect(self.on_drop)
        v.addWidget(self.drop)

        h = QtWidgets.QHBoxLayout()
        self.btn_select = QtWidgets.QPushButton("Выбрать файл(ы)…")
        self.btn_process = QtWidgets.QPushButton("Обработать")
        self.btn_save = QtWidgets.QPushButton("Сохранить Excel")
        self.btn_opt = QtWidgets.QPushButton("Опции")
        self.btn_tests = QtWidgets.QPushButton("Run sample tests")

        self.btn_select.clicked.connect(self.select_files)
        self.btn_process.clicked.connect(self.process)
        self.btn_save.clicked.connect(self.save_excel)
        self.btn_opt.clicked.connect(self.show_options)
        self.btn_tests.clicked.connect(self.run_tests)

        for b in [self.btn_select, self.btn_process, self.btn_save, self.btn_opt, self.btn_tests]:
            h.addWidget(b)
        v.addLayout(h)

        self.progress = QtWidgets.QProgressBar()
        self.progress.setRange(0, 100)
        v.addWidget(self.progress)

        self.log = QtWidgets.QPlainTextEdit()
        self.log.setReadOnly(True)
        v.addWidget(self.log)

        self.qt_handler = QtLogHandler(self.append_log)
        self.logger = build_logger("gui", self.options.log_level, self.qt_handler)

    def append_log(self, msg: str):
        self.log.appendPlainText(msg)

    def select_files(self):
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Выберите PDF", "", "PDF files (*.pdf)")
        if files:
            self.selected_paths = [Path(x) for x in files]
            self.logger.info("Выбрано файлов: %s", len(files))

    def on_drop(self, files: List[str]):
        self.selected_paths = [Path(x) for x in files]
        self.logger.info("Получено через drag&drop: %s", len(files))
        if self.options.auto_start_drop:
            self.process()

    def show_options(self):
        dlg = OptionsDialog(self.options, self)
        if dlg.exec_():
            self.logger = build_logger("gui", self.options.log_level, self.qt_handler)
            self.logger.info("Опции обновлены")

    def process(self):
        if not self.selected_paths:
            QtWidgets.QMessageBox.warning(self, "Нет файлов", "Выберите PDF-файлы или папку.")
            return
        try:
            parser = LotParser(self.options, self.logger)
            self.progress.setValue(15)
            self.rows = parser.parse_inputs(self.selected_paths)
            self.progress.setValue(85)
            self.error_file = parser.save_errors(self.options.output_folder)
            self.progress.setValue(100)
            self._final_dialog(len(self.rows), len(parser.errors))
        except Exception as exc:
            self.logger.error("Ошибка: %s", exc)
            self.logger.debug(traceback.format_exc())
            QtWidgets.QMessageBox.critical(self, "Ошибка", str(exc))

    def save_excel(self):
        if not self.rows:
            QtWidgets.QMessageBox.warning(self, "Нет данных", "Сначала нажмите «Обработать».")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel",
            str(self.options.output_folder / "lots_clean.xlsx"),
            "Excel (*.xlsx)",
        )
        if not path:
            return
        out = export_excel(self.rows, Path(path), save_csv=self.options.save_csv)
        self.logger.info("Сохранено: %s", out)
        QtWidgets.QMessageBox.information(self, "Готово", f"Excel сохранён:\n{out}")

    def _final_dialog(self, lots: int, errs: int):
        box = QtWidgets.QMessageBox(self)
        box.setWindowTitle("Обработка завершена")
        box.setText(f"Обработано лотов: {lots}\nОшибок/отфильтрованных: {errs}")
        b1 = box.addButton("Открыть папку результатов", QtWidgets.QMessageBox.ActionRole)
        b2 = box.addButton("Открыть parsing_errors.csv", QtWidgets.QMessageBox.ActionRole)
        box.addButton("OK", QtWidgets.QMessageBox.AcceptRole)
        box.exec_()
        clicked = box.clickedButton()
        if clicked == b1:
            open_path(self.options.output_folder)
        if clicked == b2 and self.error_file:
            open_path(self.error_file)

    def run_tests(self):
        report = run_sample_tests(logger=self.logger)
        QtWidgets.QMessageBox.information(self, "Тесты", report)


def open_path(path: Path):
    try:
        if os.name == "nt":
            subprocess.Popen(["explorer", str(path)])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


def collect_inputs(items: Sequence[str]) -> List[Path]:
    out = []
    for i in items:
        p = Path(i)
        if p.exists():
            out.append(p)
    return out


def run_cli(args: argparse.Namespace) -> int:
    options = ParserOptions(
        ocr_mode=args.ocr_mode,
        output_folder=Path(args.output_dir),
        auto_start_drop=False,
        log_level=args.log_level,
        save_csv=args.save_csv,
        quality_threshold=args.quality_threshold,
    )
    logger = build_logger("cli", options.log_level)
    parser = LotParser(options, logger)
    inputs = collect_inputs(args.input)
    if not inputs:
        print("Не найдены входные файлы/папки")
        return 2

    rows = parser.parse_inputs(inputs)
    out = export_excel(rows, Path(args.output), save_csv=options.save_csv)
    err = parser.save_errors(options.output_folder)

    print(f"Готово. Лотов: {len(rows)}; ошибок/фильтров: {len(parser.errors)}")
    print(f"Excel: {out}")
    print(f"Ошибки: {err}")
    return 0 if rows else 1


def run_sample_tests(logger: Optional[logging.Logger] = None) -> str:
    logger = logger or build_logger("tests", "INFO")
    opts = ParserOptions(ocr_mode="Never", output_folder=Path("output_test"), quality_threshold=0.3)
    p = LotParser(opts, logger)

    sample = (
        "1--- (Москва) Монитор AOC M2060SWD2 19,5' s/n: GLXK5HA053007 --- Диагональ-19,5' /\n"
        "Затемнения и полосы на матрице-Нет / Видеовходы-VGA/DVI / Тип питания-Кабель /\n"
        "Сетевой шнур или блок питания в наличии - Да/Следы эксплуатации.\n"
        "700\nг. Москва,\nшоссе\nЭнтузиастов\nд.14\nGLXK5HA053007"
    )

    row = p.parse_lot_block(
        sample,
        source_file="sample.pdf",
        source_page=1,
        processed_at=datetime.now(),
        header_city="Москва",
        header_dt=datetime(2026, 3, 16, 12, 0),
    )

    checks = {
        "номер": bool(row and row["НомерЛота"] == 1),
        "бренд": bool(row and row["Бренд"] == "AOC"),
        "модель": bool(row and row["Модель"] == "M2060SWD2"),
        "диагональ": bool(row and row["Диагональ"] == '19.5"'),
        "serial": bool(row and row["СерийныйНомер"] == "GLXK5HA053007"),
        "цена": bool(row and row["СтартоваяЦенаРуб"] == 700),
    }

    pc_sample = (
        "11--- (Москва) ПК APPLE Mac mini s/n: C02XXX111 --- Оперативная память-8GB / "
        "Процессор-Apple M1 / Жесткий диск-256GB SSD / Видеовыходы-HDMI/DisplayPort/Type-c / "
        "Наличие сетевого шнура-да / Следы эксплуатации.\n50000\nг. Москва, ул. Пример, д.1"
    )
    pc_row = p.parse_lot_block(
        pc_sample,
        source_file="pc.pdf",
        source_page=1,
        processed_at=datetime.now(),
        header_city="Москва",
        header_dt=datetime(2026, 3, 16, 12, 0),
    )
    checks.update(
        {
            "apple_brand": bool(pc_row and pc_row["Бренд"] == "APPLE"),
            "процессор": bool(pc_row and "M1" in str(pc_row["Процессор"])),
            "озу": bool(pc_row and "8GB" in str(pc_row["ОЗУ"])),
            "накопитель": bool(pc_row and "256GB" in str(pc_row["Накопитель"])),
        }
    )

    called = {"ok": False}
    real = p._ocr_page

    def fake_ocr(page):
        called["ok"] = True
        return "ocr text"

    p._ocr_page = fake_ocr  # type: ignore[assignment]
    if p._need_ocr(""):
        p._ocr_page(None)  # type: ignore[arg-type]
    p._ocr_page = real  # type: ignore[assignment]
    checks["ocr_fallback"] = called["ok"]

    passed = sum(1 for v in checks.values() if v)
    report = "\n".join([f"{k}: {'PASS' if v else 'FAIL'}" for k, v in checks.items()] + [f"Итого: {passed}/{len(checks)}"])
    logger.info("Результаты sample tests:\n%s", report)
    return report


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Парсер лотов PDF")
    ap.add_argument("--no-gui", action="store_true", help="Запуск в CLI")
    ap.add_argument("--input", nargs="+", default=[], help="Пути к PDF или папкам")
    ap.add_argument("--output", default=str(Path(DEFAULT_OUTPUT_DIR) / "lots_clean.xlsx"), help="Выходной Excel")
    ap.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR, help="Папка служебных файлов")
    ap.add_argument("--ocr-mode", choices=["Auto", "Always", "Never"], default="Auto")
    ap.add_argument("--save-csv", action="store_true", help="Сохранить CSV")
    ap.add_argument("--log-level", default="INFO")
    ap.add_argument("--quality-threshold", type=float, default=0.35)
    ap.add_argument("--run-sample-tests", action="store_true")
    return ap.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)

    if args.run_sample_tests:
        print(run_sample_tests())
        return 0

    if args.no_gui:
        return run_cli(args)

    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    w = MainWindow()
    w.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
