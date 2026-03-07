from postprocess_lots import postprocess_rows


def test_filters_noise_and_deduplicates_and_drops_photo_columns():
    rows = [
        {
            "SourceFile": "a.pdf",
            "SourcePage": "1",
            "LotPosition": "",
            "AuctionCity": "самовывоз",
            "RawText": "Позиция лота Наименование оборудования Стартовая цена Фото 1 Фото 2",
            "Photo1": "img1.jpg",
        },
        {
            "SourceFile": "a.pdf",
            "SourcePage": "1",
            "LotPosition": "1",
            "AuctionCity": "Москва",
            "ItemType": "Монитор",
            "Brand": "AOC",
            "Model": "M2060SWD2",
            "SerialNumber": "GLXK5HA053007",
            "StartPriceRUB": "700",
            "RawText": "1 --- (Москва) Монитор AOC M2060SWD2 19,5' s/n: GLXK5HA053007",
            "Photo1": "img1.jpg",
        },
        {
            "SourceFile": "a.pdf",
            "SourcePage": "1",
            "LotPosition": "1",
            "AuctionCity": "",
            "ItemType": "",
            "Brand": "",
            "Model": "",
            "SerialNumber": "",
            "StartPriceRUB": "",
            "RawText": "1 --- (Москва) Монитор AOC M2060SWD2 19,5' s/n: GLXK5HA053007",
            "Photo1": "img1.jpg",
        },
    ]

    out = postprocess_rows(rows)

    assert len(out) == 1
    assert "Photo1" not in out[0]
    assert out[0]["Позиция_лота"] == "1"
    assert out[0]["Наличие_серийного"] == "Да"
    assert int(out[0]["Полнота_заполнения_проц"]) >= 70


def test_fills_fields_from_raw_text():
    rows = [
        {
            "SourceFile": "b.pdf",
            "SourcePage": "2",
            "LotPosition": "",
            "AuctionCity": "",
            "ItemType": "",
            "Brand": "",
            "Model": "",
            "SerialNumber": "",
            "StartPriceRUB": "",
            "RawText": "8 --- (Москва) Монитор AOC M2060SWD2 19,5' s/n: GLXK5HA054308 700 руб",
        }
    ]

    out = postprocess_rows(rows)
    row = out[0]

    assert row["Позиция_лота"] == "8"
    assert row["Город_торгов"] == "Москва"
    assert row["Бренд"] == "AOC"
    assert row["Модель"] == "M2060SWD2"
    assert row["Серийный_номер"] == "GLXK5HA054308"
