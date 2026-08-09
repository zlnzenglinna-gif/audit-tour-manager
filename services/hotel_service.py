from __future__ import annotations

import re
from pathlib import Path
from typing import BinaryIO

import pandas as pd
from rapidfuzz import fuzz


COMPANY_SUFFIX_PATTERNS = [
    r"\bgmbh\b",
    r"\bmbh\b",
    r"\bkgaa\b",
    r"\bag\b",
    r"\bkg\b",
    r"\bs\.?\s*r\.?\s*o\.?\b",
    r"\bd\.?\s*o\.?\s*o\.?\b",
    r"\bkft\b",
    r"\bltd\b",
    r"\blimited\b",
    r"\binc\b",
    r"\bllc\b",
    r"\bsa\b",
]


def read_hotel_overview(
    file_source: str | Path | BinaryIO,
) -> pd.DataFrame:
    """
    读取 Hotelübersicht，并清理基础数据。
    """

    hotel_df = pd.read_excel(
        file_source,
        engine="openpyxl",
    )

    hotel_df.columns = [
        str(column).strip()
        for column in hotel_df.columns
    ]

    required_columns = [
        "Kunde",
        "Hotel",
        "Tripadvisor Link",
        "Kommentar",
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in hotel_df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Hotelübersicht缺少字段："
            + ", ".join(missing_columns)
        )

    hotel_df = hotel_df.dropna(
        how="all",
        subset=["Kunde", "Hotel"],
    ).copy()

    text_columns = [
        "Kunde",
        "Hotel",
        "Tripadvisor Link",
        "Kommentar",
    ]

    for column in text_columns:
        hotel_df[column] = (
            hotel_df[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    return hotel_df.reset_index(drop=True)


def normalize_location_name(value: object) -> str:
    """
    清理工厂/地点名称。

    例如：
    Grupo Antolin Bratislava s.r.o.
    -> grupo antolin bratislava

    BMW Werk GmbH
    -> bmw werk
    """

    if value is None:
        return ""

    text = str(value).strip().lower()

    for pattern in COMPANY_SUFFIX_PATTERNS:
        text = re.sub(
            pattern,
            " ",
            text,
            flags=re.IGNORECASE,
        )

    text = re.sub(
        r"[.,;:()/_-]+",
        " ",
        text,
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def find_historical_hotels(
    location_name: str,
    hotel_df: pd.DataFrame,
    minimum_score: int = 70,
) -> pd.DataFrame:
    """
    根据 Tourplan 工厂名称，返回所有相似度达到
    minimum_score 的 Kunde 对应酒店。

    同时加入 Match Score，方便开发阶段检查匹配质量。
    """

    if not location_name:
        return hotel_df.iloc[0:0].copy()

    normalized_location = normalize_location_name(
        location_name
    )

    if not normalized_location:
        return hotel_df.iloc[0:0].copy()

    working_df = hotel_df.copy()

    working_df["_normalized_kunde"] = (
        working_df["Kunde"]
        .fillna("")
        .apply(normalize_location_name)
    )

    working_df["Match Score"] = (
        working_df["_normalized_kunde"]
        .apply(
            lambda kunde_name: fuzz.token_set_ratio(
                normalized_location,
                kunde_name,
            )
            if kunde_name
            else 0
        )
    )

    matched_df = working_df[
        working_df["Match Score"] >= minimum_score
    ].copy()

    if matched_df.empty:
        return hotel_df.iloc[0:0].copy()

    matched_df = matched_df.sort_values(
        by=[
            "Match Score",
            "Kunde",
            "Hotel",
        ],
        ascending=[
            False,
            True,
            True,
        ],
    )

    matched_df = matched_df.drop(
        columns=["_normalized_kunde"]
    )

    return matched_df.reset_index(drop=True)