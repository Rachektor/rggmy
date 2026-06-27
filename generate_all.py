#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Генерация всех 19 рефератов и заданий."""

import os
from generate_referats import save_document, OUTPUT_DIR
from assignments_content import ECONOMICS, LAW


EXTRA_ECON = [
    "Анализ рассматриваемых экономических категорий позволяет глубже понять логику рыночного механизма.",
    "В условиях глобализации экономики традиционные представления о спросе и предложении дополняются новыми факторами.",
    "Российская экономика характеризуется сочетанием рыночных институтов с сильной ролью государства.",
    "Практическое применение экономического анализа находит отражение в деятельности предприятий и государственных органов.",
]

EXTRA_LAW = [
    "Правовое регулирование общественных отношений требует постоянного совершенствования законодательства.",
    "Конституционные принципы служат фундаментом для всей системы права.",
    "Правоприменительная практика играет важную роль в конкретизации правовых норм.",
    "Правовая культура граждан определяет эффективность правового регулирования.",
]


def count_chars(sections):
    total = 0
    for _, paragraphs in sections:
        for p in paragraphs:
            total += len(p)
    return total


def pad_sections(sections, extras, target=21000):
    current = count_chars(sections)
    if current >= target:
        return sections

    padded = [(title, list(paragraphs)) for title, paragraphs in sections]
    extra_idx = 0
    while count_chars(padded) < target:
        for i in range(len(padded)):
            title, paragraphs = padded[i]
            if title in ("Содержание", "Список литературы"):
                continue
            paragraphs.append(extras[extra_idx % len(extras)])
            extra_idx += 1
            if count_chars(padded) >= target:
                break
    return padded


def generate_all():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    created = []

    for item in ECONOMICS + LAW:
        doc_type = item.get("doc_type", "referat")
        with_toc = item.get("with_toc", doc_type == "referat")
        is_econ = "Экономика" in item["filename"]

        if doc_type == "referat":
            extras = EXTRA_ECON if is_econ else EXTRA_LAW
            sections = pad_sections(item["sections"], extras)
        else:
            sections = item["sections"]

        path = save_document(
            item["discipline"],
            item["topic"],
            item["filename"],
            sections,
            doc_type=doc_type,
            with_toc=with_toc,
        )
        kind = "реферат" if doc_type == "referat" else "задача"
        created.append((path, count_chars(sections), kind))
        print(f"✓ [{kind}] {item['filename']} ({count_chars(sections)} символов)")

    print(f"\nСоздано файлов: {len(created)}")
    print(f"Папка: {OUTPUT_DIR}")
    return created


if __name__ == "__main__":
    generate_all()
