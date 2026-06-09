#!/usr/bin/env python3
"""
ESG-рейтинг золотодобычи. Калькулятор методологии v3 (disclosure-based).

Использование:
    python calculate_rating.py input.json > output.json
    python calculate_rating.py input.json --pretty
    python calculate_rating.py input_year1.json input_year2.json --dynamic

Формат input.json — см. INPUT_SCHEMA ниже.
Формат output.json — полный паспорт расчёта.
"""
import json
import sys
import argparse
from typing import Any

# ============================================================================
# КАТАЛОГ ПОКАЗАТЕЛЕЙ (веса и метаданные)
# ============================================================================

INDICATORS = {
    # --- БЛОК E (сумма = 333) ---
    "E1":  {"name": "Удельные выбросы ПГ (Scope 1+2)", "weight": 54, "block": "E", "type": "minimizer_quant"},
    "E2":  {"name": "Полнота раскрытия Scope 3",        "weight": 18, "block": "E", "type": "maximizer_fraction"},
    "E3":  {"name": "Энергоинтенсивность",              "weight": 30, "block": "E", "type": "minimizer_quant"},
    "E4":  {"name": "Доля ВИЭ",                          "weight": 25, "block": "E", "type": "maximizer_fraction"},
    "E5":  {"name": "Удельное водопотребление + стресс","weight": 30, "block": "E", "type": "composite"},
    "E6":  {"name": "Управление TSF (GISTM)",           "weight": 60, "block": "E", "type": "maturity_0_4"},
    "E7":  {"name": "Опасные вещества (цианид/ртуть)",  "weight": 35, "block": "E", "type": "maturity_0_4"},
    "E8":  {"name": "Биоразнообразие",                  "weight": 30, "block": "E", "type": "composite"},
    "E9":  {"name": "Рекультивация / ARO",              "weight": 25, "block": "E", "type": "maturity_0_4"},
    "E10": {"name": "Экологические штрафы",             "weight": 16, "block": "E", "type": "minimizer_quant"},
    "E11": {"name": "ISO 14001",                         "weight": 10, "block": "E", "type": "binary_with_coverage"},
    # --- БЛОК S (сумма = 333) ---
    "S1":  {"name": "LTIFR",                             "weight": 55, "block": "S", "type": "minimizer_quant"},
    "S2":  {"name": "Fatalities",                        "weight": 60, "block": "S", "type": "minimizer_quant_capped"},
    "S3":  {"name": "ISO 45001",                         "weight": 18, "block": "S", "type": "binary_with_coverage"},
    "S4":  {"name": "Проф. заболевания",                 "weight": 30, "block": "S", "type": "minimizer_quant"},
    "S5":  {"name": "Текучесть персонала",               "weight": 20, "block": "S", "type": "minimizer_quant"},
    "S6":  {"name": "Обучение персонала",                "weight": 15, "block": "S", "type": "maximizer_quant"},
    "S7":  {"name": "Права КМНС / соглашения",           "weight": 55, "block": "S", "type": "maturity_0_4_capped"},
    "S8":  {"name": "Due diligence поставщиков",         "weight": 20, "block": "S", "type": "maximizer_fraction"},
    "S9":  {"name": "Механизм обращений",                "weight": 15, "block": "S", "type": "maximizer_fraction"},
    "S10": {"name": "Социальные инвестиции",             "weight": 45, "block": "S", "type": "maximizer_quant"},
    # --- БЛОК G (сумма = 333) ---
    "G1":  {"name": "Независимые директора",             "weight": 33, "block": "G", "type": "fraction_by_thresholds"},
    "G2":  {"name": "Разделение председатель/CEO",       "weight": 12, "block": "G", "type": "binary"},
    "G3":  {"name": "Антикоррупция (ISO 37001)",         "weight": 22, "block": "G", "type": "three_levels"},
    "G4":  {"name": "Случаи коррупции",                  "weight": 40, "block": "G", "type": "minimizer_quant_capped"},
    "G5":  {"name": "Надзор СД за ESG",                  "weight": 40, "block": "G", "type": "maturity_0_4"},
    "G6":  {"name": "Управление рисками",                "weight": 30, "block": "G", "type": "maturity_0_4"},
    "G7":  {"name": "Whistleblower канал",               "weight": 25, "block": "G", "type": "composite"},
    "G8":  {"name": "LBMA/WGC/RJC/ICMM",                 "weight": 45, "block": "G", "type": "maturity_0_4"},
    "G9":  {"name": "Запасы (JORC)",                     "weight": 30, "block": "G", "type": "maturity_0_4"},
    "G10": {"name": "Налоговая прозрачность / EITI",     "weight": 25, "block": "G", "type": "maturity_0_4"},
    "G11": {"name": "Качество раскрытия ESG",            "weight": 22, "block": "G", "type": "maturity_0_4"},
    "G12": {"name": "Финансовая прозрачность (АКРА)",    "weight":  9, "block": "G", "type": "acra_composite"},
}

# Проверка суммы весов
assert sum(v["weight"] for v in INDICATORS.values() if v["block"] == "E") == 333
assert sum(v["weight"] for v in INDICATORS.values() if v["block"] == "S") == 333
assert sum(v["weight"] for v in INDICATORS.values() if v["block"] == "G") == 333

# ============================================================================
# ABSOLUTE THRESHOLDS (fallback по ICMM benchmark)
# ============================================================================

ABSOLUTE_THRESHOLDS = {
    "E1": [(0.3, 1.00), (0.5, 0.80), (0.7, 0.60), (0.9, 0.40), (1.2, 0.20), (float("inf"), 0.00)],
    "E3": [(4, 1.00), (6, 0.80), (9, 0.60), (12, 0.40), (15, 0.20), (float("inf"), 0.00)],
    "S1": [(0.10, 1.00), (0.25, 0.85), (0.50, 0.70), (1.00, 0.55), (2.00, 0.40), (4.00, 0.20), (float("inf"), 0.00)],
    "S2_fatalities_abs": {0: 1.00, 1: 0.50, 2: 0.40, 3: 0.30, 4: 0.20, 5: 0.10},  # special key=abs count
    "S4": [(0.1, 1.00), (0.3, 0.80), (0.7, 0.60), (1.5, 0.40), (3.0, 0.20), (float("inf"), 0.00)],
    "S5": [(0.08, 1.00), (0.12, 0.80), (0.18, 0.60), (0.25, 0.40), (0.35, 0.20), (float("inf"), 0.00)],
    "S6_max": [(100, 1.00), (70, 0.80), (45, 0.60), (25, 0.40), (10, 0.20), (0, 0.00)],  # maximizer: higher = better
    "S10_max": [(0.02, 1.00), (0.01, 0.80), (0.005, 0.60), (0.002, 0.40), (0.0005, 0.20), (0, 0.00)],
    "E10": [(0, 1.00), (0.5, 0.80), (2, 0.60), (5, 0.40), (10, 0.20), (float("inf"), 0.00)],
}

RATING_SCALE = [
    (849, 999, "AAA (ESG-1)", "Лидер отрасли"),
    (749, 848, "AA (ESG-2)",  "Высокий уровень соответствия"),
    (649, 748, "A (ESG-3)",   "Уверенно выше среднего"),
    (549, 648, "BBB (ESG-4)", "Средний уровень"),
    (449, 548, "BB (ESG-5)",  "Ниже среднего"),
    (299, 448, "B (ESG-6)",   "Низкий уровень"),
    (0,   298, "C (ESG-7)",   "Критически низкий"),
]


# ============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================================

def apply_thresholds_minimizer(value: float, thresholds: list) -> float:
    """Для минимайзера: порог задаёт ВЕРХНЮЮ границу значения для данного x. Меньше = лучше."""
    for upper, x in thresholds:
        if value <= upper:
            return x
    return 0.0


def apply_thresholds_maximizer(value: float, thresholds: list) -> float:
    """Для максимайзера: порог задаёт НИЖНЮЮ границу значения для данного x. Больше = лучше."""
    for lower, x in thresholds:
        if value >= lower:
            return x
    return 0.0


def percentile_minimizer(value: float, peer_values: list) -> float:
    """Для минимайзера: x = 1 - percentile_rank. Наименьшее значение = x=1."""
    if not peer_values:
        return None
    all_vals = sorted(peer_values + [value])
    rank = all_vals.index(value) / max(1, (len(all_vals) - 1))
    return 1 - rank


def percentile_maximizer(value: float, peer_values: list) -> float:
    """Для максимайзера: x = percentile_rank. Наибольшее значение = x=1."""
    if not peer_values:
        return None
    all_vals = sorted(peer_values + [value])
    rank = all_vals.index(value) / max(1, (len(all_vals) - 1))
    return rank


def clip(x: float, lo=0.0, hi=1.0) -> float:
    return max(lo, min(hi, x))


# ============================================================================
# РАСЧЁТ ОТДЕЛЬНЫХ ПОКАЗАТЕЛЕЙ
# ============================================================================

def calculate_indicator(code: str, data: dict, peer_set: dict) -> dict:
    """
    Рассчитывает один показатель.
    data — словарь с ключами: value (или level), disclosed (bool), caveats (list[str]),
           source_page (int), raw_quote (str), not_applicable (bool).
    peer_set — словарь с peer-данными: {code: [values]}.
    Возвращает: {x, score, explanation, caveats_applied}.
    """
    meta = INDICATORS[code]
    weight = meta["weight"]
    ind_type = meta["type"]

    # Если показатель помечен как не применимый — возвращаем None (обработается снаружи)
    if data.get("not_applicable", False):
        return {
            "code": code, "name": meta["name"], "weight": weight,
            "x": None, "score": None, "applicable": False,
            "explanation": data.get("na_reason", "Помечен как не применимый"),
            "source_page": None, "caveats_applied": []
        }

    # Если ничего не раскрыто — x = 0
    if not data.get("disclosed", False):
        # Особый случай: если по G8/G9/G10 заявлен sanctions_barrier при disclosed=false,
        # явно фиксируем это в результате для академической прозрачности.
        # Бонус всё равно НЕ начисляется (нет фактической практики), но запись попадёт
        # в раздел 3.2 Word-отчёта «Санкционные флаги без начисления бонуса».
        sanctions_note = None
        caveats = ["not_disclosed"]
        if code in ("G8", "G9", "G10") and data.get("sanctions_barrier", False):
            sanctions_note = {
                "triggered": False,
                "x_raw": 0.0,
                "x_final": 0.0,
                "limit_used": "—",
                "bonus_amount": 0.0,
                "trigger_reason": data.get("sanctions_reason", "Санкционный барьер заявлен"),
                "non_application_reason": (
                    "Санкционный бонус не начислен: показатель не раскрыт в отчёте, "
                    "поэтому отсутствует фактическая практика для амплификации. "
                    "Anti-washing защита: бонус не «вытаскивает из ноля» компанию без реальных практик."
                ),
            }
            caveats.append("sanctions_bonus_not_applied_no_disclosure")
        return {
            "code": code, "name": meta["name"], "weight": weight,
            "x": 0.0, "score": 0.0, "applicable": True,
            "explanation": "Не раскрыто в отчёте → x = 0 (disclosure-based).",
            "source_page": None,
            "pdf_file": None,
            "pdf_page_ref": None,
            "raw_quote": None,
            "caveats_applied": caveats,
            "sanctions_bonus": sanctions_note,
        }

    caveats_applied = []
    sanctions_applied_note = None
    x = None
    explanation = ""

    # ---------- По типам ----------
    if ind_type == "maturity_0_4":
        level = data.get("level", 0)
        x = clip(level / 4.0)
        explanation = f"Уровень зрелости {level} из 4 → x = {level}/4 = {x:.2f}."

    elif ind_type == "maturity_0_4_capped":
        # S7: потолок при конфликтах с сообществом
        level = data.get("level", 0)
        x = clip(level / 4.0)
        explanation = f"Уровень зрелости {level} из 4 → x = {level}/4 = {x:.2f}."
        if data.get("active_conflict", False):
            x = min(x, 0.5)
            caveats_applied.append("active_community_conflict_cap")
            explanation += " ⚠ Раскрыт действующий конфликт с сообществом → ЖЁСТКИЙ ПОТОЛОК x ≤ 0,5."

    elif ind_type == "minimizer_quant":
        value = data.get("value")
        if value is None:
            return {"code": code, "name": meta["name"], "weight": weight, "x": 0.0, "score": 0.0,
                    "applicable": True, "explanation": "Значение отсутствует → x = 0.",
                    "source_page": data.get("source_page"), "caveats_applied": ["missing_value"]}
        peers = peer_set.get(code, [])
        if len(peers) >= 5:
            x = percentile_minimizer(value, peers)
            explanation = f"Значение {value} в peer-set ({len(peers)} компаний) → перцентиль → x = {x:.2f}."
        elif code in ABSOLUTE_THRESHOLDS:
            x = apply_thresholds_minimizer(value, ABSOLUTE_THRESHOLDS[code])
            caveats_applied.append("absolute_thresholds_applied")
            explanation = f"Значение {value}, применены absolute thresholds (ICMM) → x = {x:.2f}."
        else:
            x = 0.5  # нейтральное значение в безвыходной ситуации
            caveats_applied.append("no_benchmark_available")
            explanation = f"Значение {value}, но peer-set < 5 и нет absolute thresholds → x = 0,5 (нейтрально)."

    elif ind_type == "minimizer_quant_capped":
        # S2 fatalities и G4 коррупция
        value = data.get("value", 0)
        if code == "S2":
            # специальная шкала по абсолюту
            abs_count = data.get("fatalities_abs", value)
            # применяем по таблице или 0.1 если >5
            base_x = ABSOLUTE_THRESHOLDS["S2_fatalities_abs"].get(abs_count, 0.1)
            x = base_x
            explanation = f"Fatalities = {abs_count}. "
            if abs_count > 0:
                x = min(x, 0.5)
                caveats_applied.append("zero_harm_cap")
                explanation += "ЖЁСТКИЙ ПОТОЛОК x ≤ 0,5 (ICMM zero harm)."
            explanation += f" x = {x:.2f}."
        elif code == "G4":
            abs_count = value
            if abs_count == 0:
                x = 1.0
                explanation = "0 подтверждённых случаев коррупции → x = 1,0."
            else:
                x = 0.5  # базовый потолок
                # при больших числах ещё ниже
                if abs_count >= 3: x = 0.3
                if abs_count >= 5: x = 0.1
                x = min(x, 0.5)
                caveats_applied.append("corruption_cap")
                explanation = f"{abs_count} подтверждённых случаев коррупции → ЖЁСТКИЙ ПОТОЛОК x ≤ 0,5. x = {x:.2f}."

    elif ind_type == "maximizer_quant":
        value = data.get("value")
        peers = peer_set.get(code, [])
        if len(peers) >= 5:
            x = percentile_maximizer(value, peers)
            explanation = f"Значение {value} в peer-set → перцентиль → x = {x:.2f}."
        else:
            # абсолютные пороги (в зависимости от показателя)
            thr_key = f"{code}_max"
            if thr_key in ABSOLUTE_THRESHOLDS:
                x = apply_thresholds_maximizer(value, ABSOLUTE_THRESHOLDS[thr_key])
                caveats_applied.append("absolute_thresholds_applied")
                explanation = f"Значение {value}, absolute thresholds → x = {x:.2f}."
            else:
                x = 0.5
                caveats_applied.append("no_benchmark_available")
                explanation = f"Значение {value}, нет benchmarks → x = 0,5."

    elif ind_type == "maximizer_fraction":
        x = clip(data.get("fraction", 0))
        explanation = f"Доля = {x:.2f} → x = {x:.2f}."

    elif ind_type == "composite":
        # E5, E8, G7 — каждый имеет свою формулу, задаваемую через data
        if code == "E5":
            # x = (1 - perc(m3/oz)) * (1 - stress*0.3) + bonus за оборот
            base_x = data.get("water_base_x", 0.5)  # рассчитывается извне или через peer-set
            stress = data.get("stress_share", 0.0)
            closed_loop = data.get("closed_loop_disclosed", False)
            x = base_x * (1 - stress * 0.3)
            if closed_loop:
                x = min(1.0, x + 0.05)
            explanation = f"Базовый x по воде = {base_x:.2f}, доля в зонах стресса = {stress:.2%}, "
            explanation += f"бонус за оборот воды = +0,05. Итог x = {x:.2f}."
        elif code == "E8":
            level = data.get("level", 0)
            protected_share = data.get("protected_zones_share", 0.0)
            x = (level / 4.0) * 0.7 + (1 - protected_share) * 0.3
            explanation = f"Уровень {level}/4 (×0,7) + доля вне охраняемых зон ×0,3 = x = {x:.2f}."
        elif code == "G7":
            has_channel = 1 if data.get("has_channel", False) else 0
            resolved_share = data.get("resolved_share", 0.0) if has_channel else 0.0
            x = 0.5 * has_channel + 0.5 * resolved_share
            explanation = f"Канал = {has_channel}, доля урегулированных = {resolved_share:.2f}. x = {x:.2f}."
        else:
            x = 0.0
            explanation = f"Нет реализации для composite-показателя {code}."

    elif ind_type == "binary":
        # G2: разделение СД и CEO
        x = 1.0 if data.get("separated", False) else 0.0
        explanation = f"Должности {'разделены' if x == 1 else 'совмещены'} → x = {x}."

    elif ind_type == "binary_with_coverage":
        # E11, S3
        certified = data.get("certified", False)
        coverage = data.get("coverage", 0.0)
        if not certified:
            x = 0.0
        elif coverage >= 0.8:
            x = 1.0
        elif coverage >= 0.5:
            x = 0.5
        else:
            x = 0.0
        # S3 bonus
        if code == "S3" and data.get("additional_ohsas", False):
            x = clip(x + 0.05)
        explanation = f"Сертификация = {certified}, охват = {coverage:.0%} → x = {x:.2f}."

    elif ind_type == "fraction_by_thresholds":
        # G1: доля независимых директоров
        share = data.get("share", 0.0)
        criteria_disclosed = data.get("criteria_disclosed", True)
        if share >= 0.5:
            x = 1.0
        elif share >= 0.33:
            x = 0.7
        else:
            x = 0.4
        if not criteria_disclosed:
            x = min(x, 0.5)
            caveats_applied.append("criteria_not_disclosed")
        explanation = f"Доля независимых = {share:.0%}, критерии {'раскрыты' if criteria_disclosed else 'НЕ раскрыты'}. x = {x:.2f}."

    elif ind_type == "three_levels":
        # G3: 0 / 0.5 / 1
        level = data.get("level", 0)  # 0, 1, 2
        x = [0.0, 0.5, 1.0][level]
        explanation = f"Уровень антикоррупционной системы = {level}/2 → x = {x}."

    elif ind_type == "acra_composite":
        # G12: Финансовая прозрачность (АКРА), v4
        # 4 суб-компонента: a (МСФО), b (аудит/доступ), c (IR), d (структура)
        # Жёсткий потолок: нет МСФО → x = 0
        ifrs = data.get("ifrs_published", False)         # bool: консолидированная МСФО опубликована
        quarterly = data.get("quarterly_ifrs", False)    # bool: квартальная (иначе годовая)
        audit_open = data.get("audit_open_access", False)  # bool: аудит + открытый доступ (не только аналитикам)
        audit_closed = data.get("audit_closed_only", False) # bool: есть аудит, но только закрытое предоставление
        ir_materials = data.get("ir_materials", False)   # bool: IR-презентации + операционные KPI
        ir_partial = data.get("ir_partial", False)       # bool: только IR-презентации или только KPI
        structure_transparent = data.get("structure_transparent", False)  # bool: полностью прозрачная структура
        structure_partial = data.get("structure_partial", False)          # bool: частично раскрыта

        if not ifrs:
            x = 0.0
            caveats_applied.append("no_ifrs_hard_cap")
            explanation = "МСФО отсутствует (только РСБУ или нет данных) → ЖЁСТКИЙ ПОТОЛОК x = 0 (методология АКРА)."
        else:
            a = 0.40 + (0.15 if quarterly else 0.0)   # суб-компонент a: 0.40 / 0.55
            b = 0.20 if audit_open else (0.10 if audit_closed else 0.0)
            c = 0.15 if ir_materials else (0.08 if ir_partial else 0.0)
            d = 0.10 if structure_transparent else (0.05 if structure_partial else 0.0)
            x = clip(a + b + c + d)
            # АКРА-уровень для паспорта
            if x >= 0.90:
                acra_level = "Очень высокий"
            elif x >= 0.65:
                acra_level = "Высокий"
            elif x >= 0.35:
                acra_level = "Средний"
            else:
                acra_level = "Низкий"
            explanation = (
                f"МСФО: {'квартальная' if quarterly else 'годовая'} (a={a:.2f}); "
                f"аудит+доступ (b={b:.2f}); IR-материалы (c={c:.2f}); структура группы (d={d:.2f}). "
                f"x = {x:.2f}. Уровень АКРА: {acra_level}."
            )
            data["_acra_level"] = acra_level  # для отчёта
        x = 0.0
        explanation = f"Тип показателя '{ind_type}' не поддержан."

    # ---------- Применение общих понижающих коэффициентов ----------
    for caveat in data.get("caveats", []):
        if caveat == "boundaries_not_described" and code == "E1":
            x = x * 0.9
            caveats_applied.append("boundaries_x_0.9")
        elif caveat == "abs_only_no_specific" and code == "E1":
            x = min(x, 0.6)
            caveats_applied.append("no_specific_x_max_0.6")
        elif caveat == "energy_balance_not_disclosed" and code == "E3":
            x = min(x, 0.7)
            caveats_applied.append("no_energy_breakdown_x_max_0.7")
        elif caveat == "certificates_only_no_ppa" and code == "E4":
            x = min(x, 0.8)
            caveats_applied.append("certs_only_x_max_0.8")
        elif caveat == "hours_base_not_stated" and code == "S1":
            x = min(x, 0.6)
            caveats_applied.append("no_hours_base_x_max_0.6")
        elif caveat == "employees_only_no_contractors" and code == "S1":
            x = min(x, 0.8)
            caveats_applied.append("no_contractors_x_max_0.8")
        elif caveat == "methodology_not_disclosed" and code == "S4":
            x = min(x, 0.6)
            caveats_applied.append("no_methodology_x_max_0.6")
        elif caveat == "critical_criteria_not_disclosed" and code == "S8":
            x = min(x, 0.7)
            caveats_applied.append("no_criticality_criteria_x_max_0.7")

    # ---------- Санкционный бонус (применяется ТОЛЬКО к G8, G9, G10) ----------
    # Логика теперь явно различает три случая:
    #  (1) бонус сработал → triggered=True
    #  (2) флаг есть, но x_raw=0 → triggered=False, reason='no_practice_to_amplify'
    #      (anti-washing: бонус не "вытаскивает из ноля" компанию без реальных практик)
    #  (3) флага нет → sanctions_applied_note = None
    if code in ("G8", "G9", "G10") and data.get("sanctions_barrier", False):
        x_raw_before_bonus = x
        x_bonus_candidate = x_raw_before_bonus + 0.20
        x_doubled_cap = x_raw_before_bonus * 2.0
        x_after = min(x_bonus_candidate, x_doubled_cap, 1.0)

        if x_after > x_raw_before_bonus:
            # Случай (1): бонус сработал
            if abs(x_after - x_bonus_candidate) < 1e-6:
                limit_used = "x_raw + 0.20"
            elif abs(x_after - x_doubled_cap) < 1e-6:
                limit_used = "x_raw × 2 (anti-washing cap)"
            else:
                limit_used = "1.0 (scale max)"
            sanctions_applied_note = {
                "triggered": True,
                "x_raw": round(x_raw_before_bonus, 4),
                "x_final": round(x_after, 4),
                "limit_used": limit_used,
                "bonus_amount": round(x_after - x_raw_before_bonus, 4),
                "trigger_reason": data.get("sanctions_reason", "Санкционный барьер к международной инфраструктуре"),
            }
            x = x_after
            caveats_applied.append("sanctions_bonus_applied")
        else:
            # Случай (2): флаг есть, но бонус не сработал — потому что x_raw = 0.
            # Это методологически верное поведение (anti-washing), но академически важно его явно показать.
            sanctions_applied_note = {
                "triggered": False,
                "x_raw": round(x_raw_before_bonus, 4),
                "x_final": round(x_raw_before_bonus, 4),
                "limit_used": "—",
                "bonus_amount": 0.0,
                "trigger_reason": data.get("sanctions_reason", "Санкционный барьер заявлен"),
                "non_application_reason": (
                    "Санкционный бонус не начислен: причина — отсутствие раскрытой "
                    "фактической практики (x_raw = 0), а не санкционный барьер. "
                    "Anti-washing защита: 0 × 2 = 0, бонус не «вытаскивает из ноля»."
                ),
            }
            caveats_applied.append("sanctions_bonus_not_applied_no_practice")

    # Финальная обрезка
    x = clip(x)
    score = weight * x

    return {
        "code": code, "name": meta["name"], "weight": weight,
        "x": round(x, 4), "score": round(score, 2), "applicable": True,
        "explanation": explanation,
        "source_page": data.get("source_page"),
        "pdf_file": data.get("pdf_file"),
        "pdf_page_ref": data.get("pdf_page_ref") or (f"стр. {data['source_page']}" if data.get("source_page") else None),
        "raw_quote": data.get("raw_quote"),
        "caveats_applied": caveats_applied,
        "sanctions_bonus": sanctions_applied_note,
    }


# ============================================================================
# РАСЧЁТ ПОЛНОГО РЕЙТИНГА
# ============================================================================

def get_rating(score: float) -> dict:
    for lo, hi, label, interp in RATING_SCALE:
        if lo <= score <= hi:
            return {"label": label, "interpretation": interp}
    return {"label": "N/A", "interpretation": "Out of scale"}


def calculate_full_rating(input_data: dict) -> dict:
    """
    input_data = {
        "company": str,
        "year": int,
        "peer_set": {code: [values]},   # опционально
        "indicators": {code: {disclosed, value/level/..., caveats, source_page, raw_quote}}
    }
    """
    company = input_data.get("company", "Unknown")
    year = input_data.get("year")
    peer_set = input_data.get("peer_set", {})
    indicators_data = input_data.get("indicators", {})

    results = []
    block_scores = {"E": 0.0, "S": 0.0, "G": 0.0}
    block_max = {"E": 333, "S": 333, "G": 333}
    block_applicable_max = {"E": 0, "S": 0, "G": 0}

    for code in INDICATORS:
        data = indicators_data.get(code, {})
        result = calculate_indicator(code, data, peer_set)
        results.append(result)
        block = INDICATORS[code]["block"]
        if result.get("applicable", True):
            block_scores[block] += (result.get("score") or 0)
            block_applicable_max[block] += INDICATORS[code]["weight"]

    total_score = sum(block_scores.values())
    # Если есть N/A — пропорционально масштабируем (редкий случай)
    total_applicable_max = sum(block_applicable_max.values())
    if total_applicable_max > 0 and total_applicable_max < 999:
        scaling = 999 / total_applicable_max
        total_score_adj = total_score * scaling
        scaling_note = f"Применено масштабирование из-за N/A-показателей: фактор {scaling:.3f}"
    else:
        total_score_adj = total_score
        scaling_note = None

    return {
        "company": company,
        "year": year,
        "block_scores": {k: round(v, 2) for k, v in block_scores.items()},
        "block_maximums": block_max,
        "block_applicable_max": block_applicable_max,
        "total_score_raw": round(total_score, 2),
        "total_score": round(total_score_adj, 2),
        "rating": get_rating(total_score_adj),
        "indicators": results,
        "scaling_note": scaling_note,
    }


# ============================================================================
# ДИНАМИЧЕСКИЙ АНАЛИЗ
# ============================================================================

def dynamic_analysis(ratings: list) -> dict:
    """Принимает список результатов calculate_full_rating, отсортированных по году."""
    if len(ratings) < 2:
        return {"note": "Нужно ≥2 года для динамического анализа."}

    ratings_sorted = sorted(ratings, key=lambda r: r["year"])
    years = [r["year"] for r in ratings_sorted]
    scores = [r["total_score"] for r in ratings_sorted]

    # 1. Дельта
    deltas = []
    for i in range(1, len(ratings_sorted)):
        delta = scores[i] - scores[i-1]
        deltas.append({"year": years[i], "delta_score": round(delta, 2),
                       "prev_year": years[i-1]})

    # 2. CAGR за весь период
    n = len(ratings_sorted) - 1
    if scores[0] > 0 and n > 0:
        cagr = (scores[-1] / scores[0]) ** (1/n) - 1
    else:
        cagr = None

    # 3. Декомпозиция по блокам (между первым и последним годом)
    first = ratings_sorted[0]
    last = ratings_sorted[-1]
    block_deltas = {
        block: round(last["block_scores"].get(block, 0) - first["block_scores"].get(block, 0), 2)
        for block in ["E", "S", "G"]
    }

    # 4. Топ-5 вкладчиков в рост и падение (по показателям)
    contributions = []
    first_inds = {ind["code"]: ind for ind in first["indicators"]}
    last_inds = {ind["code"]: ind for ind in last["indicators"]}
    for code in INDICATORS:
        f_score = first_inds.get(code, {}).get("score") or 0
        l_score = last_inds.get(code, {}).get("score") or 0
        delta = l_score - f_score
        contributions.append({"code": code, "name": INDICATORS[code]["name"],
                             "delta_score": round(delta, 2),
                             "first_year_score": f_score, "last_year_score": l_score})

    contributions.sort(key=lambda c: c["delta_score"], reverse=True)
    top5_risers = contributions[:5]
    top5_fallers = contributions[-5:][::-1]

    # 5. Рейтинги по годам
    rating_history = [{"year": r["year"], "score": r["total_score"], "rating": r["rating"]["label"]}
                     for r in ratings_sorted]

    return {
        "years": years,
        "score_history": scores,
        "rating_history": rating_history,
        "year_over_year_deltas": deltas,
        "cagr": round(cagr, 4) if cagr is not None else None,
        "block_deltas_first_to_last": block_deltas,
        "top5_risers": top5_risers,
        "top5_fallers": top5_fallers,
    }


# ============================================================================
# CLI
# ============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input_files", nargs="+", help="JSON файлы с входными данными")
    parser.add_argument("--dynamic", action="store_true", help="Провести динамический анализ")
    parser.add_argument("--pretty", action="store_true", help="Человекочитаемый вывод")
    parser.add_argument("-o", "--output", default=None, help="Файл для записи результата")
    args = parser.parse_args()

    ratings = []
    for path in args.input_files:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        rating = calculate_full_rating(data)
        ratings.append(rating)

    output = {"ratings": ratings}
    if args.dynamic and len(ratings) > 1:
        output["dynamic_analysis"] = dynamic_analysis(ratings)

    result = json.dumps(output, indent=2 if args.pretty else None, ensure_ascii=False)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(result)
    else:
        print(result)


if __name__ == "__main__":
    main()
