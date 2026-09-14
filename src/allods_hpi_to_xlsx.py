import argparse
import json
import os
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import quote_sheetname

API_BASE = "https://allods.ru/api/rating"
YANDEX_DISK_API_BASE = "https://cloud-api.yandex.net/v1/disk"
HPI_ENDPOINT = "/hpi/{shard_id}/{type_id}"
HPI_ASTRAL_ENDPOINT = "/hpi/astral/{type_id}/{shard_id}"
ARENA_ENDPOINT = "/arena/{shard_id}/0/{class_id}"
DEFAULT_SHARD_ID = 601
DEFAULT_SHARD_NAME = "\u041c\u043e\u043b\u043e\u0434\u0430\u044f \u0413\u0432\u0430\u0440\u0434\u0438\u044f"
OVERALL_TOP_LIMIT = 50
ISLAND_TOP_LIMIT = 100
ARENA_TOP_LIMIT = 100
OVERALL_TYPE_IDS = {740212236, 740113687}
GROUP_MAX_SIZE = 6
GROUP_RATING_DELTA_THRESHOLD = 3_000_000

FONT_NAME = "Times New Roman"
FONT_SIZE = 14
ROW_HEIGHT = 22.5

CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
THIN_SIDE = Side(style="thin", color="000000")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
TOP_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
OUT_TOP24_FILL = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
GROW_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
DECLINE_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
NEW_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
REMOVED_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
ANOMALY_FILL = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")

HPI_TYPES: list[tuple[int, str]] = [
    (740212236, "РћР±С‰РёР№ СЂРµР№С‚РёРЅРі"),
    (740212225, "Р›Р°Р±РѕСЂР°С‚РѕСЂРёСЏ в„–731"),
    (740212222, "Р—РµРјР»СЏ РўС‹СЃСЏС‡Рё РљСЂС‹Р»СЊРµРІ"),
    (740212219, "РћР±РёС‚РµР»СЊ Р¤РµРЅРёРё"),
    (740212224, "РР·СѓРјСЂСѓРґРЅС‹Р№ РѕСЃС‚СЂРѕРІ"),
    (740212220, "РЈС‚РёРЅС‹Р№ РїР»С‘СЃ"),
    (740212223, "РЎР°РЅР°С‚РѕСЂРёР№ В«РЎРЅРµР¶РёРЅРєР°В»"),
    (740212217, "РњРµРґРЅР°СЏ РіРѕСЂР°"),
    (740212226, "Р›СѓРјРёСЃР°Р°СЂ"),
    (740212221, "РћРіРЅРµС…Р»Р°Рґ"),
    (740212218, "Р‘РµР·РјРѕР»РІРЅР°СЏ РїР°РґСЊ"),
    (740212227, "РџРёРѕРЅРµСЂР»Р°РіРµСЂСЊ В«РњРёСЂР°Р¶В»"),
    (740214334, "Р’РёСЃСЏС‡РёРµ СЃР°РґС‹"),
    (740226908, "РћР°Р·РёСЃ Р™РѕР»РёС†С‚Р»Рё"),
]

HPI_ASTRAL_TYPES: list[tuple[int, str]] = [
    (740113687, "РћР±С‰РёР№ СЂРµР№С‚РёРЅРі"),
    (740113454, "Р›Р°Р±РѕСЂР°С‚РѕСЂРёСЏ в„–731"),
    (740113455, "Р—РµРјР»СЏ РўС‹СЃСЏС‡Рё РљСЂС‹Р»СЊРµРІ"),
    (740113457, "РћР±РёС‚РµР»СЊ Р¤РµРЅРёРё"),
    (740113458, "РР·СѓРјСЂСѓРґРЅС‹Р№ РѕСЃС‚СЂРѕРІ"),
    (740113459, "РЈС‚РёРЅС‹Р№ РїР»С‘СЃ"),
    (740113460, "РЎР°РЅР°С‚РѕСЂРёР№ В«РЎРЅРµР¶РёРЅРєР°В»"),
    (740113461, "РњРµРґРЅР°СЏ РіРѕСЂР°"),
    (740113462, "Р›СѓРјРёСЃР°Р°СЂ"),
    (740113463, "РћРіРЅРµС…Р»Р°Рґ"),
    (740113291, "Р‘РµР·РјРѕР»РІРЅР°СЏ РїР°РґСЊ"),
    (740207909, "РџРёРѕРЅРµСЂР»Р°РіРµСЂСЊ В«РњРёСЂР°Р¶В»"),
    (740214333, "Р’РёСЃСЏС‡РёРµ СЃР°РґС‹"),
    (740226907, "РћР°Р·РёСЃ Р™РѕР»РёС†С‚Р»Рё"),
]

ARENA_CLASSES: list[tuple[int, str]] = [
    (0, "\u0412\u0441\u0435 \u043a\u043b\u0430\u0441\u0441\u044b"),
    (274401281, "\u0411\u0430\u0440\u0434"),
    (61117, "\u0412\u043e\u0438\u043d"),
    (64104, "\u0412\u043e\u043b\u0448\u0435\u0431\u043d\u0438\u043a"),
    (740021330, "\u0414\u0435\u043c\u043e\u043d\u043e\u043b\u043e\u0433"),
    (61119, "\u0416\u0440\u0435\u0446"),
    (739833443, "\u0418\u043d\u0436\u0435\u043d\u0435\u0440"),
    (64105, "\u041c\u0438\u0441\u0442\u0438\u043a"),
    (61121, "\u041d\u0435\u043a\u0440\u043e\u043c\u0430\u043d\u0442"),
    (61123, "\u0420\u0430\u0437\u0432\u0435\u0434\u0447\u0438\u043a"),
    (61120, "\u0425\u0440\u0430\u043c\u043e\u0432\u043d\u0438\u043a"),
    (64103, "\u042f\u0437\u044b\u0447\u043d\u0438\u043a"),
]

GUILD_SHORTCUTS = {
    "РўС‘РјРЅС‹Р№ Р¤РµРЅРёРєСЃ": "Р¤РµРЅРёРєСЃ",
    "РўРµРјРЅС‹Р№ Р¤РµРЅРёРєСЃ": "Р¤РµРЅРёРєСЃ",
    "РїРѕРєРѕСЂРёС‚РµР»Рё РђРЎРўР РђР›Рђ": "РџРђ",
    "РРЅРєРІРёР·РёС‚РѕСЂС‹": "РРЅРєРё",
    "Р С‹С†Р°СЂРё РљСЂРѕРІРєРё": "Р Рљ",
    "Р С‹С†Р°СЂРё РљСЂРѕРІРё": "Р Рљ",
    "Р РµР·РёРґРµРЅС†РёСЏ Р­Р»РёС‚РЅС‹С… Р‘РѕРјР¶РµР№": "Р‘РѕРјР¶Рё",
    "Р РµР·РёРґРµРЅС†РёСЏ Р­Р»РёС‚С‹С… Р‘РѕРјР¶РµР№": "Р‘РѕРјР¶Рё",
}

PLAYER_SHORTCUTS = {
    "РЎР°РјР°Р’СЂРµРґРЅРѕСЃС‚СЊ": "Р’РёРІ",
    "РљР°РїРёС‚Р°РЅРџСЂР°Р№": "РҐРµРЅРЅСЌСЃСЃРё",
    "РљР°РїРёС‚Р°РЅРџСЂР°Р№СЃ": "РҐРµРЅРЅСЌСЃСЃРё",
}

COMPASS_LEVEL_BY_DIFFICULTY: dict[int, int] = {}
COMPASS_LEVEL_BY_DIFFICULTY[740033086] = 1
for lvl in range(2, 26):
    COMPASS_LEVEL_BY_DIFFICULTY[740033169 + lvl] = lvl
for lvl in range(1, 26):
    COMPASS_LEVEL_BY_DIFFICULTY[740033194 + lvl] = lvl
for lvl in range(1, 26):
    COMPASS_LEVEL_BY_DIFFICULTY[740033219 + lvl] = lvl
for lvl in range(1, 26):
    COMPASS_LEVEL_BY_DIFFICULTY[740033244 + lvl] = lvl
for lvl in range(1, 26):
    COMPASS_LEVEL_BY_DIFFICULTY[740040110 + lvl] = lvl

_emerald = [
    740212144, 740212141, 740212138, 740212135, 740212132,
    740212129, 740212151, 740212149, 740212147, 740212143,
    740212140, 740212137, 740212134, 740212131, 740212128,
    740212150, 740212148, 740212146, 740212145, 740212142,
    740212139, 740212136, 740212133, 740212130, 740212127,
]
for idx, difficulty_id in enumerate(_emerald, start=1):
    COMPASS_LEVEL_BY_DIFFICULTY[difficulty_id] = idx


@dataclass
class RatingRecord:
    source: str
    type_name: str
    place: int | None
    character: str
    guild: str
    achievement: int | None
    compass_level: int | None
    clear_time: int | None
    shard: str


@dataclass
class Entity:
    label: str
    by_type: dict[str, RatingRecord]
    overall_place: int
    members: list[str]
    compare_key: str


@dataclass
class ArenaRecord:
    place: int | None
    character: str
    class_name: str
    guild: str
    achievement: int | None
    shard: str


@dataclass(frozen=True)
class ShardConfig:
    shard_id: int
    name: str
    output_name: str
    state_name: str
    yadisk_path: str


SHARD_CONFIGS: tuple[ShardConfig, ...] = (
    ShardConfig(
        shard_id=DEFAULT_SHARD_ID,
        name=DEFAULT_SHARD_NAME,
        output_name="allods_hpi_molodaya_gvardiya.xlsx",
        state_name="allods_state_history.json",
        yadisk_path="/allods/allods_hpi_molodaya_gvardiya.xlsx",
    ),
    ShardConfig(
        shard_id=101,
        name="\u041d\u0430\u0441\u043b\u0435\u0434\u0438\u0435 \u0411\u043e\u0433\u043e\u0432",
        output_name="allods_hpi_nasledie_bogov.xlsx",
        state_name="allods_state_history_nasledie_bogov.json",
        yadisk_path="/allods/allods_hpi_nasledie_bogov.xlsx",
    ),
)

FOCUS_CHARACTERS: tuple[tuple[str, str], ...] = (
    ("\u0424\u043e\u043a\u0443\u0441_\u041d\u0443\u043b\u0451\u0432\u044b\u0439", "\u041d\u0443\u043b\u0451\u0432\u044b\u0439"),
    ("\u0424\u043e\u043a\u0443\u0441_2", ""),
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Р­РєСЃРїРѕСЂС‚ СЂРµР№С‚РёРЅРіРѕРІ hpi/hpi-astral РІ XLSX.")
    parser.add_argument("--output", default=None)
    parser.add_argument("--shard-id", type=int, default=None)
    parser.add_argument("--timeout", type=int, default=30)
    parser.add_argument("--state-file", default=None)
    parser.add_argument("--yadisk-path", default="")
    parser.add_argument("--yadisk-token-env", default="YADISK_TOKEN")
    parser.add_argument("--upload-to-yadisk", action="store_true")
    return parser.parse_args()


def fetch_json(url: str, timeout: int) -> Any:
    req = Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        },
    )
    with urlopen(req, timeout=timeout) as response:
        if response.status != 200:
            raise RuntimeError(f"HTTP {response.status}: {url}")
        body = response.read()
    return json.loads(body.decode("utf-8"))


def request_json(
    url: str,
    timeout: int,
    method: str = "GET",
    headers: dict[str, str] | None = None,
    data: bytes | None = None,
) -> Any:
    req = Request(url, headers=headers or {}, method=method, data=data)
    with urlopen(req, timeout=timeout) as response:
        body = response.read()
    if not body:
        return None
    return json.loads(body.decode("utf-8"))


def as_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


CP1251_REVERSE_MAP = {chr(idx): idx for idx in range(256)}
for _idx in range(256):
    try:
        CP1251_REVERSE_MAP[bytes([_idx]).decode("cp1251")] = _idx
    except UnicodeDecodeError:
        continue


ARENA_HIGHLIGHT_CHARACTERS = {
    "Мгла",
    "Ройс",
    "Волкан",
    "Альфадок",
    "Хави",
    "Альтушечка",
    "Царица",
    "Рукия",
    "Лизелотт",
    "Рангику",
    "ДиректорСвалки",
    "Тродус",
    "Авель",
    "АвелДенан",
}


def _try_decode_mojibake_cp1251(value: str) -> str | None:
    try:
        raw = bytes(CP1251_REVERSE_MAP[ch] for ch in value)
        return raw.decode("utf-8")
    except (KeyError, UnicodeDecodeError):
        return None


def _try_decode_mojibake_cp1252(value: str) -> str | None:
    try:
        return value.encode("cp1252").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return None


def _try_decode_mojibake_latin1(value: str) -> str | None:
    try:
        raw = bytes(ord(ch) for ch in value)
        return raw.decode("utf-8")
    except (ValueError, UnicodeDecodeError):
        return None


def _mojibake_penalty(value: str) -> int:
    penalty = 0
    penalty += len(re.findall(r"[ÐÑ][^\s]", value)) * 10
    penalty += len(re.findall(r"[РС][\u0400-\u04FF]", value)) * 10
    penalty += value.count("в„") * 12
    penalty += value.count("В«") * 4
    penalty += value.count("В»") * 4
    penalty += value.count("\ufffd") * 20
    return penalty


def _text_quality(value: str) -> int:
    cyrillic_count = sum(1 for ch in value if "\u0400" <= ch <= "\u04FF")
    alpha_count = sum(1 for ch in value if ch.isalpha())
    digit_count = sum(1 for ch in value if ch.isdigit())
    whitespace_count = sum(1 for ch in value if ch.isspace())
    return (cyrillic_count * 4) + alpha_count + digit_count + whitespace_count - _mojibake_penalty(value)


def fix_mojibake_text(value: str) -> str:
    best = value
    for _ in range(3):
        best_score = _text_quality(best)
        improved = best
        for candidate in (
            _try_decode_mojibake_cp1251(best),
            _try_decode_mojibake_cp1252(best),
            _try_decode_mojibake_latin1(best),
        ):
            if not candidate or candidate == best:
                continue
            candidate_score = _text_quality(candidate)
            if candidate_score > best_score:
                improved = candidate
                best_score = candidate_score
        if improved == best:
            break
        best = improved
    return best


def human_text(value: Any) -> str:
    return fix_mojibake_text(str(value))


def cli_print(value: Any) -> None:
    print(human_text(value))


def cli_exit(value: Any) -> SystemExit:
    return SystemExit(human_text(value))


def as_int(value: Any) -> int | None:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def short_guild_name(name: str) -> str:
    cleaned = name.strip()
    if not cleaned:
        return "Р‘РµР· РіРёР»СЊРґРёРё"
    return GUILD_SHORTCUTS.get(cleaned, cleaned)


def short_player_name(name: str) -> str:
    cleaned = name.strip()
    if not cleaned:
        return cleaned
    return PLAYER_SHORTCUTS.get(cleaned, cleaned)


def entity_primary_guild(entity: Entity) -> str:
    guild_counter = Counter(rec.guild for rec in entity.by_type.values() if rec.guild)
    if not guild_counter:
        return "Р‘РµР· РіРёР»СЊРґРёРё"
    return short_guild_name(guild_counter.most_common(1)[0][0])


def compass_level_from_difficulty(difficulty: int | None) -> int | None:
    if difficulty is None or difficulty <= 0:
        return None
    return COMPASS_LEVEL_BY_DIFFICULTY.get(difficulty)


def pick_better_record(current: RatingRecord | None, candidate: RatingRecord) -> RatingRecord:
    if current is None:
        return candidate
    current_ach = current.achievement if current.achievement is not None else -1
    candidate_ach = candidate.achievement if candidate.achievement is not None else -1
    if candidate_ach > current_ach:
        return candidate
    if candidate_ach < current_ach:
        return current
    current_place = current.place if current.place is not None else 10**9
    candidate_place = candidate.place if candidate.place is not None else 10**9
    return candidate if candidate_place < current_place else current


def top_limit_for_type_id(type_id: int) -> int:
    return OVERALL_TOP_LIMIT if type_id in OVERALL_TYPE_IDS else ISLAND_TOP_LIMIT


def fetch_source_records(
    source: str,
    endpoint_template: str,
    type_pairs: list[tuple[int, str]],
    shard_id: int,
    timeout: int,
    default_shard_name: str,
) -> list[RatingRecord]:
    rows: list[RatingRecord] = []
    for type_id, type_name in type_pairs:
        type_limit = top_limit_for_type_id(type_id)
        url = API_BASE + endpoint_template.format(shard_id=shard_id, type_id=type_id)
        payload = fetch_json(url, timeout)
        if not isinstance(payload, list):
            raise RuntimeError(f"РќРµРєРѕСЂСЂРµРєС‚РЅС‹Р№ С„РѕСЂРјР°С‚ РѕС‚РІРµС‚Р°: {url}")

        for index, item in enumerate(payload):
            if not isinstance(item, dict):
                continue
            place = as_int(item.get("rating_index")) or (index + 1)
            if place > type_limit:
                continue
            rows.append(
                RatingRecord(
                    source=source,
                    type_name=type_name,
                    place=place,
                    character=as_str(item.get("name")),
                    guild=as_str(item.get("guild")),
                    achievement=as_int(item.get("achievement")),
                    compass_level=compass_level_from_difficulty(as_int(item.get("difficulty"))),
                    clear_time=as_int(item.get("time")),
                    shard=as_str(item.get("shard")) or default_shard_name,
                )
            )
    return rows


def fetch_arena_records(
    shard_id: int,
    timeout: int,
    default_shard_name: str,
) -> tuple[list[ArenaRecord], dict[int, list[ArenaRecord]]]:
    by_class_id = dict(ARENA_CLASSES)
    overall_rows: list[ArenaRecord] = []
    class_rows: dict[int, list[ArenaRecord]] = {}

    for class_id, class_title in ARENA_CLASSES:
        url = API_BASE + ARENA_ENDPOINT.format(shard_id=shard_id, class_id=class_id)
        payload = fetch_json(url, timeout)
        if not isinstance(payload, list):
            raise RuntimeError(f"Р СњР ВµР С”Р С•РЎР‚РЎР‚Р ВµР С”РЎвЂљР Р…РЎвЂ№Р в„– РЎвЂћР С•РЎР‚Р СР В°РЎвЂљ Р С•РЎвЂљР Р†Р ВµРЎвЂљР В°: {url}")

        rows_for_class: list[ArenaRecord] = []
        for index, item in enumerate(payload):
            if not isinstance(item, dict):
                continue
            place = as_int(item.get("rating_index")) or (index + 1)
            if place > ARENA_TOP_LIMIT:
                continue
            class_name = as_str(item.get("class")) or class_title or by_class_id.get(class_id, "")
            record = ArenaRecord(
                place=place,
                character=as_str(item.get("name")),
                class_name=class_name,
                guild=as_str(item.get("guild")),
                achievement=as_int(item.get("achievement")),
                shard=as_str(item.get("shard")) or default_shard_name,
            )
            rows_for_class.append(record)

        if class_id == 0:
            overall_rows = rows_for_class
        else:
            class_rows[class_id] = rows_for_class

    return overall_rows, class_rows


def build_character_entities(rows: list[RatingRecord], ordered_types: list[str]) -> list[Entity]:
    grouped: dict[str, list[RatingRecord]] = defaultdict(list)
    for row in rows:
        if row.character:
            grouped[row.character].append(row)

    entities: list[Entity] = []
    for character, items in grouped.items():
        per_type: dict[str, RatingRecord] = {}
        for item in items:
            per_type[item.type_name] = pick_better_record(per_type.get(item.type_name), item)
        overall = per_type.get("РћР±С‰РёР№ СЂРµР№С‚РёРЅРі")
        overall_place = overall.place if overall and overall.place is not None else 10**9
        entities.append(
            Entity(
                label=character,
                by_type=per_type,
                overall_place=overall_place,
                members=[character],
                compare_key=character,
            )
        )
    entities.sort(key=lambda e: (e.overall_place, e.label.lower()))
    return entities


def build_group_entities(character_entities: list[Entity], ordered_types: list[str]) -> list[Entity]:
    def entity_signature(entity: Entity) -> tuple[int | None, ...]:
        return tuple(
            entity.by_type.get(type_name).achievement if entity.by_type.get(type_name) else None
            for type_name in ordered_types
        )

    def overall_achievement(entity: Entity) -> int | None:
        rec = entity.by_type.get("РћР±С‰РёР№ СЂРµР№С‚РёРЅРі")
        return rec.achievement if rec and rec.achievement is not None else None

    def compass_25_times(entity: Entity) -> set[tuple[str, int]]:
        times: set[tuple[str, int]] = set()
        for type_name, rec in entity.by_type.items():
            if rec.compass_level == 25 and rec.clear_time is not None and rec.clear_time > 0:
                times.add((type_name, rec.clear_time))
        return times

    signatures = {id(entity): entity_signature(entity) for entity in character_entities}
    ratings = {id(entity): overall_achievement(entity) for entity in character_entities}
    c25_times = {id(entity): compass_25_times(entity) for entity in character_entities}

    def can_be_grouped(anchor: Entity, candidate: Entity) -> bool:
        anchor_id = id(anchor)
        candidate_id = id(candidate)
        if signatures[anchor_id] == signatures[candidate_id]:
            return True
        anchor_rating = ratings[anchor_id]
        candidate_rating = ratings[candidate_id]
        if anchor_rating is None or candidate_rating is None:
            return False
        if abs(anchor_rating - candidate_rating) > GROUP_RATING_DELTA_THRESHOLD:
            return False
        return bool(c25_times[anchor_id] & c25_times[candidate_id])

    pool = sorted(character_entities, key=lambda e: (e.overall_place, e.label.lower()))
    groups: list[Entity] = []
    while pool:
        anchor = pool.pop(0)
        anchor_id = id(anchor)
        compatible: list[tuple[int, int, int, str, Entity]] = []
        for candidate in pool:
            if not can_be_grouped(anchor, candidate):
                continue
            candidate_id = id(candidate)
            exact_match = signatures[anchor_id] == signatures[candidate_id]
            anchor_rating = ratings[anchor_id]
            candidate_rating = ratings[candidate_id]
            rating_delta = (
                abs(anchor_rating - candidate_rating)
                if anchor_rating is not None and candidate_rating is not None
                else 10**9
            )
            compatible.append(
                (
                    0 if exact_match else 1,
                    rating_delta,
                    candidate.overall_place,
                    candidate.label.lower(),
                    candidate,
                )
            )
        compatible.sort()
        selected = [entry[4] for entry in compatible[: max(0, GROUP_MAX_SIZE - 1)]]
        selected_ids = {id(entity) for entity in selected}
        pool = [entity for entity in pool if id(entity) not in selected_ids]
        members = [anchor, *selected]
        members_sorted = sorted(members, key=lambda e: (e.overall_place, e.label.lower()))
        representative = members_sorted[0]
        label_representative = sorted(members, key=lambda e: e.label.lower())[0]

        guild_counter = Counter(
            rec.guild for member in members_sorted for rec in member.by_type.values() if rec.guild
        )
        group_guild = short_guild_name(guild_counter.most_common(1)[0][0] if guild_counter else "Р‘РµР· РіРёР»СЊРґРёРё")
        compare_key = label_representative.label.strip() or short_player_name(label_representative.label)
        label = f"{short_player_name(label_representative.label)}({group_guild})"
        if len(members_sorted) > 1:
            label = f"{label}x{len(members_sorted)}"

        members_text: list[str] = []
        for member in members_sorted:
            member_name = short_player_name(member.label)
            member_guild = entity_primary_guild(member)
            if member_guild != group_guild:
                member_name = f"{member_name} ({member_guild})"
            members_text.append(member_name)

        groups.append(
            Entity(
                label=label,
                by_type=representative.by_type,
                overall_place=representative.overall_place,
                members=members_text,
                compare_key=compare_key,
            )
        )
    groups.sort(key=lambda e: (e.overall_place, e.label.lower()))
    return groups


def build_top_values_by_type(rows: list[RatingRecord], ordered_types: list[str], metric: str) -> dict[str, int | None]:
    values: dict[str, list[int]] = defaultdict(list)
    for row in rows:
        if metric == "achievement":
            val = row.achievement
        elif metric == "compass":
            val = row.compass_level
        else:
            val = row.clear_time if row.clear_time and row.clear_time > 0 else None
        if val is not None:
            values[row.type_name].append(val)

    out: dict[str, int | None] = {}
    for type_name in ordered_types:
        vals = values.get(type_name, [])
        if not vals:
            out[type_name] = None
        elif metric == "time":
            out[type_name] = min(vals)
        else:
            out[type_name] = max(vals)
    return out


def format_time_ms(value: int | None) -> str | None:
    if value is None or value <= 0:
        return None
    total_ms = int(value)
    minutes = total_ms // 60000
    seconds = (total_ms % 60000) // 1000
    return f"{minutes}:{seconds:02d}"


def format_duration_ms(value: int | None) -> str:
    if value is None or value < 0:
        return "-"
    total_ms = int(value)
    minutes = total_ms // 60000
    seconds = (total_ms % 60000) // 1000
    return f"{minutes}:{seconds:02d}"


def format_time_change(raw_ms: int | None) -> str:
    if raw_ms is None or raw_ms <= 0:
        return "-"
    return format_time_ms(raw_ms) or "-"


def format_signed_duration_ms(value: int | None) -> str:
    if value is None:
        return "-"
    sign = "-" if value < 0 else "+"
    return f"{sign}{format_duration_ms(abs(value))}"


def build_island_time_delta_cells(entities: list[Entity], type_name: str) -> list[str]:
    records = [
        record
        for entity in entities
        if (record := entity.by_type.get(type_name)) is not None
        and record.clear_time is not None
        and record.clear_time > 0
    ][:3]

    def delta_cell(left_index: int, right_index: int) -> str:
        if right_index >= len(records):
            return "-"
        left = records[left_index].clear_time
        right = records[right_index].clear_time
        if left is None or right is None:
            return "-"
        return format_signed_duration_ms(right - left)

    return [delta_cell(0, 1), delta_cell(1, 2)]


def metric_value(record: RatingRecord | None, metric: str) -> int | str | None:
    if record is None:
        return None
    if metric == "achievement":
        return record.achievement
    if metric == "compass":
        return record.compass_level
    if metric == "compass_time":
        compass = record.compass_level
        time_text = format_time_ms(record.clear_time)
        if compass is None or time_text is None:
            return None
        return f"{compass} - {time_text}"
    return format_time_ms(record.clear_time)


def total_time_cell_for_types(entity: Entity, type_names: list[str]) -> str | None:
    total = 0
    completed = 0
    for type_name in type_names:
        rec = entity.by_type.get(type_name)
        if rec is None or rec.clear_time is None or rec.clear_time <= 0:
            continue
        total += rec.clear_time
        completed += 1

    remaining = len(type_names) - completed
    total_text = format_time_ms(total if completed > 0 else None)
    if remaining <= 0:
        return total_text
    if total_text is None:
        return f"+{remaining}"
    return f"{total_text} +{remaining}"


def autosize_columns(
    sheet,
    min_first: int = 38,
    min_other: int = 34,
    max_width: int = 160,
) -> None:
    for col in sheet.columns:
        max_len = 0
        col_idx = col[0].column
        for cell in col:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        width = max_len + 2
        if col_idx == 1:
            width = max(width, min_first)
        else:
            width = max(width, min_other)
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(width, max_width)


def apply_layout(sheet) -> None:
    for row in sheet.iter_rows():
        for cell in row:
            is_bold = bool(cell.font and cell.font.bold)
            cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=is_bold)
            cell.alignment = CENTER_ALIGNMENT
            cell.border = THIN_BORDER
    for row_idx in range(1, sheet.max_row + 1):
        sheet.row_dimensions[row_idx].height = ROW_HEIGHT


def _safe_sheet_title(title: str, existing_titles: set[str]) -> str:
    cleaned = "".join(ch for ch in title if ch not in '[]:*?/\\').strip().strip("'")
    if not cleaned:
        cleaned = "Sheet"
    if len(cleaned) <= 31 and cleaned not in existing_titles:
        return cleaned

    base = cleaned[:31]
    if base not in existing_titles:
        return base

    counter = 2
    while True:
        suffix = f"_{counter}"
        prefix_len = max(1, 31 - len(suffix))
        candidate = f"{cleaned[:prefix_len]}{suffix}"
        if candidate not in existing_titles:
            return candidate
        counter += 1


def create_sheet_safe(workbook: Workbook, title: str):
    safe_title = _safe_sheet_title(title, set(workbook.sheetnames))
    return workbook.create_sheet(safe_title)


def is_best_compass_time(
    record: RatingRecord | None,
    records: list[RatingRecord],
) -> bool:
    if record is None or record.compass_level is None or record.clear_time is None or record.clear_time <= 0:
        return False
    candidates = [
        r for r in records
        if r.compass_level is not None and r.clear_time is not None and r.clear_time > 0
    ]
    if not candidates:
        return False
    best = sorted(
        candidates,
        key=lambda r: (-(r.compass_level or -1), r.clear_time or 10**12, r.place or 10**9),
    )[0]
    return best.compass_level == record.compass_level and best.clear_time == record.clear_time


def write_metric_matrix_sheet(
    workbook: Workbook,
    sheet_title: str,
    entities: list[Entity],
    rows: list[RatingRecord],
    ordered_types: list[str],
    mode: str,
) -> None:
    sheet = create_sheet_safe(workbook, sheet_title)
    headers = ["Р РµР№С‚РёРЅРі/РћСЃС‚СЂРѕРІ"]
    if mode == "compass_time":
        headers.extend(["Разница 1-2", "Разница 2-3"])
    headers.extend(e.label for e in entities)
    sheet.append(headers)

    row_specs = [("Р РµР№С‚РёРЅРі", "achievement")] if mode == "rating" else [("РљРѕРјРїР°СЃ - Р’СЂРµРјСЏ", "compass_time")]
    top_by_type = build_top_values_by_type(rows, ordered_types, "achievement")
    top_cells: list[tuple[int, int]] = []
    records_by_type: dict[str, list[RatingRecord]] = defaultdict(list)
    for row in rows:
        records_by_type[row.type_name].append(row)

    visible_types = ordered_types if mode == "rating" else [t for t in ordered_types if t != "РћР±С‰РёР№ СЂРµР№С‚РёРЅРі"]
    row_idx = 2
    for type_name in visible_types:
        for _, metric_key in row_specs:
            row_values: list[Any] = [type_name]
            if mode == "compass_time":
                row_values.extend(build_island_time_delta_cells(entities, type_name))
            row_records: list[RatingRecord | None] = []
            for entity in entities:
                rec = entity.by_type.get(type_name)
                row_records.append(rec)
                row_values.append(metric_value(rec, metric_key))
            sheet.append(row_values)

            entity_start_col = 4 if mode == "compass_time" else 2
            for col_idx, rec in enumerate(row_records, start=entity_start_col):
                if metric_key == "achievement":
                    top_value = top_by_type.get(type_name)
                    if top_value is not None and rec is not None and rec.achievement == top_value:
                        top_cells.append((row_idx, col_idx))
                elif metric_key == "compass_time" and is_best_compass_time(rec, records_by_type[type_name]):
                    top_cells.append((row_idx, col_idx))
            row_idx += 1

    if mode == "compass_time":
        total_row: list[Any] = ["РћР±С‰РµРµ РІСЂРµРјСЏ Р·Р° РІСЃРµ РѕСЃС‚СЂРѕРІР°"]
        total_row.extend(["-", "-"])
        for entity in entities:
            total_row.append(total_time_cell_for_types(entity, visible_types))
        sheet.append(total_row)

    entity_start_col = 4 if mode == "compass_time" else 2
    for col_idx, entity in enumerate(entities, start=entity_start_col):
        if entity.overall_place > 24:
            for r in range(1, sheet.max_row + 1):
                sheet.cell(row=r, column=col_idx).fill = OUT_TOP24_FILL

    for r, c in top_cells:
        sheet.cell(row=r, column=c).fill = TOP_FILL

    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    for cell in sheet["A"]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)

    if mode == "compass_time":
        apply_filter_and_freeze(sheet, header_row=1, freeze_cell="D2")
        autosize_columns(sheet, min_first=57, min_other=30, max_width=72)
        for column_letter in ("B", "C"):
            sheet.column_dimensions[column_letter].width = max(sheet.column_dimensions[column_letter].width or 0, 18)
    else:
        apply_filter_and_freeze(sheet, header_row=1, freeze_cell="B2")
        autosize_columns(sheet, min_first=38, min_other=28, max_width=96)
    apply_layout(sheet)


def write_group_description_sheet(workbook: Workbook, sheet_title: str, entities: list[Entity]) -> None:
    sheet = create_sheet_safe(workbook, sheet_title)
    sheet.append(["Р“СЂСѓРїРїР°", "РЎРѕСЃС‚Р°РІ РіСЂСѓРїРїС‹"])
    for entity in entities:
        if len(entity.members) <= 1:
            continue
        sheet.append([entity.label, ", ".join(entity.members)])

    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    autosize_columns(sheet, min_first=42, min_other=40, max_width=220)
    current_b = sheet.column_dimensions["B"].width or 0
    sheet.column_dimensions["B"].width = max(current_b, 102)
    current_a = sheet.column_dimensions["A"].width or 0
    sheet.column_dimensions["A"].width = max(current_a, 64)
    apply_filter_and_freeze(sheet, header_row=1, freeze_cell="B2")
    apply_layout(sheet)


def build_snapshot_for_source(
    source_name: str,
    groups: list[Entity],
    ordered_types: list[str],
) -> dict[str, Any]:
    source_snapshot: dict[str, Any] = {"groups": {}}
    for group in groups:
        type_map: dict[str, Any] = {}
        for type_name in ordered_types:
            rec = group.by_type.get(type_name)
            if rec is None:
                type_map[type_name] = None
            else:
                type_map[type_name] = {
                    "achievement": rec.achievement,
                    "compass": rec.compass_level,
                    "time": rec.clear_time,
                }
        source_snapshot["groups"][group.compare_key] = type_map
    return source_snapshot


def normalize_group_key_for_compare(name: str) -> str:
    cleaned = name.strip()
    if not cleaned:
        return cleaned
    cleaned = re.sub(r"x\d+$", "", cleaned).strip()
    if cleaned.endswith(")") and "(" in cleaned:
        return cleaned[:cleaned.rfind("(")].strip()
    return cleaned


def normalize_groups_map_for_compare(groups: dict[str, Any]) -> dict[str, dict[str, Any]]:
    normalized: dict[str, dict[str, Any]] = {}
    for raw_name in sorted(groups.keys()):
        raw_types = groups.get(raw_name)
        if not isinstance(raw_types, dict):
            continue
        normalized_name = normalize_group_key_for_compare(raw_name)
        if not normalized_name:
            continue
        if normalized_name not in normalized:
            normalized[normalized_name] = {}
        normalized[normalized_name].update(raw_types)
    return normalized


def apply_filter_and_freeze(sheet, header_row: int = 1, freeze_cell: str = "B2") -> None:
    if sheet.max_row >= header_row and sheet.max_column >= 1:
        end_col = get_column_letter(sheet.max_column)
        sheet.auto_filter.ref = f"A{header_row}:{end_col}{sheet.max_row}"
    sheet.freeze_panes = freeze_cell


def add_back_to_toc_link(sheet, toc_title: str) -> None:
    cell = sheet["A1"]
    cell.hyperlink = f"#{quote_sheetname(toc_title)}!A1"
    cell.font = Font(name=FONT_NAME, size=FONT_SIZE, underline="single", color="0563C1", bold=True)


def load_state_history(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(payload, dict) and isinstance(payload.get("runs"), list):
            return [run for run in payload["runs"] if isinstance(run, dict)]
    except Exception:
        return []
    return []


def save_state(path: Path, state: dict[str, Any]) -> None:
    history: dict[str, Any] = {"version": 1, "runs": []}
    if path.exists():
        try:
            loaded = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(loaded, dict) and isinstance(loaded.get("runs"), list):
                history = loaded
        except Exception:
            history = {"version": 1, "runs": []}

    history["runs"].append(state)
    path.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")


def classify_change(prev_val: Any, curr_val: Any, rating_delta: int | None, prev_time: int | None, curr_time: int | None) -> tuple[str, bool]:
    if prev_val is None and isinstance(curr_val, dict):
        return "РќРѕРІР°СЏ РіСЂСѓРїРїР°", False
    if isinstance(prev_val, dict) and curr_val is None:
        return "Р“СЂСѓРїРїР° РїСЂРѕРїР°Р»Р°", True
    if curr_val is None:
        return "РќРµС‚ РґР°РЅРЅС‹С…", True
    if isinstance(rating_delta, int) and rating_delta < 0:
        return "РђРЅРѕРјР°Р»РёСЏ РґР°РЅРЅС‹С…", True
    if isinstance(prev_time, int) and isinstance(curr_time, int) and prev_time > 0 and curr_time > 0 and curr_time < prev_time:
        return "РЎРёР»СЊРЅРѕ Р»СѓС‡С€РµРµ РІСЂРµРјСЏ", True
    return "РР·РјРµРЅРµРЅРёРµ СЂРµР·СѓР»СЊС‚Р°С‚Р°", False


def change_fill_by_type(change_type: str, odd_case: bool, rating_delta: int | None) -> PatternFill | None:
    if odd_case:
        return ANOMALY_FILL
    if change_type == "РќРѕРІР°СЏ РіСЂСѓРїРїР°":
        return NEW_FILL
    if change_type == "Р“СЂСѓРїРїР° РїСЂРѕРїР°Р»Р°":
        return REMOVED_FILL
    if isinstance(rating_delta, int):
        if rating_delta > 0:
            return GROW_FILL
        if rating_delta < 0:
            return DECLINE_FILL
    return None


def collect_changes(
    previous: dict[str, Any],
    current: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    changes: list[dict[str, Any]] = []
    odd_changes: list[dict[str, Any]] = []
    data_errors: list[dict[str, Any]] = []

    for source_name in ("РџСЂС‹РіРѕРј", "РљРѕСЂРѕРј"):
        prev_source = previous.get(source_name, {})
        curr_source = current.get(source_name, {})

        prev_groups_raw = prev_source.get("groups", {}) if isinstance(prev_source.get("groups"), dict) else {}
        curr_groups_raw = curr_source.get("groups", {}) if isinstance(curr_source.get("groups"), dict) else {}
        prev_groups = normalize_groups_map_for_compare(prev_groups_raw)
        curr_groups = normalize_groups_map_for_compare(curr_groups_raw)

        all_group_names = sorted(set(prev_groups.keys()) | set(curr_groups.keys()))
        for group_name in all_group_names:
            prev_types = prev_groups.get(group_name, {})
            curr_types = curr_groups.get(group_name, {})
            all_types = sorted(set(prev_types.keys()) | set(curr_types.keys()))
            for type_name in all_types:
                prev_val = prev_types.get(type_name)
                curr_val = curr_types.get(type_name)
                if prev_val == curr_val:
                    continue

                prev_ach = prev_val.get("achievement") if isinstance(prev_val, dict) else None
                curr_ach = curr_val.get("achievement") if isinstance(curr_val, dict) else None
                if isinstance(prev_ach, int) and isinstance(curr_ach, int):
                    rating_delta = curr_ach - prev_ach
                elif prev_ach is None and isinstance(curr_ach, int):
                    rating_delta = curr_ach
                elif isinstance(prev_ach, int) and curr_ach is None:
                    rating_delta = -prev_ach
                else:
                    rating_delta = None

                prev_compass = prev_val.get("compass") if isinstance(prev_val, dict) else None
                curr_compass = curr_val.get("compass") if isinstance(curr_val, dict) else None
                prev_time = prev_val.get("time") if isinstance(prev_val, dict) else None
                curr_time = curr_val.get("time") if isinstance(curr_val, dict) else None
                prev_ct = f"{prev_compass or '-'} / {format_time_change(prev_time)}" if isinstance(prev_val, dict) else "-"
                curr_ct = f"{curr_compass or '-'} / {format_time_change(curr_time)}" if isinstance(curr_val, dict) else "-"
                if prev_ct == curr_ct and rating_delta in (None, 0):
                    continue

                change_type, odd_case = classify_change(prev_val, curr_val, rating_delta, prev_time, curr_time)
                payload = {
                    "source": source_name,
                    "group": group_name,
                    "type_name": type_name,
                    "metric": "РљРѕРјРїР°СЃ/Р’СЂРµРјСЏ",
                    "change_type": change_type,
                    "before": prev_ct,
                    "after": curr_ct,
                    "rating_delta": rating_delta,
                    "odd_case": odd_case,
                }

                if isinstance(rating_delta, int) and rating_delta < 0:
                    payload["change_type"] = "РћС€РёР±РєР° РґР°РЅРЅС‹С… (СЂРµР№С‚РёРЅРі СѓРјРµРЅСЊС€РёР»СЃСЏ)"
                    payload["odd_case"] = True
                    data_errors.append(payload)
                    continue

                if odd_case:
                    odd_changes.append(payload)
                else:
                    changes.append(payload)

    return changes, odd_changes, data_errors


def _write_changes_payload(sheet, payloads: list[dict[str, Any]], empty_msg: str) -> None:
    headers = [
        "РСЃС‚РѕС‡РЅРёРє",
        "Р“СЂСѓРїРїР°",
        "РћСЃС‚СЂРѕРІ/Р РµР№С‚РёРЅРі",
        "РњРµС‚СЂРёРєР°",
        "РўРёРї РёР·РјРµРЅРµРЅРёСЏ",
        "Р‘С‹Р»Рѕ",
        "РЎС‚Р°Р»Рѕ",
        "Р”РµР»СЊС‚Р° СЂРµР№С‚РёРЅРіР°",
    ]
    sheet.append(headers)

    if not payloads:
        sheet.append(["-", "-", "-", "-", empty_msg, "-", "-", "-"])
    else:
        for item in payloads:
            sheet.append(
                [
                    item["source"],
                    item["group"],
                    item["type_name"],
                    item["metric"],
                    item["change_type"],
                    item["before"],
                    item["after"],
                    str(item["rating_delta"]) if isinstance(item["rating_delta"], int) else "-",
                ]
            )
            fill = change_fill_by_type(item["change_type"], bool(item["odd_case"]), item["rating_delta"])
            if fill is not None:
                for col in range(1, len(headers) + 1):
                    sheet.cell(row=sheet.max_row, column=col).fill = fill

    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    autosize_columns(sheet, min_first=16, min_other=24, max_width=180)
    sheet.column_dimensions["B"].width = min(max(sheet.column_dimensions["B"].width or 24, 30), 180)
    sheet.column_dimensions["C"].width = min(max(sheet.column_dimensions["C"].width or 24, 30), 180)
    apply_filter_and_freeze(sheet, header_row=1, freeze_cell="B2")
    apply_layout(sheet)


def write_changes_sheet(workbook: Workbook, previous: dict[str, Any], current: dict[str, Any]) -> None:
    sheet = create_sheet_safe(workbook, "РР·РјРµРЅРµРЅРёСЏ")
    changes, _, _ = collect_changes(previous, current)
    _write_changes_payload(sheet, changes, "РќРµС‚ РёР·РјРµРЅРµРЅРёР№")


def write_changes_sheets_with_odd(
    workbook: Workbook,
    previous: dict[str, Any],
    current: dict[str, Any],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    changes, odd_changes, data_errors = collect_changes(previous, current)
    sheet = create_sheet_safe(workbook, "РР·РјРµРЅРµРЅРёСЏ")
    odd_sheet = create_sheet_safe(workbook, "РЎС‚СЂР°РЅРЅС‹Рµ РёР·РјРµРЅРµРЅРёСЏ")
    err_sheet = create_sheet_safe(workbook, "РћС€РёР±РєРё РґР°РЅРЅС‹С…")
    _write_changes_payload(sheet, changes, "РќРµС‚ РёР·РјРµРЅРµРЅРёР№")
    _write_changes_payload(odd_sheet, odd_changes, "РќРµС‚ СЃС‚СЂР°РЅРЅС‹С… РёР·РјРµРЅРµРЅРёР№")
    _write_changes_payload(err_sheet, data_errors, "РќРµС‚ РѕС€РёР±РѕРє РґР°РЅРЅС‹С…")
    return changes, odd_changes, data_errors


def best_record_for_character(rows: list[RatingRecord], type_name: str, character: str) -> RatingRecord | None:
    best: RatingRecord | None = None
    for row in rows:
        if row.type_name != type_name:
            continue
        if row.character != character:
            continue
        best = pick_better_record(best, row)
    return best


def best_record_for_type(rows: list[RatingRecord], type_name: str) -> RatingRecord | None:
    best: RatingRecord | None = None
    for row in rows:
        if row.type_name != type_name:
            continue
        best = pick_better_record(best, row)
    return best


def overall_cohort_characters(
    rows: list[RatingRecord],
    place_from: int,
    place_to: int,
) -> set[str]:
    return {
        r.character
        for r in rows
        if r.type_name == "РћР±С‰РёР№ СЂРµР№С‚РёРЅРі"
        and r.character
        and r.place is not None
        and place_from <= r.place <= place_to
    }


def target_record_for_type(
    rows: list[RatingRecord],
    type_name: str,
    character: str,
    cohort_chars: set[str] | None = None,
) -> RatingRecord | None:
    if character == "Р¦Р°СЂРёС†Р°" and cohort_chars:
        candidates = [r for r in rows if r.type_name == type_name and r.character in cohort_chars]
        best: RatingRecord | None = None
        for candidate in candidates:
            best = pick_better_record(best, candidate)
        if best is not None:
            return best
    return best_record_for_type(rows, type_name)


def min_achievement_for_type(rows: list[RatingRecord], type_name: str) -> int | None:
    vals = [r.achievement for r in rows if r.type_name == type_name and r.achievement is not None]
    return min(vals) if vals else None


def compass_stats_for_type(rows: list[RatingRecord], type_name: str) -> tuple[int | None, int | None]:
    vals = sorted(
        [
            r.compass_level
            for r in rows
            if r.type_name == type_name and r.compass_level is not None and r.compass_level > 0
        ]
    )
    if not vals:
        return None, None
    median = vals[len(vals) // 2]
    top = vals[-1]
    return median, top


def render_record_cell(record: RatingRecord | None) -> str:
    if record is None or record.achievement is None:
        return "-"
    parts: list[str] = [f"Рњ{record.place or '-'}", f"Р {record.achievement}"]
    if record.compass_level is not None:
        parts.append(f"Рљ{record.compass_level}")
    tm = format_time_ms(record.clear_time)
    if tm:
        parts.append(tm)
    return " | ".join(parts)


def recommendation_for_type(
    type_name: str,
    hero: RatingRecord | None,
    top: RatingRecord | None,
    entry_threshold: int | None,
    type_rows: list[RatingRecord],
) -> str:
    if type_name == "РћР±С‰РёР№ СЂРµР№С‚РёРЅРі":
        return "-"

    median_compass, top_compass = compass_stats_for_type(type_rows, type_name)

    if hero is None:
        target_compass = top_compass or median_compass
        if target_compass is not None:
            return f"РџСЂРѕР№РґРё РѕСЃС‚СЂРѕРІ РЅР° РєРѕРјРїР°СЃРµ {target_compass}, С‡С‚РѕР±С‹ РІРѕР№С‚Рё РІ С‚РѕРї-100."
        if entry_threshold is not None:
            return f"РџСЂРѕР№РґРё РѕСЃС‚СЂРѕРІ РЅР° РїРѕРґС…РѕРґСЏС‰РµРј РєРѕРјРїР°СЃРµ: С†РµР»СЊ СЂРµР№С‚РёРЅРіР° РЅРµ РЅРёР¶Рµ {entry_threshold}."
        return "РџСЂРѕР№РґРё РѕСЃС‚СЂРѕРІ Рё РІРѕР№РґРё РІ С‚РѕРї-100."

    if top is None or hero.achievement is None or top.achievement is None:
        if hero.compass_level is not None:
            return f"РЈРґРµСЂР¶РёРІР°Р№ РїСЂРѕС…РѕР¶РґРµРЅРёРµ РЅР° РєРѕРјРїР°СЃРµ {hero.compass_level}."
        return "Р’РѕР№С‚Рё РІ С‚РѕРї-100 РїРѕ СЌС‚РѕРјСѓ РѕСЃС‚СЂРѕРІСѓ."

    if hero.achievement == top.achievement:
        if hero.clear_time and top.clear_time and hero.clear_time > top.clear_time:
            return f"РћСЃС‚Р°РІСЊ РєРѕРјРїР°СЃ {hero.compass_level or '-'}, РЅРѕ СѓР»СѓС‡С€Рё РІСЂРµРјСЏ РґРѕ ~{format_time_ms(top.clear_time)}."
        return "РўРѕРї-СЂРµР·СѓР»СЊС‚Р°С‚, СѓРґРµСЂР¶РёРІР°Р№ С‚РµРєСѓС‰РёР№ СѓСЂРѕРІРµРЅСЊ."

    hero_compass = hero.compass_level
    top_compass = top.compass_level
    if median_compass is not None and (hero_compass or 0) < median_compass:
        if hero_compass is None:
            return f"РџСЂРѕР№РґРё РѕСЃС‚СЂРѕРІ РЅР° РєРѕРјРїР°СЃРµ {median_compass} Рё РІС‹С€Рµ."
        return f"РџРµСЂРµРїСЂРѕР№РґРё РѕСЃС‚СЂРѕРІ: СЃРµР№С‡Р°СЃ РєРѕРјРїР°СЃ {hero_compass}, С†РµР»СЊ РјРёРЅРёРјСѓРј {median_compass}."

    hero_time = hero.clear_time if hero.clear_time and hero.clear_time > 0 else None
    top_time = top.clear_time if top.clear_time and top.clear_time > 0 else None
    if hero_time and top_time and hero_time > top_time:
        if hero_compass is not None:
            return f"РЈР»СѓС‡С€Рё РІСЂРµРјСЏ РЅР° РєРѕРјРїР°СЃРµ {hero_compass}: С†РµР»СЊ ~{format_time_ms(top_time)} Рё Р±С‹СЃС‚СЂРµРµ."
        return f"РЈР»СѓС‡С€Рё РІСЂРµРјСЏ: С†РµР»СЊ ~{format_time_ms(top_time)} Рё Р±С‹СЃС‚СЂРµРµ."

    gap = top.achievement - hero.achievement
    if hero_compass is not None:
        return f"РџРµСЂРµРїСЂРѕР№РґРё РѕСЃС‚СЂРѕРІ РЅР° РєРѕРјРїР°СЃРµ {hero_compass}, РґРѕР±РµСЂРё ~{gap} СЂРµР№С‚РёРЅРіР°."
    return f"РџСЂРѕР№РґРё РѕСЃС‚СЂРѕРІ РЅР° Р±РѕР»РµРµ РІС‹СЃРѕРєРѕРј РєРѕРјРїР°СЃРµ, РґРѕР±РµСЂРё ~{gap} СЂРµР№С‚РёРЅРіР°."


def write_focus_sheet(
    workbook: Workbook,
    sheet_title: str,
    character: str,
    hpi_rows: list[RatingRecord],
    hpi_types: list[str],
    astral_rows: list[RatingRecord],
    astral_types: list[str],
) -> None:
    sheet = create_sheet_safe(workbook, sheet_title)
    sheet.append([f"РђРЅР°Р»РёР· РїРµСЂСЃРѕРЅР°Р¶Р°: {character}"])
    if character == "Р¦Р°СЂРёС†Р°":
        sheet.append(["РћСЂРёРµРЅС‚РёСЂ", "Р”Р»СЏ Р¦Р°СЂРёС†Р° С†РµР»РµРІР°СЏ РєРѕРіРѕСЂС‚Р° Р±РµСЂС‘С‚СЃСЏ РїРѕ РѕР±С‰РµРјСѓ СЂРµР№С‚РёРЅРіСѓ: РјРµСЃС‚Р° 7-12."])
    sheet.append(["РСЃС‚РѕС‡РЅРёРє", "Р РµР№С‚РёРЅРі/РћСЃС‚СЂРѕРІ", "Р РµР·СѓР»СЊС‚Р°С‚ РїРµСЂСЃРѕРЅР°Р¶Р°", "Р¦РµР»РµРІРѕР№ РѕСЂРёРµРЅС‚РёСЂ", "РћС‚СЃС‚Р°РІР°РЅРёРµ РїРѕ СЂРµР№С‚РёРЅРіСѓ", "Р РµРєРѕРјРµРЅРґР°С†РёСЏ"])

    gaps_by_source: dict[str, list[tuple[int, str, str]]] = {"РџСЂС‹РіРѕРј": [], "РљРѕСЂРѕРј": []}
    goals_by_source: dict[str, dict[str, str]] = {"РџСЂС‹РіРѕРј": {}, "РљРѕСЂРѕРј": {}}

    def append_block(source_name: str, rows: list[RatingRecord], ordered_types: list[str]) -> None:
        type_rows_map: dict[str, list[RatingRecord]] = defaultdict(list)
        for row in rows:
            type_rows_map[row.type_name].append(row)

        cohort_chars: set[str] | None = None
        if character == "Р¦Р°СЂРёС†Р°":
            cohort_chars = overall_cohort_characters(rows, place_from=7, place_to=12)

        for type_name in ordered_types:
            hero = best_record_for_character(rows, type_name, character)
            target = target_record_for_type(rows, type_name, character, cohort_chars=cohort_chars)
            threshold = min_achievement_for_type(rows, type_name)

            if hero is not None and target is not None and hero.achievement is not None and target.achievement is not None:
                gap = max(0, target.achievement - hero.achievement)
                gap_value: int | str | None = gap
            else:
                gap_value = None

            if type_name != "РћР±С‰РёР№ СЂРµР№С‚РёРЅРі" and target is not None and target.achievement is not None:
                if hero is not None and hero.achievement is not None:
                    priority_gap = max(0, target.achievement - hero.achievement)
                else:
                    priority_gap = target.achievement
                gaps_by_source[source_name].append((priority_gap, source_name, type_name))
                goal_compass = target.compass_level
                goal_time = format_time_ms(target.clear_time)
                current_compass = hero.compass_level if hero else None
                current_time = format_time_ms(hero.clear_time) if hero else None
                if current_compass is None or current_time is None:
                    current_text = "СЃРµР№С‡Р°СЃ: РЅРµС‚ РІ С‚РѕРї-100"
                else:
                    current_text = f"СЃРµР№С‡Р°СЃ: РєРѕРјРїР°СЃ {current_compass} / РІСЂРµРјСЏ {current_time}"
                if goal_compass is not None and goal_time is not None:
                    goals_by_source[source_name][type_name] = (
                        f"{current_text}; С†РµР»СЊ: РєРѕРјРїР°СЃ {goal_compass} / РІСЂРµРјСЏ {goal_time}"
                    )
                elif goal_compass is not None:
                    goals_by_source[source_name][type_name] = (
                        f"{current_text}; С†РµР»СЊ: РєРѕРјРїР°СЃ {goal_compass} / РІСЂРµРјСЏ -"
                    )
                elif goal_time is not None:
                    goals_by_source[source_name][type_name] = (
                        f"{current_text}; С†РµР»СЊ: РєРѕРјРїР°СЃ - / РІСЂРµРјСЏ {goal_time}"
                    )
                else:
                    goals_by_source[source_name][type_name] = f"{current_text}; С†РµР»СЊ: РєРѕРјРїР°СЃ - / РІСЂРµРјСЏ -"

            sheet.append(
                [
                    source_name,
                    type_name,
                    render_record_cell(hero),
                    render_record_cell(target),
                    gap_value,
                    recommendation_for_type(
                        type_name=type_name,
                        hero=hero,
                        top=target,
                        entry_threshold=threshold,
                        type_rows=type_rows_map[type_name],
                    ),
                ]
            )

    append_block("РџСЂС‹РіРѕРј", hpi_rows, hpi_types)
    append_block("РљРѕСЂРѕРј", astral_rows, astral_types)

    sheet.append([])
    sheet.append(["5 СЂРµРєРѕРјРµРЅРґР°С†РёР№ РґР»СЏ РџСЂС‹РіРѕРј", "Р¦РµР»СЊ"])
    prygom = sorted(gaps_by_source["РџСЂС‹РіРѕРј"], key=lambda x: x[0], reverse=True)[:5]
    if prygom:
        for idx, (gap, _, type_name) in enumerate(prygom, start=1):
            goal = goals_by_source["РџСЂС‹РіРѕРј"].get(type_name, "РєРѕРјРїР°СЃ - / РІСЂРµРјСЏ -")
            sheet.append([f"РџСЂС‹РіРѕРј #{idx}", f"{type_name}: {goal}"])
    else:
        sheet.append(["РџСЂС‹РіРѕРј", "РќРµС‚ РґР°РЅРЅС‹С… РґР»СЏ СЃСЂР°РІРЅРµРЅРёСЏ"])

    sheet.append([])
    sheet.append(["5 СЂРµРєРѕРјРµРЅРґР°С†РёР№ РґР»СЏ РљРѕСЂРѕРј", "Р¦РµР»СЊ"])
    korom = sorted(gaps_by_source["РљРѕСЂРѕРј"], key=lambda x: x[0], reverse=True)[:5]
    if korom:
        for idx, (gap, _, type_name) in enumerate(korom, start=1):
            goal = goals_by_source["РљРѕСЂРѕРј"].get(type_name, "РєРѕРјРїР°СЃ - / РІСЂРµРјСЏ -")
            sheet.append([f"РљРѕСЂРѕРј #{idx}", f"{type_name}: {goal}"])
    else:
        sheet.append(["РљРѕСЂРѕРј", "РќРµС‚ РґР°РЅРЅС‹С… РґР»СЏ СЃСЂР°РІРЅРµРЅРёСЏ"])

    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    header_row = 3 if character == "Р¦Р°СЂРёС†Р°" else 2
    for cell in sheet[header_row]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    for row_idx in range(1, sheet.max_row + 1):
        if sheet.cell(row_idx, 1).value in {"5 СЂРµРєРѕРјРµРЅРґР°С†РёР№ РґР»СЏ РџСЂС‹РіРѕРј", "5 СЂРµРєРѕРјРµРЅРґР°С†РёР№ РґР»СЏ РљРѕСЂРѕРј"}:
            for col in range(1, 3):
                sheet.cell(row_idx, col).font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)

    autosize_columns(sheet, min_first=24, min_other=28, max_width=180)
    sheet.column_dimensions["A"].width = max(sheet.column_dimensions["A"].width or 0, 48)
    sheet.column_dimensions["B"].width = max(sheet.column_dimensions["B"].width or 0, 126)
    sheet.column_dimensions["C"].width = max(sheet.column_dimensions["C"].width or 0, 36)
    sheet.column_dimensions["D"].width = max(sheet.column_dimensions["D"].width or 0, 36)
    sheet.column_dimensions["E"].width = max(sheet.column_dimensions["E"].width or 0, 28)
    sheet.column_dimensions["F"].width = max(sheet.column_dimensions["F"].width or 0, 90)
    apply_filter_and_freeze(sheet, header_row=header_row, freeze_cell=f"B{header_row + 1}")
    apply_layout(sheet)


def write_trends_sheet(workbook: Workbook, runs: list[dict[str, Any]], lookback: int = 8) -> None:
    sheet = create_sheet_safe(workbook, "РўСЂРµРЅРґС‹")
    sheet.append(["РСЃС‚РѕС‡РЅРёРє", "Р“СЂСѓРїРїР°", "Р РµР№С‚РёРЅРі (РїРѕСЃР»РµРґРЅРёРµ N)", "Р”РµР»СЊС‚Р°", "Р’СЂРµРјСЏ 25 (РїРѕСЃР»РµРґРЅРёРµ N)", "Р”РµР»СЊС‚Р° РІСЂРµРјРµРЅРё"])

    if not runs:
        sheet.append(["-", "-", "РќРµС‚ РґР°РЅРЅС‹С…", "-", "РќРµС‚ РґР°РЅРЅС‹С…", "-"])
    else:
        window = runs[-lookback:]
        current = window[-1]
        for source_name in ("РџСЂС‹РіРѕРј", "РљРѕСЂРѕРј"):
            source = current.get(source_name, {})
            groups = source.get("groups", {}) if isinstance(source.get("groups"), dict) else {}
            current_groups = normalize_groups_map_for_compare(groups)
            for group_name in sorted(current_groups.keys()):
                rating_points: list[int] = []
                time_points: list[int] = []
                for run in window:
                    run_source = run.get(source_name, {})
                    run_groups_raw = run_source.get("groups", {}) if isinstance(run_source.get("groups"), dict) else {}
                    run_groups = normalize_groups_map_for_compare(run_groups_raw)
                    group_map = run_groups.get(group_name, {})
                    overall = group_map.get("РћР±С‰РёР№ СЂРµР№С‚РёРЅРі")
                    if isinstance(overall, dict) and isinstance(overall.get("achievement"), int):
                        rating_points.append(overall["achievement"])

                    best_25: int | None = None
                    for type_payload in group_map.values():
                        if not isinstance(type_payload, dict):
                            continue
                        if type_payload.get("compass") == 25 and isinstance(type_payload.get("time"), int) and type_payload["time"] > 0:
                            if best_25 is None or type_payload["time"] < best_25:
                                best_25 = type_payload["time"]
                    if best_25 is not None:
                        time_points.append(best_25)

                rating_text = " -> ".join(str(v) for v in rating_points) if rating_points else "-"
                time_text = " -> ".join(format_time_ms(v) or "-" for v in time_points) if time_points else "-"
                rating_delta = rating_points[-1] - rating_points[0] if len(rating_points) >= 2 else None
                time_delta = time_points[-1] - time_points[0] if len(time_points) >= 2 else None

                sheet.append(
                    [
                        source_name,
                        group_name,
                        rating_text,
                        rating_delta if rating_delta is not None else "-",
                        time_text,
                        format_time_ms(abs(time_delta)) if time_delta is not None else "-",
                    ]
                )
                current_row = sheet.max_row
                if isinstance(rating_delta, int):
                    if rating_delta > 0:
                        for col in range(1, 7):
                            sheet.cell(row=current_row, column=col).fill = GROW_FILL
                    elif rating_delta < 0:
                        for col in range(1, 7):
                            sheet.cell(row=current_row, column=col).fill = ANOMALY_FILL

    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    autosize_columns(sheet, min_first=16, min_other=24, max_width=180)
    sheet.column_dimensions["C"].width = max(sheet.column_dimensions["C"].width or 0, 60)
    sheet.column_dimensions["E"].width = max(sheet.column_dimensions["E"].width or 0, 54)
    apply_filter_and_freeze(sheet, header_row=1, freeze_cell="B2")
    apply_layout(sheet)


def write_run_summary_sheet(
    workbook: Workbook,
    previous: dict[str, Any],
    current: dict[str, Any],
    changes: list[dict[str, Any]],
    odd_changes: list[dict[str, Any]],
    data_errors: list[dict[str, Any]],
) -> None:
    sheet = create_sheet_safe(workbook, "Сводка запуска")
    sheet.append(["Показатель", "Значение"])

    created = sum(1 for c in changes if fix_mojibake_text(as_str(c.get("change_type"))) == "Новая группа")
    removed = sum(1 for c in changes if fix_mojibake_text(as_str(c.get("change_type"))) == "Группа пропала")
    rating_deltas = [c["rating_delta"] for c in changes if isinstance(c["rating_delta"], int)]
    avg_delta = round(sum(rating_deltas) / len(rating_deltas), 2) if rating_deltas else 0
    top_up = sorted(
        [c for c in changes if isinstance(c["rating_delta"], int) and c["rating_delta"] > 0],
        key=lambda x: x["rating_delta"],
        reverse=True,
    )[:5]
    anomalies = (data_errors + odd_changes)[:5]

    prev_count = 0
    curr_count = 0
    for source_name in ("РџСЂС‹РіРѕРј", "РљРѕСЂРѕРј"):
        prev_source = previous.get(source_name, {})
        curr_source = current.get(source_name, {})
        prev_groups_raw = prev_source.get("groups", {}) if isinstance(prev_source.get("groups"), dict) else {}
        curr_groups_raw = curr_source.get("groups", {}) if isinstance(curr_source.get("groups"), dict) else {}
        prev_count += len(normalize_groups_map_for_compare(prev_groups_raw))
        curr_count += len(normalize_groups_map_for_compare(curr_groups_raw))

    summary_rows = [
        ("Дата запуска", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Групп в прошлом запуске", prev_count),
        ("Групп в текущем запуске", curr_count),
        ("Обычных изменений", len(changes)),
        ("Странных изменений", len(odd_changes)),
        ("Ошибок данных", len(data_errors)),
        ("Новых групп", created),
        ("Пропавших групп", removed),
        ("Средняя дельта рейтинга", avg_delta),
    ]
    for row in summary_rows:
        sheet.append(list(row))

    sheet.append([])
    sheet.append(["Топ-5 роста", "Источник / Группа / Остров / Дельта"])
    if top_up:
        for item in top_up:
            source = fix_mojibake_text(as_str(item.get("source")))
            group = fix_mojibake_text(as_str(item.get("group")))
            type_name = fix_mojibake_text(as_str(item.get("type_name")))
            sheet.append(["Рост", f"{source} / {group} / {type_name} / {item['rating_delta']}"])
    else:
        sheet.append(["Рост", "Нет данных"])

    sheet.append([])
    sheet.append(["Аномалии (должно быть 0)", "Источник / Группа / Остров / Тип"])
    if anomalies:
        for item in anomalies:
            source = fix_mojibake_text(as_str(item.get("source")))
            group = fix_mojibake_text(as_str(item.get("group")))
            type_name = fix_mojibake_text(as_str(item.get("type_name")))
            change_type = fix_mojibake_text(as_str(item.get("change_type")))
            sheet.append(["Аномалия", f"{source} / {group} / {type_name} / {change_type}"])
    else:
        sheet.append(["Аномалия", "Нет данных"])

    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    for row_idx in range(2, sheet.max_row + 1):
        title = sheet.cell(row=row_idx, column=1).value
        if title in {"Рост", "Аномалия"}:
            fill = GROW_FILL if title == "Рост" else ANOMALY_FILL
            for col in range(1, 3):
                sheet.cell(row=row_idx, column=col).fill = fill
    autosize_columns(sheet, min_first=28, min_other=48, max_width=220)
    apply_filter_and_freeze(sheet, header_row=1, freeze_cell="B2")
    apply_layout(sheet)


def write_recent_improvements_sheet(
    workbook: Workbook,
    runs: list[dict[str, Any]],
    source_name: str,
    sheet_title: str,
    max_records_per_pair: int = 5,
) -> None:
    sheet = create_sheet_safe(workbook, sheet_title)
    headers = ["Р“СЂСѓРїРїР°", "РћСЃС‚СЂРѕРІ"] + [f"#{i}" for i in range(1, max_records_per_pair + 1)]
    sheet.append(headers)

    improvements_by_pair: dict[tuple[str, str], list[tuple[str, str, str, int | None, int | None]]] = defaultdict(list)
    current_overall_by_group: dict[str, int] = {}
    if runs:
        latest_run = runs[-1]
        latest_source = latest_run.get(source_name, {}) if isinstance(latest_run, dict) else {}
        latest_groups_raw = latest_source.get("groups", {}) if isinstance(latest_source.get("groups"), dict) else {}
        latest_groups = normalize_groups_map_for_compare(latest_groups_raw)
        for group_name, type_map in latest_groups.items():
            overall = type_map.get("РћР±С‰РёР№ СЂРµР№С‚РёРЅРі") if isinstance(type_map, dict) else None
            if isinstance(overall, dict) and isinstance(overall.get("achievement"), int):
                current_overall_by_group[group_name] = overall["achievement"]

    for idx in range(1, len(runs)):
        prev_run = runs[idx - 1]
        curr_run = runs[idx]
        run_at = curr_run.get("generated_at") if isinstance(curr_run.get("generated_at"), str) else "-"

        prev_source = prev_run.get(source_name, {}) if isinstance(prev_run, dict) else {}
        curr_source = curr_run.get(source_name, {}) if isinstance(curr_run, dict) else {}

        prev_groups_raw = prev_source.get("groups", {}) if isinstance(prev_source.get("groups"), dict) else {}
        curr_groups_raw = curr_source.get("groups", {}) if isinstance(curr_source.get("groups"), dict) else {}
        prev_groups = normalize_groups_map_for_compare(prev_groups_raw)
        curr_groups = normalize_groups_map_for_compare(curr_groups_raw)

        for group_name in sorted(curr_groups.keys()):
            curr_types = curr_groups.get(group_name, {})
            prev_types = prev_groups.get(group_name, {})
            if not isinstance(curr_types, dict):
                continue

            for type_name, curr_val in curr_types.items():
                if type_name == "РћР±С‰РёР№ СЂРµР№С‚РёРЅРі" or not isinstance(curr_val, dict):
                    continue
                prev_val = prev_types.get(type_name)
                prev_val = prev_val if isinstance(prev_val, dict) else None

                prev_ach = prev_val.get("achievement") if prev_val else None
                curr_ach = curr_val.get("achievement")
                prev_time = prev_val.get("time") if prev_val else None
                curr_time = curr_val.get("time")

                rating_improved = isinstance(curr_ach, int) and (not isinstance(prev_ach, int) or curr_ach > prev_ach)
                time_improved_same_rating = (
                    isinstance(curr_ach, int)
                    and isinstance(prev_ach, int)
                    and curr_ach == prev_ach
                    and isinstance(prev_time, int)
                    and prev_time > 0
                    and isinstance(curr_time, int)
                    and curr_time > 0
                    and curr_time < prev_time
                )
                if not rating_improved and not time_improved_same_rating:
                    continue

                before_text = f"{prev_val.get('compass') or '-'} / {format_time_change(prev_time)}" if prev_val else "-"
                after_text = f"{curr_val.get('compass') or '-'} / {format_time_change(curr_time)}"

                rating_delta = None
                if isinstance(curr_ach, int) and isinstance(prev_ach, int):
                    rating_delta = curr_ach - prev_ach
                elif isinstance(curr_ach, int) and prev_ach is None:
                    rating_delta = curr_ach

                time_delta = None
                if isinstance(curr_time, int) and isinstance(prev_time, int) and prev_time > 0 and curr_time > 0:
                    time_delta = prev_time - curr_time

                improvements_by_pair[(group_name, type_name)].append((run_at, before_text, after_text, rating_delta, time_delta))

    if not improvements_by_pair:
        sheet.append(["-", "-", "РќРµС‚ СѓР»СѓС‡С€РµРЅРёР№", "-", "-", "-", "-"])
    else:
        sorted_pairs = sorted(
            improvements_by_pair.keys(),
            key=lambda pair: (
                -(current_overall_by_group.get(pair[0], -1)),
                pair[0].lower(),
                pair[1].lower(),
            ),
        )
        for (group_name, type_name) in sorted_pairs:
            events = improvements_by_pair[(group_name, type_name)][-max_records_per_pair:][::-1]
            compact_events: list[str] = []
            for run_at, before_text, after_text, rating_delta, time_delta in events:
                # Avoid duplicate columns when only rating changed but compass/time did not.
                if after_text not in compact_events:
                    compact_events.append(after_text)

            row: list[Any] = [group_name, type_name]
            for i in range(max_records_per_pair):
                row.append(compact_events[i] if i < len(compact_events) else "-")
            sheet.append(row)

    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    autosize_columns(sheet, min_first=28, min_other=34, max_width=140)
    sheet.column_dimensions["A"].width = max(sheet.column_dimensions["A"].width or 0, 34)
    sheet.column_dimensions["B"].width = max(sheet.column_dimensions["B"].width or 0, 40)
    for col in ("C", "D", "E", "F", "G"):
        sheet.column_dimensions[col].width = max(sheet.column_dimensions[col].width or 0, 46)
    apply_filter_and_freeze(sheet, header_row=1, freeze_cell="C2")
    apply_layout(sheet)


def _median_int(values: list[int]) -> int | None:
    if not values:
        return None
    sorted_vals = sorted(values)
    mid = len(sorted_vals) // 2
    if len(sorted_vals) % 2 == 1:
        return sorted_vals[mid]
    return (sorted_vals[mid - 1] + sorted_vals[mid]) // 2


def write_arena_top100_class_stats_sheet(
    workbook: Workbook,
    sheet_title: str,
    arena_top_rows: list[ArenaRecord],
) -> None:
    sheet = create_sheet_safe(workbook, sheet_title)
    headers = [
        "\u041a\u043b\u0430\u0441\u0441",
        "\u0412 \u0442\u043e\u043f-100",
        "\u0414\u043e\u043b\u044f, %",
        "\u041b\u0443\u0447\u0448\u0435\u0435 \u043c\u0435\u0441\u0442\u043e",
        "\u0421\u0440\u0435\u0434\u043d\u0435\u0435 \u043c\u0435\u0441\u0442\u043e",
        "\u0421\u0440\u0435\u0434\u043d\u0438\u0439 \u0440\u0435\u0439\u0442\u0438\u043d\u0433",
        "\u041c\u0430\u043a\u0441. \u0440\u0435\u0439\u0442\u0438\u043d\u0433",
        "\u0420\u0435\u0439\u0442\u0438\u043d\u0433 \u0432\u0445\u043e\u0434\u0430",
    ]

    sheet.append(headers)
    if not arena_top_rows:
        sheet.append(["-", 0, 0, "-", "-", "-", "-", "-"])
    else:
        total = len(arena_top_rows)
        class_counter = Counter(row.class_name for row in arena_top_rows if row.class_name)
        best_place_by_class: dict[str, int] = {}
        place_avg_by_class: dict[str, float] = {}
        avg_ach_by_class: dict[str, float] = {}
        max_ach_by_class: dict[str, int] = {}
        min_ach_by_class: dict[str, int] = {}

        for class_name in class_counter.keys():
            rows = [r for r in arena_top_rows if r.class_name == class_name]
            places = [r.place for r in rows if r.place is not None]
            achievements = [r.achievement for r in rows if r.achievement is not None]
            if places:
                best_place_by_class[class_name] = min(places)
                place_avg_by_class[class_name] = round(sum(places) / len(places), 2)
            if achievements:
                avg_ach_by_class[class_name] = round(sum(achievements) / len(achievements), 2)
                max_ach_by_class[class_name] = max(achievements)
                min_ach_by_class[class_name] = min(achievements)

        sorted_classes = sorted(
            class_counter.keys(),
            key=lambda name: (-class_counter[name], best_place_by_class.get(name, 10**9), name.lower()),
        )
        for class_name in sorted_classes:
            count = class_counter[class_name]
            share = round((count / total) * 100, 2) if total else 0
            sheet.append(
                [
                    class_name,
                    count,
                    share,
                    best_place_by_class.get(class_name, "-"),
                    place_avg_by_class.get(class_name, "-"),
                    avg_ach_by_class.get(class_name, "-"),
                    max_ach_by_class.get(class_name, "-"),
                    min_ach_by_class.get(class_name, "-"),
                ]
            )
        sheet.append(["\u0418\u0442\u043e\u0433\u043e", total, 100.0, "-", "-", "-", "-", "-"])

    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    autosize_columns(sheet, min_first=22, min_other=20, max_width=90)
    apply_filter_and_freeze(sheet, header_row=1, freeze_cell="B2")
    apply_layout(sheet)


def write_arena_class_details_sheet(
    workbook: Workbook,
    sheet_title: str,
    arena_top_rows: list[ArenaRecord],
    arena_rows_by_class: dict[int, list[ArenaRecord]],
) -> None:
    sheet = create_sheet_safe(workbook, sheet_title)
    headers = [
        "\u041a\u043b\u0430\u0441\u0441",
        "\u0417\u0430\u043f\u0438\u0441\u0435\u0439 \u0432 \u0440\u0435\u0439\u0442\u0438\u043d\u0433\u0435 \u043a\u043b\u0430\u0441\u0441\u0430",
        "\u0412 \u0442\u043e\u043f-100 \u043e\u0431\u0449\u0435\u0433\u043e",
        "\u0414\u043e\u043b\u044f \u0432 \u043e\u0431\u0449\u0435\u043c \u0442\u043e\u043f-100, %",
        "\u041b\u0443\u0447\u0448\u0438\u0439 \u0440\u0435\u0439\u0442\u0438\u043d\u0433",
        "\u0421\u0440\u0435\u0434\u043d\u0438\u0439 \u0440\u0435\u0439\u0442\u0438\u043d\u0433",
        "\u041c\u0435\u0434\u0438\u0430\u043d\u0430 \u0440\u0435\u0439\u0442\u0438\u043d\u0433\u0430",
        "\u0420\u0435\u0439\u0442\u0438\u043d\u0433 \u0432\u0445\u043e\u0434\u0430",
        "\u0422\u043e\u043f-1 \u043f\u0435\u0440\u0441\u043e\u043d\u0430\u0436",
        "\u0413\u0438\u043b\u044c\u0434\u0438\u044f \u0442\u043e\u043f-1",
    ]
    sheet.append(headers)
    for class_id, class_name in ARENA_CLASSES:
        if class_id == 0:
            continue
        rows = arena_rows_by_class.get(class_id, [])
        display_name = rows[0].class_name if rows else class_name
        top_count = sum(1 for r in arena_top_rows if r.class_name == display_name)
        top_share = round((top_count / ARENA_TOP_LIMIT) * 100, 2)

        achievements = [r.achievement for r in rows if r.achievement is not None]
        top_record = min(
            rows,
            key=lambda r: (r.place if r.place is not None else 10**9, -(r.achievement or 0), r.character.lower()),
            default=None,
        )

        sheet.append(
            [
                display_name,
                len(rows),
                top_count,
                top_share,
                max(achievements) if achievements else "-",
                round(sum(achievements) / len(achievements), 2) if achievements else "-",
                _median_int(achievements) if achievements else "-",
                min(achievements) if achievements else "-",
                top_record.character if top_record else "-",
                short_guild_name(top_record.guild) if top_record else "-",
            ]
        )

    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    autosize_columns(sheet, min_first=22, min_other=24, max_width=96)
    apply_filter_and_freeze(sheet, header_row=1, freeze_cell="B2")
    apply_layout(sheet)


def write_arena_top5_per_class_sheet(
    workbook: Workbook,
    sheet_title: str,
    arena_rows_by_class: dict[int, list[ArenaRecord]],
) -> None:
    sheet = create_sheet_safe(workbook, sheet_title)
    headers = [
        "Класс",
        "Топ 1",
        "Топ 2",
        "Топ 3",
        "Топ 4",
        "Топ 5",
    ]
    sheet.append(headers)

    rows_written = 0
    for class_id, class_name in ARENA_CLASSES:
        if class_id == 0:
            continue
        class_rows = arena_rows_by_class.get(class_id, [])
        display_name = class_rows[0].class_name if class_rows else class_name
        top_rows = sorted(
            class_rows,
            key=lambda row: (
                row.place if row.place is not None else 10**9,
                -(row.achievement or 0),
                row.character.lower(),
            ),
        )[:5]

        row_values = [display_name]
        for row in top_rows:
            character = row.character or "-"
            guild = row.guild or "-"
            waves = str(row.achievement) if row.achievement is not None else "-"
            cell_value = f"{character}({guild}) - {waves}"
            row_values.append(cell_value)
        while len(row_values) < len(headers):
            row_values.append("-")
        sheet.append(row_values)
        current_row = sheet.max_row
        for col_idx in range(2, len(headers) + 1):
            cell_value = sheet.cell(row=current_row, column=col_idx).value
            if not isinstance(cell_value, str) or cell_value == "-":
                continue
            character_name = cell_value.split("(", 1)[0].strip()
            if character_name in ARENA_HIGHLIGHT_CHARACTERS:
                sheet.cell(row=current_row, column=col_idx).fill = GROW_FILL
        rows_written += 1

    if rows_written == 0:
        sheet.append(["-", "-", "-", "-", "-", "-"])

    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)

    autosize_columns(sheet, min_first=18, min_other=34, max_width=90)
    for col_idx in range(1, sheet.max_column + 1):
        col_letter = get_column_letter(col_idx)
        current_width = sheet.column_dimensions[col_letter].width or 0
        sheet.column_dimensions[col_letter].width = current_width * 2
    apply_filter_and_freeze(sheet, header_row=1, freeze_cell="B2")
    apply_layout(sheet)


def write_toc_sheet(workbook: Workbook, sheet_descriptions: dict[str, str] | None = None) -> str:
    toc = create_sheet_safe(workbook, "РћРіР»Р°РІР»РµРЅРёРµ")
    toc_title = toc.title
    toc.append(["Р›РёСЃС‚", "РќР°Р·РЅР°С‡РµРЅРёРµ"])

    descriptions = sheet_descriptions or {}
    for title in sorted(name for name in workbook.sheetnames if name != toc_title):
        toc.append([title, descriptions.get(title, "")])
        cell = toc.cell(row=toc.max_row, column=1)
        cell.hyperlink = f"#{quote_sheetname(title)}!A1"
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, underline="single", color="0563C1")

    for cell in toc[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    autosize_columns(toc, min_first=24, min_other=36, max_width=180)
    apply_filter_and_freeze(toc, header_row=1, freeze_cell="A2")
    apply_layout(toc)
    return toc_title


def fix_workbook_text(workbook: Workbook) -> None:
    updated_titles: set[str] = set()
    title_map: dict[str, str] = {}

    for sheet in workbook.worksheets:
        original_title = sheet.title
        fixed_title = fix_mojibake_text(original_title)
        fixed_title = _safe_sheet_title(fixed_title, updated_titles)
        updated_titles.add(fixed_title)
        title_map[original_title] = fixed_title
        if fixed_title != original_title:
            sheet.title = fixed_title

    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    cell.value = fix_mojibake_text(cell.value)
                if cell.hyperlink is not None:
                    if cell.hyperlink.target:
                        cell.hyperlink.target = fix_mojibake_text(cell.hyperlink.target)
                    if cell.hyperlink.location:
                        cell.hyperlink.location = fix_mojibake_text(cell.hyperlink.location)


def yandex_disk_api_url(endpoint: str, **params: Any) -> str:
    query = urlencode({k: v for k, v in params.items() if v is not None and v != ""})
    if not query:
        return YANDEX_DISK_API_BASE + endpoint
    return f"{YANDEX_DISK_API_BASE}{endpoint}?{query}"


def ensure_yandex_disk_folder(token: str, remote_file_path: str, timeout: int) -> None:
    folder = remote_file_path.rsplit("/", 1)[0].strip()
    if not folder:
        return
    headers = {"Authorization": f"OAuth {token}"}
    try:
        request_json(
            yandex_disk_api_url("/resources", path=folder),
            timeout=timeout,
            headers=headers,
        )
        return
    except HTTPError as exc:
        if exc.code != 404:
            raise

    try:
        request_json(
            yandex_disk_api_url("/resources", path=folder),
            timeout=timeout,
            method="PUT",
            headers=headers,
        )
    except HTTPError as exc:
        if exc.code != 409:
            raise


def upload_file_to_yandex_disk(local_path: Path, remote_file_path: str, token: str, timeout: int) -> str | None:
    headers = {"Authorization": f"OAuth {token}"}
    ensure_yandex_disk_folder(token=token, remote_file_path=remote_file_path, timeout=timeout)

    upload_meta = request_json(
        yandex_disk_api_url("/resources/upload", path=remote_file_path, overwrite="true"),
        timeout=timeout,
        headers=headers,
    )
    if not isinstance(upload_meta, dict) or not isinstance(upload_meta.get("href"), str):
        raise RuntimeError("Не удалось получить ссылку загрузки Яндекс Диска.")

    upload_req = Request(
        upload_meta["href"],
        data=local_path.read_bytes(),
        method="PUT",
        headers={"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
    )
    with urlopen(upload_req, timeout=timeout) as response:
        if response.status not in (201, 202):
            raise RuntimeError(f"Ошибка загрузки на Яндекс Диск: HTTP {response.status}")

    try:
        request_json(
            yandex_disk_api_url("/resources/publish", path=remote_file_path),
            timeout=timeout,
            method="PUT",
            headers=headers,
        )
    except HTTPError as exc:
        if exc.code != 409:
            raise

    meta = request_json(
        yandex_disk_api_url("/resources", path=remote_file_path),
        timeout=timeout,
        headers=headers,
    )
    if isinstance(meta, dict):
        public_url = meta.get("public_url")
        if isinstance(public_url, str) and public_url.strip():
            return public_url.strip()
    return None


def write_legend_sheet(workbook: Workbook) -> None:
    sheet = create_sheet_safe(workbook, "Р›РµРіРµРЅРґР°")
    sheet.append(["РћР±РѕР·РЅР°С‡РµРЅРёРµ", "РћРїРёСЃР°РЅРёРµ"])
    sheet.append(["РџСѓСЃС‚Рѕ", "Р“СЂСѓРїРїР° РѕС‚СЃСѓС‚СЃС‚РІСѓРµС‚ РІ С†РµР»РµРІРѕРј С‚РѕРїРµ (РѕР±С‰РёР№: С‚РѕРї-30, РѕСЃС‚СЂРѕРІР°: С‚РѕРї-100)"])
    sheet.append(["Р—РµР»С‘РЅС‹Р№ С„РѕРЅ", "Р›СѓС‡С€РµРµ Р·РЅР°С‡РµРЅРёРµ РІ СЃС‚СЂРѕРєРµ (РґР»СЏ РІСЂРµРјРµРЅРё - РјРёРЅРёРјР°Р»СЊРЅРѕРµ)"])
    sheet.append(["РџРµСЂСЃРёРєРѕРІС‹Р№ С„РѕРЅ", "Р“СЂСѓРїРїР° РЅРµ РІС…РѕРґРёС‚ РІ С‚РѕРї-24 РѕР±С‰РµРіРѕ СЂРµР№С‚РёРЅРіР°"])
    sheet.append(["Р¤РѕСЂРјР°С‚ РІСЂРµРјРµРЅРё", "РјРёРЅСѓС‚С‹:СЃРµРєСѓРЅРґС‹"])
    sheet.append(["РћСЃРЅРѕРІРЅС‹Рµ Р»РёСЃС‚С‹", "РџСЂС‹РіРѕРј/РљРѕСЂРѕРј = РљРѕРјРїР°СЃ - Р’СЂРµРјСЏ, *_СЂРµР№С‚РёРЅРі = С†РёС„СЂС‹ СЂРµР№С‚РёРЅРіР°"])
    sheet.append(["РЎРёРЅРёР№ С„РѕРЅ", "РќРѕРІР°СЏ РіСЂСѓРїРїР° РІ Р»РёСЃС‚Рµ РёР·РјРµРЅРµРЅРёР№"])
    sheet.append(["РЎРµСЂС‹Р№ С„РѕРЅ", "Р“СЂСѓРїРїР° РїСЂРѕРїР°Р»Р° РѕС‚РЅРѕСЃРёС‚РµР»СЊРЅРѕ РїСЂРѕС€Р»РѕРіРѕ Р·Р°РїСѓСЃРєР°"])
    sheet.append(["Р–РµР»С‚С‹Р№ С„РѕРЅ", "РЎС‚СЂР°РЅРЅРѕРµ/Р°РЅРѕРјР°Р»СЊРЅРѕРµ РёР·РјРµРЅРµРЅРёРµ"])
    sheet.append(["РљСЂР°СЃРЅС‹Р№ С„РѕРЅ", "РЎРЅРёР¶РµРЅРёРµ СЂРµР№С‚РёРЅРіР°"])
    for cell in sheet[1]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    autosize_columns(sheet, min_first=26, min_other=46, max_width=220)
    current_b = sheet.column_dimensions["B"].width or 0
    sheet.column_dimensions["B"].width = max(current_b, 120)
    apply_filter_and_freeze(sheet, header_row=1, freeze_cell="B2")
    apply_layout(sheet)


def write_meta_sheet(
    workbook: Workbook,
    shard_id: int,
    shard_name: str,
    total_rows: int,
    total_chars: int,
    compared_with_run_at: str | None,
    arena_rows_total: int = 0,
) -> None:
    sheet = create_sheet_safe(workbook, "РњРµС‚Р°РґР°РЅРЅС‹Рµ")
    rows = [
        ("РЎРµСЂРІРµСЂ", shard_name),
        ("ID СЃРµСЂРІРµСЂР°", shard_id),
        ("Р”Р°С‚Р° РІС‹РіСЂСѓР·РєРё", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("РЎСЂР°РІРЅРµРЅРёРµ СЃ Р·Р°РїСѓСЃРєРѕРј", compared_with_run_at or "РќРµС‚ РїСЂРµРґС‹РґСѓС‰РµРіРѕ Р·Р°РїСѓСЃРєР°"),
        ("СЃС‚РѕС‡РЅРёРє 1", "https://allods.ru/ratings/#/hpi/"),
        ("СЃС‚РѕС‡РЅРёРє 2", "https://allods.ru/ratings/#/hpi-astral/"),
        ("СЃС‚РѕС‡РЅРёРє 3", "https://allods.ru/ratings/#/arena/"),
        ("РћРіСЂР°РЅРёС‡РµРЅРёРµ", f"РћР±С‰РёР№ СЂРµР№С‚РёРЅРі: С‚РѕРї-{OVERALL_TOP_LIMIT}; РѕСЃС‚СЂРѕРІР°: С‚РѕРї-{ISLAND_TOP_LIMIT}"),
        ("РљРѕР»РёС‡РµСЃС‚РІРѕ Р·Р°РїРёСЃРµР№", total_rows),
        ("Р—Р°РїРёСЃРµР№ Р°СЂРµРЅС‹", arena_rows_total),
        ("РљРѕР»РёС‡РµСЃС‚РІРѕ РїРµСЂСЃРѕРЅР°Р¶РµР№", total_chars),
    ]
    for k, v in rows:
        sheet.append([k, v])
    for cell in sheet["A"]:
        cell.font = Font(name=FONT_NAME, size=FONT_SIZE, bold=True)
    autosize_columns(sheet)
    apply_filter_and_freeze(sheet, header_row=1, freeze_cell="B2")
    apply_layout(sheet)


def build_source_sheets(
    workbook: Workbook,
    base_sheet_name: str,
    rows: list[RatingRecord],
    ordered_types: list[str],
) -> list[Entity]:
    characters = build_character_entities(rows, ordered_types)
    groups = build_group_entities(characters, ordered_types)

    write_metric_matrix_sheet(
        workbook=workbook,
        sheet_title=base_sheet_name,
        entities=groups,
        rows=rows,
        ordered_types=ordered_types,
        mode="compass_time",
    )
    write_metric_matrix_sheet(
        workbook=workbook,
        sheet_title=f"{base_sheet_name}_СЂРµР№С‚РёРЅРі",
        entities=groups,
        rows=rows,
        ordered_types=ordered_types,
        mode="rating",
    )
    write_group_description_sheet(
        workbook=workbook,
        sheet_title=f"{base_sheet_name}_РіСЂСѓРїРїС‹",
        entities=groups,
    )
    return groups


def generate_report(
    config: ShardConfig,
    output_path: Path,
    state_path: Path,
    timeout: int,
    yadisk_path: str,
    yadisk_token_env: str,
) -> None:

    try:
        hpi_rows = fetch_source_records(
            source="hpi",
            endpoint_template=HPI_ENDPOINT,
            type_pairs=HPI_TYPES,
            shard_id=config.shard_id,
            timeout=timeout,
            default_shard_name=config.name,
        )
        astral_rows = fetch_source_records(
            source="hpi-astral",
            endpoint_template=HPI_ASTRAL_ENDPOINT,
            type_pairs=HPI_ASTRAL_TYPES,
            shard_id=config.shard_id,
            timeout=timeout,
            default_shard_name=config.name,
        )
        arena_top_rows, arena_rows_by_class = fetch_arena_records(
            shard_id=config.shard_id,
            timeout=timeout,
            default_shard_name=config.name,
        )
    except HTTPError as exc:
        raise cli_exit(f"РћС€РёР±РєР° HTTP: {exc.code} {exc.reason}") from exc
    except URLError as exc:
        raise cli_exit(f"РћС€РёР±РєР° СЃРµС‚Рё: {exc.reason}") from exc
    except json.JSONDecodeError as exc:
        raise cli_exit(f"РћС€РёР±РєР° РїР°СЂСЃРёРЅРіР° JSON: {exc}") from exc
    except Exception as exc:
        raise cli_exit(f"РќРµ СѓРґР°Р»РѕСЃСЊ РїРѕР»СѓС‡РёС‚СЊ РґР°РЅРЅС‹Рµ: {exc}") from exc

    all_rows = hpi_rows + astral_rows
    if not all_rows:
        raise cli_exit("РџРѕР»СѓС‡РµРЅС‹ РїСѓСЃС‚С‹Рµ РґР°РЅРЅС‹Рµ СЂРµР№С‚РёРЅРіРѕРІ.")

    workbook = Workbook()
    workbook.remove(workbook.active)

    prygom_groups = build_source_sheets(
        workbook=workbook,
        base_sheet_name="РџСЂС‹РіРѕРј",
        rows=hpi_rows,
        ordered_types=[name for _, name in HPI_TYPES],
    )
    korom_groups = build_source_sheets(
        workbook=workbook,
        base_sheet_name="РљРѕСЂРѕРј",
        rows=astral_rows,
        ordered_types=[name for _, name in HPI_ASTRAL_TYPES],
    )
    for sheet_title, character in FOCUS_CHARACTERS:
        write_focus_sheet(
            workbook=workbook,
            sheet_title=sheet_title,
            character=character,
            hpi_rows=hpi_rows,
            hpi_types=[name for _, name in HPI_TYPES],
            astral_rows=astral_rows,
            astral_types=[name for _, name in HPI_ASTRAL_TYPES],
        )

    write_arena_top100_class_stats_sheet(
        workbook=workbook,
        sheet_title="РђСЂРµРЅР°_С‚РѕРї100_РєР»Р°СЃСЃС‹",
        arena_top_rows=arena_top_rows,
    )
    write_arena_top5_per_class_sheet(
        workbook=workbook,
        sheet_title="Arena_top5_by_class",
        arena_rows_by_class=arena_rows_by_class,
    )

    current_state = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "РџСЂС‹РіРѕРј": build_snapshot_for_source("РџСЂС‹РіРѕРј", prygom_groups, [name for _, name in HPI_TYPES]),
        "РљРѕСЂРѕРј": build_snapshot_for_source("РљРѕСЂРѕРј", korom_groups, [name for _, name in HPI_ASTRAL_TYPES]),
    }

    history_runs = load_state_history(state_path)
    previous_state = history_runs[-1] if history_runs else {}
    previous_run_at = previous_state.get("generated_at") if isinstance(previous_state, dict) else None
    changes, odd_changes, data_errors = write_changes_sheets_with_odd(workbook, previous_state, current_state)
    write_run_summary_sheet(workbook, previous_state, current_state, changes, odd_changes, data_errors)
    runs_with_current = history_runs + [current_state]
    write_recent_improvements_sheet(
        workbook=workbook,
        runs=runs_with_current,
        source_name="РџСЂС‹РіРѕРј",
        sheet_title="РЈР»СѓС‡С€РµРЅРёСЏ_РџСЂС‹РіРѕРј",
        max_records_per_pair=5,
    )
    write_recent_improvements_sheet(
        workbook=workbook,
        runs=runs_with_current,
        source_name="РљРѕСЂРѕРј",
        sheet_title="РЈР»СѓС‡С€РµРЅРёСЏ_РљРѕСЂРѕРј",
        max_records_per_pair=5,
    )
    save_state(state_path, current_state)

    write_legend_sheet(workbook)

    shard_name = next((row.shard for row in all_rows if row.shard), config.name)
    write_meta_sheet(
        workbook=workbook,
        shard_id=config.shard_id,
        shard_name=shard_name,
        total_rows=len(all_rows),
        total_chars=len({r.character for r in all_rows if r.character}),
        compared_with_run_at=previous_run_at,
        arena_rows_total=len(arena_top_rows),
    )

    toc_title = write_toc_sheet(
        workbook,
        sheet_descriptions={
            "РР·РјРµРЅРµРЅРёСЏ": "РћР±С‹С‡РЅС‹Рµ РёР·РјРµРЅРµРЅРёСЏ РјРµР¶РґСѓ Р·Р°РїСѓСЃРєР°РјРё",
            "РЎС‚СЂР°РЅРЅС‹Рµ РёР·РјРµРЅРµРЅРёСЏ": "РџРѕРґРѕР·СЂРёС‚РµР»СЊРЅС‹Рµ РёР·РјРµРЅРµРЅРёСЏ Рё Р°РЅРѕРјР°Р»РёРё",
            "РЎРІРѕРґРєР° Р·Р°РїСѓСЃРєР°": "РС‚РѕРіРё С‚РµРєСѓС‰РµРіРѕ Р·Р°РїСѓСЃРєР° Рё С‚РѕРї-РёР·РјРµРЅРµРЅРёР№",
            "РЈР»СѓС‡С€РµРЅРёСЏ_РџСЂС‹РіРѕРј": "РџРѕСЃР»РµРґРЅРёРµ СѓР»СѓС‡С€РµРЅРёСЏ РіСЂСѓРїРї Рё РѕСЃС‚СЂРѕРІРѕРІ (РґРѕ 5 Р·Р°РїРёСЃРµР№)",
            "РЈР»СѓС‡С€РµРЅРёСЏ_РљРѕСЂРѕРј": "РџРѕСЃР»РµРґРЅРёРµ СѓР»СѓС‡С€РµРЅРёСЏ РіСЂСѓРїРї Рё РѕСЃС‚СЂРѕРІРѕРІ (РґРѕ 5 Р·Р°РїРёСЃРµР№)",
            "РђСЂРµРЅР°_С‚РѕРї100_РєР»Р°СЃСЃС‹": "РЎС‚Р°С‚РёСЃС‚РёРєР° РєР»Р°СЃСЃРѕРІ РІ С‚РѕРї-100 РѕР±С‰РµРіРѕ СЂРµР№С‚РёРЅРіР° Р°СЂРµРЅС‹",
            "Arena_top5_by_class": "Топ-5 персонажей по каждому классу: ник, гильдия, пройдено волн",
        },
    )
    for sheet_name in workbook.sheetnames:
        if sheet_name == toc_title:
            continue
        add_back_to_toc_link(workbook[sheet_name], toc_title)

    fix_workbook_text(workbook)
    workbook.save(output_path)
    yandex_public_url: str | None = None
    if yadisk_path:
        token = os.environ.get(yadisk_token_env, "").strip()
        if not token:
            raise cli_exit(
                f"Не задан OAuth-токен Яндекс Диска. Укажите переменную окружения {yadisk_token_env}."
            )
        try:
            yandex_public_url = upload_file_to_yandex_disk(
                local_path=output_path,
                remote_file_path=yadisk_path,
                token=token,
                timeout=timeout,
            )
        except HTTPError as exc:
            raise cli_exit(f"Ошибка Яндекс Диска: {exc.code} {exc.reason}") from exc
        except URLError as exc:
            raise cli_exit(f"Ошибка сети Яндекс Диска: {exc.reason}") from exc

    cli_print(f"РЎРѕС…СЂР°РЅРµРЅРѕ: {output_path.resolve()}")
    cli_print(f"Р—Р°РїРёСЃРµР№: {len(all_rows)}")
    cli_print(f"РџРµСЂСЃРѕРЅР°Р¶РµР№: {len({r.character for r in all_rows if r.character})}")
    if yandex_public_url:
        cli_print(f"Yandex Disk URL: {yandex_public_url}")


def main() -> None:
    args = parse_args()
    if args.shard_id is None:
        configs = SHARD_CONFIGS
        if args.output or args.state_file or args.yadisk_path:
            raise cli_exit(
                "\u041f\u0430\u0440\u0430\u043c\u0435\u0442\u0440\u044b --output, --state-file \u0438 --yadisk-path \u043c\u043e\u0436\u043d\u043e \u0438\u0441\u043f\u043e\u043b\u044c\u0437\u043e\u0432\u0430\u0442\u044c \u0442\u043e\u043b\u044c\u043a\u043e \u0441 --shard-id."
            )
    else:
        config = next((item for item in SHARD_CONFIGS if item.shard_id == args.shard_id), None)
        if config is None:
            raise cli_exit(f"\u041d\u0435\u0438\u0437\u0432\u0435\u0441\u0442\u043d\u044b\u0439 ID \u0441\u0435\u0440\u0432\u0435\u0440\u0430: {args.shard_id}.")
        configs = (config,)

    for config in configs:
        output_path = Path(args.output) if args.output else Path(config.output_name)
        state_path = Path(args.state_file) if args.state_file else Path(config.state_name)
        yadisk_path = args.yadisk_path
        if args.upload_to_yadisk and not yadisk_path:
            yadisk_path = config.yadisk_path
        generate_report(
            config=config,
            output_path=output_path,
            state_path=state_path,
            timeout=args.timeout,
            yadisk_path=yadisk_path,
            yadisk_token_env=args.yadisk_token_env,
        )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(human_text("РћСЃС‚Р°РЅРѕРІР»РµРЅРѕ РїРѕР»СЊР·РѕРІР°С‚РµР»РµРј."))

