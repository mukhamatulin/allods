from __future__ import annotations

import csv
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "kompas_statistik.txt"
OUT_DIR = ROOT / "kompas_stats"

HEADER_RE = re.compile(r"^(\d{2}\.\d{2}\.\d{4})\s+-\s+(.+?):?\s*$")
DEATH_RE = re.compile(r"^(\d{1,2}:\d{2}(?::\d{2})?)\s+-\s+(.+?)\s*$")


@dataclass(frozen=True)
class DeathEvent:
    date: str
    stream_url: str
    video_time: str
    players: tuple[str, ...]
    island: str
    reason: str
    source_line: int


def split_players(value: str) -> tuple[str, ...]:
    return tuple(part.strip() for part in re.split(r"\s+и\s+", value) if part.strip())


def parse_events(source: Path) -> tuple[list[DeathEvent], list[str]]:
    events: list[DeathEvent] = []
    warnings: list[str] = []
    current_date = ""
    current_url = ""

    for line_number, raw_line in enumerate(source.read_text(encoding="utf-8-sig").splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue

        header = HEADER_RE.match(line)
        if header:
            current_date, current_url = header.groups()
            current_url = current_url.rstrip(":")
            continue

        death = DEATH_RE.match(line)
        if not death:
            warnings.append(f"Строка {line_number}: не удалось разобрать запись: {raw_line}")
            continue

        if not current_date or not current_url:
            warnings.append(f"Строка {line_number}: запись смерти идет до заголовка стрима: {raw_line}")
            continue

        video_time, payload = death.groups()
        parts = [part.strip() for part in payload.split(",", maxsplit=2)]
        if len(parts) != 3 or not all(parts):
            warnings.append(f"Строка {line_number}: ожидался формат 'имя, остров, причина': {raw_line}")
            continue

        players_raw, island, reason = parts
        players = split_players(players_raw)
        if not players:
            warnings.append(f"Строка {line_number}: не найдено имя игрока: {raw_line}")
            continue

        events.append(
            DeathEvent(
                date=current_date,
                stream_url=current_url,
                video_time=video_time,
                players=players,
                island=island,
                reason=reason,
                source_line=line_number,
            )
        )

    return events, warnings


def write_counter_csv(path: Path, headers: tuple[str, str], counter: Counter[str]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, delimiter=";")
        writer.writerow(headers)
        for name, count in counter.most_common():
            writer.writerow([name, count])


def plural_deaths(count: int) -> str:
    if count % 10 == 1 and count % 100 != 11:
        return "смерть"
    if 2 <= count % 10 <= 4 and not 12 <= count % 100 <= 14:
        return "смерти"
    return "смертей"


def write_summary(
    events: list[DeathEvent],
    warnings: list[str],
    player_counter: Counter[str],
    island_counter: Counter[str],
    reason_counter: Counter[str],
    stream_counter: Counter[str],
) -> None:
    total_deaths = sum(player_counter.values())
    multi_death_events = sum(1 for event in events if len(event.players) > 1)
    leader_count = player_counter.most_common(1)[0][1] if player_counter else 0
    leaders = [name for name, count in player_counter.items() if count == leader_count]

    lines = [
        "# Статистика смертей компаса",
        "",
        f"Источник: `{SOURCE.name}`",
        f"Всего стримов: {len(stream_counter)}",
        f"Всего записей смертей: {len(events)}",
        f"Всего индивидуальных смертей: {total_deaths}",
        f"Записей с несколькими умершими: {multi_death_events}",
        "",
        "## Кто умер больше всего",
    ]

    if leaders:
        joined_leaders = ", ".join(leaders)
        lines.append(f"Лидер: **{joined_leaders}** — {leader_count} {plural_deaths(leader_count)}.")
    else:
        lines.append("Нет данных.")

    lines.extend(["", "## Рейтинг по игрокам", "", "| Место | Игрок | Смерти |", "|---:|---|---:|"])
    for index, (player, count) in enumerate(player_counter.most_common(), start=1):
        lines.append(f"| {index} | {player} | {count} |")

    lines.extend(["", "## По островам", "", "| Остров | Смерти |", "|---|---:|"])
    for island, count in island_counter.most_common():
        lines.append(f"| {island} | {count} |")

    lines.extend(["", "## По причинам", "", "| Причина | Смерти |", "|---|---:|"])
    for reason, count in reason_counter.most_common():
        lines.append(f"| {reason} | {count} |")

    if warnings:
        lines.extend(["", "## Предупреждения разбора", ""])
        lines.extend(f"- {warning}" for warning in warnings)

    (OUT_DIR / "summary.md").write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    OUT_DIR.mkdir(exist_ok=True)
    events, warnings = parse_events(SOURCE)

    player_counter: Counter[str] = Counter()
    island_counter: Counter[str] = Counter()
    reason_counter: Counter[str] = Counter()
    stream_counter: Counter[str] = Counter()
    player_island_counter: dict[str, Counter[str]] = defaultdict(Counter)
    player_reason_counter: dict[str, Counter[str]] = defaultdict(Counter)

    parsed_rows: list[list[str | int]] = []
    for event in events:
        stream_counter[f"{event.date} - {event.stream_url}"] += len(event.players)
        for player in event.players:
            player_counter[player] += 1
            island_counter[event.island] += 1
            reason_counter[event.reason] += 1
            player_island_counter[player][event.island] += 1
            player_reason_counter[player][event.reason] += 1
            parsed_rows.append(
                [
                    event.date,
                    event.stream_url,
                    event.video_time,
                    player,
                    event.island,
                    event.reason,
                    event.source_line,
                ]
            )

    write_counter_csv(OUT_DIR / "deaths_by_player.csv", ("Игрок", "Смерти"), player_counter)
    write_counter_csv(OUT_DIR / "deaths_by_island.csv", ("Остров", "Смерти"), island_counter)
    write_counter_csv(OUT_DIR / "deaths_by_reason.csv", ("Причина", "Смерти"), reason_counter)
    write_counter_csv(OUT_DIR / "deaths_by_stream.csv", ("Стрим", "Смерти"), stream_counter)

    with (OUT_DIR / "parsed_deaths.csv").open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, delimiter=";")
        writer.writerow(["Дата", "Стрим", "Время", "Игрок", "Остров", "Причина", "Строка источника"])
        writer.writerows(parsed_rows)

    with (OUT_DIR / "player_details.csv").open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file, delimiter=";")
        writer.writerow(["Игрок", "Всего смертей", "Топ остров", "Топ причина"])
        for player, count in player_counter.most_common():
            top_island, top_island_count = player_island_counter[player].most_common(1)[0]
            top_reason, top_reason_count = player_reason_counter[player].most_common(1)[0]
            writer.writerow(
                [
                    player,
                    count,
                    f"{top_island} ({top_island_count})",
                    f"{top_reason} ({top_reason_count})",
                ]
            )

    if warnings:
        (OUT_DIR / "parse_warnings.txt").write_text("\n".join(warnings) + "\n", encoding="utf-8")
    else:
        warning_file = OUT_DIR / "parse_warnings.txt"
        if warning_file.exists():
            warning_file.unlink()

    write_summary(events, warnings, player_counter, island_counter, reason_counter, stream_counter)


if __name__ == "__main__":
    main()
