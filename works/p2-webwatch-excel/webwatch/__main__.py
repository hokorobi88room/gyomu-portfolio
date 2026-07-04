"""エントリポイント: python -m webwatch [config.yaml]"""
from __future__ import annotations

import sys

from .config import ConfigError, load_config
from .differ import SnapshotStore, diff_rows
from .excel_out import write_report
from .extractor import extract
from .fetcher import FetchError, Fetcher, logger, setup_logging


def main(argv: list[str] | None = None) -> int:
    setup_logging()
    args = argv if argv is not None else sys.argv[1:]
    config_path = args[0] if args else "config.yaml"

    try:
        cfg = load_config(config_path)
    except ConfigError as e:
        logger.error(f"config error: {e}")
        return 2

    fetcher = Fetcher(cfg.user_agent, cfg.interval_seconds, cfg.timeout_seconds)
    store = SnapshotStore(cfg.output_dir / "snapshots")

    results: dict[str, list[dict[str, str]]] = {}
    diffs = {}
    failures = 0

    for target in cfg.targets:
        try:
            html = fetcher.fetch(target.url)
            rows = extract(html, target)
            previous = store.load(target.name)
            diffs[target.name] = diff_rows(previous, rows)
            results[target.name] = rows
            store.save(target.name, rows)
            logger.info(
                f"{target.name}: {len(rows)} rows "
                f"(+{len(diffs[target.name].added)} / -{len(diffs[target.name].removed)})"
            )
        except FetchError as e:
            # 1対象の失敗で全体を止めない。ただし必ず記録して失敗数に数える
            logger.error(f"{target.name}: {e}")
            failures += 1

    if not results:
        logger.error("全対象の取得に失敗しました")
        return 1

    path = write_report(results, diffs, cfg.output_dir)
    logger.info(f"report written: {path}")

    changed = sum(1 for d in diffs.values() if d.has_changes)
    print(f"\n完了: {len(results)}対象を取得(失敗{failures})/ 変化あり {changed}件 → {path}")
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
