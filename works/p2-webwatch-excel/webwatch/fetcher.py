"""HTTP取得(robots.txt尊重・間隔制御・UA明示・タイムアウト)。"""
from __future__ import annotations

import json
import logging
import sys
import time
import urllib.robotparser
from urllib.parse import urlparse

import requests

logger = logging.getLogger("webwatch")


def setup_logging() -> None:
    """構造化ログ(JSON 1行)を標準出力へ。"""
    handler = logging.StreamHandler(sys.stdout)

    class JsonFormatter(logging.Formatter):
        def format(self, record: logging.LogRecord) -> str:
            return json.dumps(
                {
                    "ts": self.formatTime(record, "%Y-%m-%dT%H:%M:%S"),
                    "level": record.levelname,
                    "module": "webwatch",
                    "event": record.getMessage(),
                },
                ensure_ascii=False,
            )

    handler.setFormatter(JsonFormatter())
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)


class FetchError(RuntimeError):
    """取得失敗(接続・HTTPエラー・robots拒否)。"""


class Fetcher:
    """アクセスマナーを実装レベルで強制するHTTPクライアント。

    - robots.txt を確認し、拒否されたURLは取得しない(FetchError)
    - 同一ドメインへのアクセスは interval_seconds 以上あける
    - User-Agent を明示する
    """

    def __init__(self, user_agent: str, interval_seconds: float, timeout_seconds: float) -> None:
        self.user_agent = user_agent
        self.interval = interval_seconds
        self.timeout = timeout_seconds
        self._last_access: dict[str, float] = {}
        self._robots: dict[str, urllib.robotparser.RobotFileParser] = {}

    def fetch(self, url: str) -> str:
        domain = urlparse(url).netloc

        if not self._robots_allowed(url):
            raise FetchError(f"robots.txt により取得が許可されていません: {url}")

        # 同一ドメインへの間隔制御
        wait = self.interval - (time.monotonic() - self._last_access.get(domain, 0.0))
        if wait > 0:
            time.sleep(wait)

        try:
            resp = requests.get(
                url, headers={"User-Agent": self.user_agent}, timeout=self.timeout
            )
            resp.raise_for_status()
        except requests.Timeout as e:
            raise FetchError(f"タイムアウト({self.timeout}秒): {url}") from e
        except requests.RequestException as e:
            raise FetchError(f"取得失敗: {url} ({e})") from e
        finally:
            self._last_access[domain] = time.monotonic()

        resp.encoding = resp.apparent_encoding or resp.encoding
        logger.info(f"fetched {url} ({len(resp.text)} chars)")
        return resp.text

    def _robots_allowed(self, url: str) -> bool:
        parsed = urlparse(url)
        base = f"{parsed.scheme}://{parsed.netloc}"
        if base not in self._robots:
            rp = urllib.robotparser.RobotFileParser()
            rp.set_url(f"{base}/robots.txt")
            try:
                rp.read()
            except OSError:
                # robots.txt が取得できないサイトは保守的に「許可」扱いとしつつ記録する
                logger.warning(f"robots.txt を取得できません(許可として続行): {base}")
                rp = None  # type: ignore[assignment]
            self._robots[base] = rp
        rp = self._robots[base]
        return True if rp is None else rp.can_fetch(self.user_agent, url)
