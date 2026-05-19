"""
investment-signals: 종가 기준 자동 매매 시그널 시스템
지침 v3 기반 — 매일 장 마감 후 종가로 점수 산출
"""
import json
import os
import shutil
import sys
from datetime import datetime, timezone, timedelta

import certifi
import requests

# curl_cffi는 경로에 비ASCII 문자가 있으면 SSL 인증서를 로드하지 못함 (Windows 한글 경로 등)
# yfinance import 전에 CA 번들을 ASCII 경로로 복사해서 환경변수로 지정
_ca_src = certifi.where()
if any(ord(c) > 127 for c in _ca_src):
    # TEMP 경로도 한글일 수 있으므로 항상 ASCII 경로로 고정
    _ca_dst = "C:\\Windows\\Temp\\cacert.pem"
    shutil.copy2(_ca_src, _ca_dst)
    os.environ["CURL_CA_BUNDLE"] = _ca_dst
    os.environ["SSL_CERT_FILE"] = _ca_dst

from pathlib import Path
import yfinance as yf
import pandas as pd
import pandas_ta_classic as ta

# 프로젝트 루트 경로 (main.py가 있는 폴더)
ROOT = Path(__file__).parent
TICKERS_PATH = ROOT / "tickers.json"
PORTFOLIO_PATH = ROOT / "portfolio.json"
DATA_DIR = ROOT / "data"


def load_config():
    """Google Sheets에서 관심종목과 보유종목 정보를 읽어온다.
    환경변수 GOOGLE_SERVICE_ACCOUNT_JSON과 GOOGLE_SHEET_ID가 있으면 시트를 읽고,
    없으면 로컬 fallback으로 tickers.json/portfolio.json을 읽는다 (개발용).

    Returns:
        (tickers_dict, portfolio_dict): 기존 JSON 파일 구조와 호환되는 형태
    """
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    sheet_id = os.environ.get("GOOGLE_SHEET_ID")

    if not sa_json or not sheet_id:
        # 로컬 개발 fallback: 기존 JSON 파일 읽기
        print("⚠️  GOOGLE_SERVICE_ACCOUNT_JSON 또는 GOOGLE_SHEET_ID 환경변수가 없음 → 로컬 JSON 파일 사용")
        with open(TICKERS_PATH, "r", encoding="utf-8") as f:
            tickers = json.load(f)
        with open(PORTFOLIO_PATH, "r", encoding="utf-8") as f:
            portfolio = json.load(f)
        return tickers, portfolio

    # Google Sheets 인증
    import gspread
    from google.oauth2.service_account import Credentials

    try:
        creds_dict = json.loads(sa_json)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"GOOGLE_SERVICE_ACCOUNT_JSON 파싱 실패: {e}")

    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)

    # holdings 탭 읽기
    try:
        holdings_ws = sh.worksheet("holdings")
    except Exception as e:
        raise RuntimeError(f"'holdings' 탭을 찾을 수 없음: {e}")
    holdings_rows = holdings_ws.get_all_records()

    # watchlist 탭 읽기 (관심종목 + holdings_only 통합)
    try:
        watchlist_ws = sh.worksheet("watchlist")
    except Exception as e:
        raise RuntimeError(f"'watchlist' 탭을 찾을 수 없음: {e}")
    watchlist_rows = watchlist_ws.get_all_records()

    # === holdings 데이터 변환 ===
    portfolio = {"holdings": []}
    for i, row in enumerate(holdings_rows, start=2):  # 2행부터 데이터
        # 빈 행 건너뛰기
        if not row.get("ticker") or str(row.get("ticker")).strip() == "":
            continue
        try:
            holding = {
                "ticker": str(row["ticker"]).strip(),
                "name": str(row["name"]).strip(),
                "shares": float(row["shares"]),
                "avg_price": float(row["avg_price"]),
                "currency": str(row["currency"]).strip().upper(),
            }
        except (KeyError, ValueError, TypeError) as e:
            raise RuntimeError(
                f"❌ holdings 탭 {i}행 형식 오류: {e}\n"
                f"   해당 행 데이터: {row}\n"
                f"   확인: name/ticker/shares/avg_price/currency 컬럼이 모두 있고, "
                f"shares와 avg_price는 숫자여야 합니다."
            )
        if holding["currency"] not in ("KRW", "USD"):
            raise RuntimeError(
                f"❌ holdings 탭 {i}행: currency는 'KRW' 또는 'USD'만 가능 (현재 값: '{holding['currency']}')"
            )
        portfolio["holdings"].append(holding)

    # === watchlist 데이터 변환 ===
    tickers = {"watchlist": {"major": [], "minor": []}, "holdings_only": []}
    for i, row in enumerate(watchlist_rows, start=2):
        if not row.get("ticker") or str(row.get("ticker")).strip() == "":
            continue
        try:
            item = {
                "name": str(row["name"]).strip(),
                "ticker": str(row["ticker"]).strip(),
                "market": str(row["market"]).strip().upper(),
            }
            category = str(row["category"]).strip().lower()
        except (KeyError, ValueError, TypeError) as e:
            raise RuntimeError(
                f"❌ watchlist 탭 {i}행 형식 오류: {e}\n"
                f"   해당 행 데이터: {row}"
            )

        if item["market"] not in ("KR", "US"):
            raise RuntimeError(
                f"❌ watchlist 탭 {i}행: market은 'KR' 또는 'US'만 가능 (현재 값: '{item['market']}')"
            )

        if category == "major":
            tickers["watchlist"]["major"].append(item)
        elif category == "minor":
            tickers["watchlist"]["minor"].append(item)
        elif category == "holdings_only":
            tickers["holdings_only"].append(item)
        else:
            raise RuntimeError(
                f"❌ watchlist 탭 {i}행: category는 'major', 'minor', 'holdings_only' 중 하나여야 함 "
                f"(현재 값: '{category}')"
            )

    # 최소 검증
    total_watchlist = len(tickers["watchlist"]["major"]) + len(tickers["watchlist"]["minor"]) + len(tickers["holdings_only"])
    if total_watchlist == 0:
        raise RuntimeError("❌ watchlist 탭에 종목이 하나도 없음")
    if len(portfolio["holdings"]) == 0:
        print("⚠️  holdings 탭이 비어있음 — 보유 종목 없음으로 진행")

    print(f"✅ 시트 읽기 성공: 관심종목 {total_watchlist}개 (major {len(tickers['watchlist']['major'])} / minor {len(tickers['watchlist']['minor'])} / holdings_only {len(tickers['holdings_only'])}), 보유종목 {len(portfolio['holdings'])}개")

    return tickers, portfolio


def get_all_monitored_tickers(tickers):
    """관심종목(메이저+마이너) + 보유분 자동 모니터링 종목을 하나의 리스트로 합쳐 반환.
    각 항목은 dict로 ticker, name, market, category 키를 가진다.
    중복 티커는 제거 (현대백화점·한화에어로스페이스처럼 관심+보유 양쪽인 종목)."""
    result = []
    seen = set()

    for item in tickers["watchlist"]["major"]:
        if item["ticker"] not in seen:
            result.append({**item, "category": "major"})
            seen.add(item["ticker"])

    for item in tickers["watchlist"]["minor"]:
        if item["ticker"] not in seen:
            result.append({**item, "category": "minor"})
            seen.add(item["ticker"])

    for item in tickers["holdings_only"]:
        if item["ticker"] not in seen:
            result.append({**item, "category": "holdings_only"})
            seen.add(item["ticker"])

    return result


def fetch_price_data(ticker: str, period: str = "1y") -> pd.DataFrame:
    """yfinance에서 OHLCV 데이터를 가져온다.

    Args:
        ticker: yfinance 티커 (예: "AMZN", "000660.KS")
        period: 가져올 기간 (기본 1년 — 120일 SMA 계산을 위해 최소 1년 필요)

    Returns:
        Open/High/Low/Close/Volume 컬럼을 가진 DataFrame.
        가져오기 실패 시 빈 DataFrame 반환.
    """
    try:
        df = yf.download(
            ticker,
            period=period,
            interval="1d",
            auto_adjust=True,
            progress=False,
        )
        if df.empty:
            return df
        # yfinance가 가끔 MultiIndex 컬럼으로 주는 경우가 있어 평탄화
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
        # 결측 행 제거
        df = df.dropna()
        return df
    except Exception as e:
        print(f"[ERROR] {ticker} 데이터 수집 실패: {e}")
        return pd.DataFrame()


def calculate_indicators(df: pd.DataFrame, market: str) -> dict:
    """OHLCV 데이터로부터 지침 4장의 모든 지표를 계산하여 dict로 반환.

    Args:
        df: fetch_price_data가 반환한 OHLCV DataFrame
        market: "KR" 또는 "US" (이동평균 기간 결정)

    Returns:
        다음 키를 가진 dict:
            close, prev_close       : 최근 종가, 전일 종가
            ema_short, ema_long     : 단기/중기 EMA (한국 5/20, 미국 9/21)
            sma_mid, sma_long       : 중기/장기 SMA (한국 60/120, 미국 50/200)
            ema_short_prev, ema_long_prev : 전일 EMA (크로스 판정용)
            bb_upper, bb_middle, bb_lower : 볼린저밴드 (20일, ±2σ)
            bb_lower_prev, bb_upper_prev  : 전일 BB 상/하단 (반등·꺾임 판정용)
            bb_width, bb_width_avg  : 현재 밴드폭, 최근 60일 평균 밴드폭 (스퀴즈 판정용)
            adx, plus_di, minus_di  : ADX 14일과 +DI, -DI
            adx_prev                : 전일 ADX (30→미만 전환 판정용)
            volume, vol_avg20       : 현재 거래량, 20일 평균 거래량
            vol_ratio               : 현재 거래량 / 20일 평균 비율
        데이터 부족 시 None 반환.
    """
    # 최소 데이터 길이 체크 (200일 SMA 계산을 위해 200일 이상 필요)
    if df.empty or len(df) < 200:
        return None

    # 이동평균 기간 (지침 4-1)
    if market == "KR":
        short_period, long_period = 5, 20      # EMA
        mid_period, long_sma_period = 60, 120  # SMA
    else:  # US (일본 인펙스는 미국 기준 준용이지만 현재 모니터링 제외)
        short_period, long_period = 9, 21      # EMA
        mid_period, long_sma_period = 50, 200  # SMA

    close = df["Close"]
    high = df["High"]
    low = df["Low"]
    volume = df["Volume"]

    # === 이동평균 (지침 4-1) ===
    ema_short = ta.ema(close, length=short_period)
    ema_long = ta.ema(close, length=long_period)
    sma_mid = ta.sma(close, length=mid_period)
    sma_long = ta.sma(close, length=long_sma_period)

    # === 볼린저 밴드 (지침 4-2): 20일 SMA ± 2σ ===
    bb = ta.bbands(close, length=20, std=2.0)
    # pandas-ta-classic의 bbands 컬럼: BBL_20_2.0, BBM_20_2.0, BBU_20_2.0, BBB_20_2.0, BBP_20_2.0
    bb_lower = bb.iloc[:, 0]   # 하단
    bb_middle = bb.iloc[:, 1]  # 중심
    bb_upper = bb.iloc[:, 2]   # 상단
    bb_width = bb_upper - bb_lower
    bb_width_avg = bb_width.rolling(window=60).mean()

    # === ADX (지침 4-3): 14일 ===
    adx_df = ta.adx(high, low, close, length=14)
    # pandas-ta-classic의 adx 컬럼: ADX_14, DMP_14 (+DI), DMN_14 (-DI)
    adx = adx_df.iloc[:, 0]
    plus_di = adx_df.iloc[:, 1]
    minus_di = adx_df.iloc[:, 2]

    # === 거래량 (지침 4-4): 최근 20영업일 평균 ===
    vol_avg20 = volume.rolling(window=20).mean()

    # 결과 추출 (마지막 행)
    def last(s):
        v = s.iloc[-1]
        return float(v) if pd.notna(v) else None

    def prev(s):
        v = s.iloc[-2]
        return float(v) if pd.notna(v) else None

    result = {
        "close": last(close),
        "prev_close": prev(close),
        "ema_short": last(ema_short),
        "ema_long": last(ema_long),
        "ema_short_prev": prev(ema_short),
        "ema_long_prev": prev(ema_long),
        "sma_mid": last(sma_mid),
        "sma_long": last(sma_long),
        "bb_upper": last(bb_upper),
        "bb_middle": last(bb_middle),
        "bb_lower": last(bb_lower),
        "bb_upper_prev": prev(bb_upper),
        "bb_lower_prev": prev(bb_lower),
        "bb_width": last(bb_width),
        "bb_width_avg": last(bb_width_avg),
        "adx": last(adx),
        "adx_prev": prev(adx),
        "plus_di": last(plus_di),
        "minus_di": last(minus_di),
        "volume": last(volume),
        "vol_avg20": last(vol_avg20),
    }
    if result["vol_avg20"] and result["vol_avg20"] > 0:
        result["vol_ratio"] = result["volume"] / result["vol_avg20"]
    else:
        result["vol_ratio"] = None

    return result


def score_buy(ind: dict) -> dict:
    """지침 5-1 매수 판정. 충족된 조건과 합산 점수를 반환.

    조건별 점수:
        1. 골든크로스 (단기 EMA가 장기 EMA를 상향 돌파)         : 5점
        2. 볼린저 밴드 하단 터치 후 반등                          : 3점
        3. 볼린저 밴드 확장 + 밴드워크 (상단 타고 상승)           : 3점
        4. ADX >= 30 + (+DI > -DI)                              : 3점
        5. 거래량 돌파 확인 (전일 대비 가격 상승 + 거래량 2배 이상): 5점

    조건 2와 3은 상호 배타적이므로 실질 만점은 16점.

    Returns:
        {"total": int, "details": [{"name": str, "score": int, "passed": bool, "note": str}, ...]}
    """
    details = []
    total = 0

    # === 1. 골든크로스 (5점) ===
    # 전일: 단기 EMA <= 장기 EMA, 오늘: 단기 EMA > 장기 EMA
    golden_cross = False
    if all(ind.get(k) is not None for k in ["ema_short", "ema_long", "ema_short_prev", "ema_long_prev"]):
        if ind["ema_short_prev"] <= ind["ema_long_prev"] and ind["ema_short"] > ind["ema_long"]:
            golden_cross = True
    details.append({
        "name": "골든크로스",
        "score": 5 if golden_cross else 0,
        "passed": golden_cross,
        "note": f"단기EMA {ind.get('ema_short', 0):.2f} vs 장기EMA {ind.get('ema_long', 0):.2f}",
    })
    if golden_cross:
        total += 5

    # === 2. BB 하단 터치 후 반등 (3점) ===
    # 전일 저가가 BB 하단 근처(또는 이하), 오늘 종가가 전일 종가보다 위 + BB 하단 위
    bb_lower_bounce = False
    if all(ind.get(k) is not None for k in ["close", "prev_close", "bb_lower", "bb_lower_prev"]):
        # 전일에 하단 터치(1% 이내 또는 이하) + 오늘 반등(전일보다 상승 + 하단 위)
        touched_yesterday = ind["prev_close"] <= ind["bb_lower_prev"] * 1.01
        bounced_today = ind["close"] > ind["prev_close"] and ind["close"] > ind["bb_lower"]
        if touched_yesterday and bounced_today:
            bb_lower_bounce = True
    details.append({
        "name": "BB 하단 반등",
        "score": 3 if bb_lower_bounce else 0,
        "passed": bb_lower_bounce,
        "note": f"종가 {ind.get('close', 0):.2f} vs BB하단 {ind.get('bb_lower', 0):.2f}",
    })
    if bb_lower_bounce:
        total += 3

    # === 3. BB 확장 + 밴드워크 (3점) — 조건 2와 상호 배타 ===
    # 밴드 폭이 60일 평균보다 크고, 종가가 상단 근처(95% 이상)에서 상승 중
    band_walk = False
    if not bb_lower_bounce and all(ind.get(k) is not None for k in ["close", "prev_close", "bb_upper", "bb_width", "bb_width_avg"]):
        band_expanded = ind["bb_width"] > ind["bb_width_avg"] * 1.1  # 평균보다 10% 이상 확장
        riding_upper = ind["close"] >= ind["bb_upper"] * 0.95 and ind["close"] > ind["prev_close"]
        if band_expanded and riding_upper:
            band_walk = True
    details.append({
        "name": "BB 밴드워크",
        "score": 3 if band_walk else 0,
        "passed": band_walk,
        "note": f"종가 {ind.get('close', 0):.2f} vs BB상단 {ind.get('bb_upper', 0):.2f}, 밴드폭 {ind.get('bb_width', 0):.2f}",
    })
    if band_walk:
        total += 3

    # === 4. ADX >= 30 + (+DI > -DI) (3점) ===
    adx_buy = False
    if all(ind.get(k) is not None for k in ["adx", "plus_di", "minus_di"]):
        if ind["adx"] >= 30 and ind["plus_di"] > ind["minus_di"]:
            adx_buy = True
    details.append({
        "name": "ADX>=30 상승추세",
        "score": 3 if adx_buy else 0,
        "passed": adx_buy,
        "note": f"ADX {ind.get('adx', 0):.2f}, +DI {ind.get('plus_di', 0):.2f}, -DI {ind.get('minus_di', 0):.2f}",
    })
    if adx_buy:
        total += 3

    # === 5. 거래량 돌파 확인 (5점) ===
    # 가격이 전일보다 상승 + 거래량이 20일 평균의 2배 이상
    volume_breakout = False
    if all(ind.get(k) is not None for k in ["close", "prev_close", "vol_ratio"]):
        if ind["close"] > ind["prev_close"] and ind["vol_ratio"] >= 2.0:
            volume_breakout = True
    details.append({
        "name": "거래량 돌파(2배+)",
        "score": 5 if volume_breakout else 0,
        "passed": volume_breakout,
        "note": f"거래량 {ind.get('vol_ratio', 0):.2f}배, 종가 {ind.get('close', 0):.2f} vs 전일 {ind.get('prev_close', 0):.2f}",
    })
    if volume_breakout:
        total += 5

    return {"total": total, "details": details}


def buy_verdict(total: int) -> str:
    """지침 5-1의 점수 → 판정 매핑."""
    if total >= 16:
        return "강력 매수"
    elif total >= 12:
        return "적극 매수"
    elif total >= 8:
        return "매수 검토"
    else:
        return "관망"


def score_warning(ind: dict) -> dict:
    """지침 5-2 경고 판정. 충족된 조건과 합산 점수를 반환.

    조건별 점수:
        1.  데드크로스 (단기 EMA가 장기 EMA를 하향 돌파)              : 5점
        2.  BB 상단 터치 후 꺾임                                       : 3점
        3.  볼린저 스퀴즈 (밴드 폭이 60일 평균의 50% 이하로 축소)      : 2점
        4.  ADX>=30 + (-DI > +DI) 또는 ADX 30 이상→미만 전환            : 3점
        5a. 하락 + 대량 거래 (전일 대비 하락 + 거래량 2배 이상)         : 5점
        5b. 상승 + 거래량 감소 (가격 상승 + 거래량 0.7배 이하)          : 3점
        5c. 거래량 클라이맥스 (거래량 3배 이상 폭발)                    : 4점

    Returns:
        {"total": int, "details": [...]}
    """
    details = []
    total = 0

    # === 1. 데드크로스 (5점) ===
    dead_cross = False
    if all(ind.get(k) is not None for k in ["ema_short", "ema_long", "ema_short_prev", "ema_long_prev"]):
        if ind["ema_short_prev"] >= ind["ema_long_prev"] and ind["ema_short"] < ind["ema_long"]:
            dead_cross = True
    details.append({
        "name": "데드크로스",
        "score": 5 if dead_cross else 0,
        "passed": dead_cross,
        "note": f"단기EMA {ind.get('ema_short', 0):.2f} vs 장기EMA {ind.get('ema_long', 0):.2f}",
    })
    if dead_cross:
        total += 5

    # === 2. BB 상단 터치 후 꺾임 (3점) ===
    bb_upper_reject = False
    if all(ind.get(k) is not None for k in ["close", "prev_close", "bb_upper", "bb_upper_prev"]):
        touched_yesterday = ind["prev_close"] >= ind["bb_upper_prev"] * 0.99
        rejected_today = ind["close"] < ind["prev_close"] and ind["close"] < ind["bb_upper"]
        if touched_yesterday and rejected_today:
            bb_upper_reject = True
    details.append({
        "name": "BB 상단 꺾임",
        "score": 3 if bb_upper_reject else 0,
        "passed": bb_upper_reject,
        "note": f"종가 {ind.get('close', 0):.2f} vs BB상단 {ind.get('bb_upper', 0):.2f}",
    })
    if bb_upper_reject:
        total += 3

    # === 3. 볼린저 스퀴즈 (2점) ===
    bb_squeeze = False
    if all(ind.get(k) is not None for k in ["bb_width", "bb_width_avg"]):
        if ind["bb_width_avg"] > 0 and ind["bb_width"] <= ind["bb_width_avg"] * 0.5:
            bb_squeeze = True
    details.append({
        "name": "볼린저 스퀴즈",
        "score": 2 if bb_squeeze else 0,
        "passed": bb_squeeze,
        "note": f"밴드폭 {ind.get('bb_width', 0):.2f} vs 60일평균 {ind.get('bb_width_avg', 0):.2f}",
    })
    if bb_squeeze:
        total += 2

    # === 4. ADX 하락추세 또는 30 이상→미만 전환 (3점) ===
    adx_warn = False
    reason = ""
    if all(ind.get(k) is not None for k in ["adx", "plus_di", "minus_di"]):
        if ind["adx"] >= 30 and ind["minus_di"] > ind["plus_di"]:
            adx_warn = True
            reason = "하락추세"
        elif ind.get("adx_prev") is not None and ind["adx_prev"] >= 30 and ind["adx"] < 30:
            adx_warn = True
            reason = "추세 약화 전환"
    details.append({
        "name": "ADX 경고",
        "score": 3 if adx_warn else 0,
        "passed": adx_warn,
        "note": f"ADX {ind.get('adx', 0):.2f} ({reason if adx_warn else '해당 없음'})",
    })
    if adx_warn:
        total += 3

    # === 5a. 하락 + 대량 거래 (5점) ===
    vol_5a = False
    if all(ind.get(k) is not None for k in ["close", "prev_close", "vol_ratio"]):
        if ind["close"] < ind["prev_close"] and ind["vol_ratio"] >= 2.0:
            vol_5a = True
    details.append({
        "name": "하락+대량거래",
        "score": 5 if vol_5a else 0,
        "passed": vol_5a,
        "note": f"거래량 {ind.get('vol_ratio', 0):.2f}배, 하락 폭 {(ind.get('close', 0) - ind.get('prev_close', 0)):.2f}",
    })
    if vol_5a:
        total += 5

    # === 5b. 상승 + 거래량 감소 (3점) ===
    vol_5b = False
    if all(ind.get(k) is not None for k in ["close", "prev_close", "vol_ratio"]):
        if ind["close"] > ind["prev_close"] and ind["vol_ratio"] <= 0.7:
            vol_5b = True
    details.append({
        "name": "상승+거래량감소",
        "score": 3 if vol_5b else 0,
        "passed": vol_5b,
        "note": f"거래량 {ind.get('vol_ratio', 0):.2f}배",
    })
    if vol_5b:
        total += 3

    # === 5c. 거래량 클라이맥스 (4점) ===
    vol_5c = False
    if ind.get("vol_ratio") is not None and ind["vol_ratio"] >= 3.0:
        vol_5c = True
    details.append({
        "name": "거래량 클라이맥스(3배+)",
        "score": 4 if vol_5c else 0,
        "passed": vol_5c,
        "note": f"거래량 {ind.get('vol_ratio', 0):.2f}배",
    })
    if vol_5c:
        total += 4

    return {"total": total, "details": details}


def warning_verdict(total: int) -> str:
    """지침 5-2의 점수 → 판정 매핑."""
    if total >= 12:
        return "전량 정리 검토"
    elif total >= 8:
        return "비중 축소"
    elif total >= 5:
        return "주의"
    else:
        return "정상"


def check_holdings_pnl(ticker: str, current_close: float, portfolio: dict) -> dict:
    """보유 종목의 현재 종가가 손절(-7%)/1차익절(+10%)/2차익절(+20%)에 도달했는지 확인.

    Args:
        ticker: yfinance 티커
        current_close: 오늘 종가
        portfolio: load_config가 반환한 portfolio dict

    Returns:
        보유분이 아니면 {"is_holding": False}.
        보유분이면 다음 키를 포함:
            is_holding=True, avg_price, shares, pnl_pct
            stop_loss_hit, target1_hit, target2_hit (bool)
            action: "손절 무조건 정리" / "1차 익절 50%" / "2차 익절 전량" / "보유 유지"
    """
    holding = None
    for h in portfolio["holdings"]:
        if h["ticker"] == ticker:
            holding = h
            break

    if holding is None:
        return {"is_holding": False}

    avg = holding["avg_price"]
    pnl_pct = (current_close - avg) / avg * 100

    stop_loss_hit = pnl_pct <= -7.0
    target1_hit = pnl_pct >= 10.0
    target2_hit = pnl_pct >= 20.0

    # 우선순위: 손절 > 2차 익절 > 1차 익절 > 보유
    if stop_loss_hit:
        action = "손절 무조건 정리 (-7% 도달)"
    elif target2_hit:
        action = "2차 익절 전량 (+20% 도달)"
    elif target1_hit:
        action = "1차 익절 50% (+10% 도달)"
    else:
        action = "보유 유지"

    return {
        "is_holding": True,
        "avg_price": avg,
        "shares": holding["shares"],
        "currency": holding["currency"],
        "current_close": current_close,
        "pnl_pct": pnl_pct,
        "stop_loss_hit": stop_loss_hit,
        "target1_hit": target1_hit,
        "target2_hit": target2_hit,
        "action": action,
    }


def build_signal_record(ticker_info: dict, portfolio: dict) -> dict:
    """한 종목에 대해 데이터 수집 → 지표 계산 → 점수 산출 → 보유분 PnL 체크를 모두 수행하고
    하나의 dict로 반환. 종목 하나당 한 번 호출.

    Args:
        ticker_info: {"ticker": ..., "name": ..., "market": ..., "category": ...}
        portfolio: load_config가 반환한 portfolio dict

    Returns:
        시그널 결과 dict. 실패 시 status="error".
    """
    ticker = ticker_info["ticker"]
    market = ticker_info["market"]

    # 1) 데이터 수집
    df = fetch_price_data(ticker)
    if df.empty:
        return {
            "ticker": ticker,
            "name": ticker_info["name"],
            "market": market,
            "category": ticker_info["category"],
            "status": "error",
            "error_message": "데이터 수집 실패",
        }

    # 2) 지표 계산
    ind = calculate_indicators(df, market)
    if ind is None:
        return {
            "ticker": ticker,
            "name": ticker_info["name"],
            "market": market,
            "category": ticker_info["category"],
            "status": "error",
            "error_message": "데이터 부족 (200일 미만)",
        }

    # 3) 매수 / 경고 점수
    buy = score_buy(ind)
    warn = score_warning(ind)

    # 4) 보유분이면 손절/익절 체크
    pnl = check_holdings_pnl(ticker, ind["close"], portfolio)

    # 5) 결과 통합
    last_date = df.index[-1].strftime("%Y-%m-%d")

    return {
        "ticker": ticker,
        "name": ticker_info["name"],
        "market": market,
        "category": ticker_info["category"],
        "status": "ok",
        "last_date": last_date,
        "indicators": ind,
        "buy": {
            "total": buy["total"],
            "verdict": buy_verdict(buy["total"]),
            "details": buy["details"],
        },
        "warning": {
            "total": warn["total"],
            "verdict": warning_verdict(warn["total"]),
            "details": warn["details"],
        },
        "holding": pnl,
    }


def run_all(verbose: bool = True) -> list:
    """tickers.json에 등록된 모든 종목에 대해 시그널을 산출하여 리스트로 반환."""
    tickers, portfolio = load_config()
    monitored = get_all_monitored_tickers(tickers)

    results = []
    total = len(monitored)
    for i, ticker_info in enumerate(monitored, start=1):
        if verbose:
            print(f"  [{i:2d}/{total}] {ticker_info['ticker']:15s} ({ticker_info['name']}) 처리 중...", end=" ", flush=True)
        rec = build_signal_record(ticker_info, portfolio)
        results.append(rec)
        if verbose:
            if rec["status"] == "ok":
                buy_pt = rec["buy"]["total"]
                warn_pt = rec["warning"]["total"]
                print(f"매수 {buy_pt}점 / 경고 {warn_pt}점")
            else:
                print(f"❌ {rec['error_message']}")
    return results


def save_results(results: list) -> dict:
    """run_all의 결과를 data/ 폴더에 두 개의 JSON 파일로 저장.

    1) data/latest.json : 항상 가장 최근 결과 (대시보드가 읽음)
    2) data/signals_YYYYMMDD.json : 일자별 누적 기록

    저장 형식은 메타데이터 + results 리스트.

    Returns:
        저장된 메타데이터 dict (run_at, run_date, summary).
    """
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    now_kst = datetime.now(timezone(timedelta(hours=9)))
    run_at = now_kst.isoformat()
    run_date = now_kst.strftime("%Y%m%d")

    # 요약 통계
    ok_count = sum(1 for r in results if r["status"] == "ok")
    err_count = len(results) - ok_count
    buy_signals = [r for r in results if r["status"] == "ok" and r["buy"]["total"] >= 8]
    warn_signals = [r for r in results if r["status"] == "ok" and r["warning"]["total"] >= 5]
    pnl_alerts = [
        r for r in results
        if r["status"] == "ok" and r["holding"].get("is_holding")
        and (r["holding"]["stop_loss_hit"] or r["holding"]["target1_hit"] or r["holding"]["target2_hit"])
    ]

    summary = {
        "total": len(results),
        "ok": ok_count,
        "error": err_count,
        "buy_signal_count": len(buy_signals),
        "warning_signal_count": len(warn_signals),
        "pnl_alert_count": len(pnl_alerts),
    }

    payload = {
        "run_at": run_at,
        "run_date_kst": now_kst.strftime("%Y-%m-%d"),
        "summary": summary,
        "results": results,
    }

    # 1) latest.json
    latest_path = DATA_DIR / "latest.json"
    with open(latest_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    # 2) signals_YYYYMMDD.json
    daily_path = DATA_DIR / f"signals_{run_date}.json"
    with open(daily_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print(f"\n💾 저장 완료:")
    print(f"   {latest_path}")
    print(f"   {daily_path}")

    return summary


# 디스코드 임베드 색깔 (10진수 정수, Discord 표준)
COLOR_STOP_LOSS  = 0xE74C3C  # 빨강 — 손절 도달
COLOR_TARGET2    = 0x2ECC71  # 초록 — 2차 익절
COLOR_TARGET1    = 0xF1C40F  # 황색 — 1차 익절
COLOR_BUY_STRONG = 0x3498DB  # 파랑 — 적극/강력 매수
COLOR_BUY_WATCH  = 0x5DADE2  # 하늘색 — 매수 검토
COLOR_WARN_HIGH  = 0xE67E22  # 주황 — 비중 축소/정리
COLOR_WARN_LOW   = 0xF7DC6F  # 노랑 — 주의


def _build_embeds(results: list) -> list:
    """results 리스트에서 알림이 필요한 종목만 골라 디스코드 임베드 리스트로 변환.

    우선순위: 손절 > 2차익절 > 1차익절 > 매수12+ > 매수8+ > 경고8+ > 경고5+
    """
    embeds = []

    # 1) 보유분 손절·익절 알림 (가장 위에)
    for r in results:
        if r["status"] != "ok" or not r["holding"].get("is_holding"):
            continue
        h = r["holding"]
        if h["stop_loss_hit"]:
            embeds.append({
                "title": f"🚨 손절 도달 — {r['name']} ({r['ticker']})",
                "description": f"**즉시 정리 검토**\n수익률: **{h['pnl_pct']:+.2f}%**",
                "color": COLOR_STOP_LOSS,
                "fields": [
                    {"name": "현재 종가", "value": f"{h['current_close']:,.2f} {h['currency']}", "inline": True},
                    {"name": "평균 단가", "value": f"{h['avg_price']:,.2f} {h['currency']}", "inline": True},
                    {"name": "보유 수량", "value": f"{h['shares']}주", "inline": True},
                ],
            })
        elif h["target2_hit"]:
            embeds.append({
                "title": f"🟢 2차 익절 도달 (+20%) — {r['name']} ({r['ticker']})",
                "description": f"**전량 익절 검토**\n수익률: **{h['pnl_pct']:+.2f}%**",
                "color": COLOR_TARGET2,
                "fields": [
                    {"name": "현재 종가", "value": f"{h['current_close']:,.2f} {h['currency']}", "inline": True},
                    {"name": "평균 단가", "value": f"{h['avg_price']:,.2f} {h['currency']}", "inline": True},
                    {"name": "보유 수량", "value": f"{h['shares']}주", "inline": True},
                ],
            })
        elif h["target1_hit"]:
            embeds.append({
                "title": f"🟡 1차 익절 도달 (+10%) — {r['name']} ({r['ticker']})",
                "description": f"**50% 익절 검토**\n수익률: **{h['pnl_pct']:+.2f}%**",
                "color": COLOR_TARGET1,
                "fields": [
                    {"name": "현재 종가", "value": f"{h['current_close']:,.2f} {h['currency']}", "inline": True},
                    {"name": "평균 단가", "value": f"{h['avg_price']:,.2f} {h['currency']}", "inline": True},
                    {"name": "보유 수량", "value": f"{h['shares']}주", "inline": True},
                ],
            })

    # 2) 매수 시그널 알림 (12점 이상 우선)
    buy_strong = [r for r in results if r["status"] == "ok" and r["buy"]["total"] >= 12]
    buy_watch = [r for r in results if r["status"] == "ok" and 8 <= r["buy"]["total"] < 12]

    for r in buy_strong + buy_watch:
        buy_pt = r["buy"]["total"]
        verdict = r["buy"]["verdict"]
        ind = r["indicators"]
        passed_items = [d for d in r["buy"]["details"] if d["passed"]]
        reasons = "\n".join([f"• {d['name']} ({d['score']}점)" for d in passed_items])
        # 손절/익절 계산
        close = ind["close"]
        target1 = close * 1.10
        target2 = close * 1.20
        stop = close * 0.93
        embeds.append({
            "title": f"💎 매수 시그널 — {r['name']} ({r['ticker']})",
            "description": (
                f"**{verdict}** ({buy_pt}점) · {r['category']} · {r['market']}\n\n"
                f"**충족 조건:**\n{reasons}\n\n"
                f"**진입가**: {close:,.2f}\n"
                f"**1차 목표 (+10%)**: {target1:,.2f}\n"
                f"**2차 목표 (+20%)**: {target2:,.2f}\n"
                f"**손절 라인 (-7%)**: {stop:,.2f}"
            ),
            "color": COLOR_BUY_STRONG if buy_pt >= 12 else COLOR_BUY_WATCH,
        })

    # 3) 경고 시그널 알림 (8점 이상 우선)
    warn_high = [r for r in results if r["status"] == "ok" and r["warning"]["total"] >= 8]
    warn_low = [r for r in results if r["status"] == "ok" and 5 <= r["warning"]["total"] < 8]

    for r in warn_high + warn_low:
        warn_pt = r["warning"]["total"]
        verdict = r["warning"]["verdict"]
        passed_items = [d for d in r["warning"]["details"] if d["passed"]]
        reasons = "\n".join([f"• {d['name']} ({d['score']}점)" for d in passed_items])
        embeds.append({
            "title": f"⚠️ 경고 시그널 — {r['name']} ({r['ticker']})",
            "description": (
                f"**{verdict}** ({warn_pt}점) · {r['category']} · {r['market']}\n\n"
                f"**충족 조건:**\n{reasons}"
            ),
            "color": COLOR_WARN_HIGH if warn_pt >= 8 else COLOR_WARN_LOW,
        })

    return embeds


def send_error_notification(error_message: str) -> None:
    """치명적 에러(시트 형식 오류 등) 발생 시 디스코드로 알림."""
    webhook_url = os.environ.get("DISCORD_WEBHOOK_URL")
    if not webhook_url:
        print(f"[ERROR] {error_message}")
        return
    payload = {
        "embeds": [{
            "title": "🚨 시스템 오류 발생 — 시트/설정 확인 필요",
            "description": error_message[:1900],
            "color": 0xC0392B,
        }],
        "content": "**자동 실행 실패** — 본인 시트를 확인해주세요.",
    }
    try:
        requests.post(webhook_url, json=payload, timeout=10)
    except Exception:
        pass


def send_discord_notification(results: list, summary: dict) -> None:
    """알림이 필요한 종목이 있으면 Discord Webhook으로 발송.
    환경변수 DISCORD_WEBHOOK_URL이 없으면 로컬 테스트 모드로 간주하고 콘솔에만 출력.
    """
    webhook_url = os.environ.get("DISCORD_WEBHOOK_URL")

    embeds = _build_embeds(results)

    # 알림 대상이 하나도 없으면 발송 생략
    if not embeds:
        print("\n📭 알림 대상 없음 (조용한 날입니다)")
        return

    print(f"\n📨 알림 대상 {len(embeds)}건 발견")

    if not webhook_url:
        print("  (DISCORD_WEBHOOK_URL 환경변수가 없어 콘솔 출력만 진행)")
        for e in embeds:
            print(f"\n  -- {e['title']}")
            print(f"     {e.get('description', '')[:200]}")
        return

    # 디스코드 임베드는 메시지당 최대 10개 → 10개씩 나눠 발송
    now_kst = datetime.now(timezone(timedelta(hours=9))).strftime("%Y-%m-%d %H:%M KST")
    header_content = f"📊 **투자 시그널 리포트** — {now_kst}\n(총 {summary['total']}개 종목 분석, 알림 {len(embeds)}건)"

    for i in range(0, len(embeds), 10):
        batch = embeds[i : i + 10]
        payload = {"embeds": batch}
        if i == 0:
            payload["content"] = header_content
        try:
            resp = requests.post(webhook_url, json=payload, timeout=10)
            if resp.status_code in (200, 204):
                print(f"  ✅ {len(batch)}건 발송 완료 (배치 {i//10 + 1})")
            else:
                print(f"  ❌ 발송 실패: HTTP {resp.status_code} — {resp.text[:200]}")
        except Exception as e:
            print(f"  ❌ 발송 예외: {e}")


if __name__ == "__main__":
    try:
        print("=" * 60)
        print(f"투자 시그널 산출 시작: {datetime.now(timezone(timedelta(hours=9))).strftime('%Y-%m-%d %H:%M KST')}")
        print("=" * 60)

        results = run_all(verbose=True)
        summary = save_results(results)

        print("\n" + "=" * 60)
        print(f"총 {len(results)}개 종목 처리 완료")
        print("=" * 60)

        print("\n[주요 시그널 요약 — 매수 8점+ / 경고 5점+]")
        found = False
        for r in results:
            if r["status"] != "ok":
                continue
            buy_pt = r["buy"]["total"]
            warn_pt = r["warning"]["total"]
            if buy_pt >= 8 or warn_pt >= 5:
                found = True
                tag = []
                if buy_pt >= 8:
                    tag.append(f"매수 {buy_pt}점 ({r['buy']['verdict']})")
                if warn_pt >= 5:
                    tag.append(f"경고 {warn_pt}점 ({r['warning']['verdict']})")
                print(f"  • {r['ticker']:15s} ({r['name']}) — {' / '.join(tag)}")
        if not found:
            print("  (해당 종목 없음)")

        print("\n[보유분 손절/익절 도달]")
        found = False
        for r in results:
            if r["status"] != "ok" or not r["holding"].get("is_holding"):
                continue
            h = r["holding"]
            if h["stop_loss_hit"] or h["target1_hit"] or h["target2_hit"]:
                found = True
                print(f"  • {r['ticker']:15s} ({r['name']}) — {h['pnl_pct']:+.2f}% — {h['action']}")
        if not found:
            print("  (해당 종목 없음)")

        send_discord_notification(results, summary)

    except Exception as e:
        import traceback
        err_msg = f"{type(e).__name__}: {e}\n\n{traceback.format_exc()}"
        print(f"\n❌ 치명적 오류 발생:\n{err_msg}")
        send_error_notification(err_msg)
        sys.exit(1)
