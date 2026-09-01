# -*- coding: utf-8 -*-
"""금리 대시보드 데이터 갱신 파이프라인.

일별 시리즈(국고채 3/10년·CD 91일·원/달러·미 국채 2/10년·양국 기준금리 스텝)를
웹에서 수집해 data_master.json 을 갱신하고 rates_data.js 를 생성한다.
GitHub Actions(매일)와 로컬(py -3) 공용. 표준 라이브러리만 사용(--caprate 시 openpyxl).

사용:
  py -3 fetch_rates.py                          # 일별 수집 + rates_data.js 재생성
  py -3 fetch_rates.py --caprate <오피스마켓DB.xlsx>  # + Cap. Rate 탭 연동(분기)
  py -3 fetch_rates.py --no-fetch               # 수집 없이 rates_data.js 만 재생성

옵션: --master/--out 경로, --window 일별 백필 시작일(기본 2026-01-02)
"""
import argparse
import json
import os
import re
import subprocess
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

KST = timezone(timedelta(hours=9))
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

DAILY_KEYS = ["base", "us", "kt3", "kt10", "cd", "fx", "ust2", "ust10",
              "krcpi", "uscpi", "pce", "unemp"]  # 뒤 4개는 월별 지표 스텝필


def http_get(url: str, timeout: int = 40, tries: int = 3, ua: str = None) -> bytes:
    """curl 우선(FRED가 urllib TLS를 차단), 실패 시 urllib 폴백.

    ua=None 이면 curl 기본 UA 사용 — FRED는 curl+브라우저UA 조합을 차단하므로
    FRED 계열은 반드시 ua=None 으로 호출한다. 네이버는 브라우저 UA 필요.
    """
    last_err = None
    for i in range(tries):
        try:
            cmd = ["curl", "-sS", "--max-time", str(timeout)]
            if ua:
                cmd += ["-A", ua]
            r = subprocess.run(cmd + [url], capture_output=True, timeout=timeout + 10)
            if r.returncode == 0 and r.stdout:
                return r.stdout
            last_err = RuntimeError(f"curl rc={r.returncode} {r.stderr[:120]!r}")
        except Exception as e:
            last_err = e
        try:
            req = urllib.request.Request(url, headers={"User-Agent": ua} if ua else {})
            with urllib.request.urlopen(req, timeout=timeout) as r2:
                return r2.read()
        except Exception as e:
            last_err = e
            time.sleep(1.5 * (i + 1))
    raise last_err


# ── 소스별 수집기: {iso날짜: 값} 반환 ─────────────────────────────


def fetch_fred(series_id: str, start: str) -> dict:
    url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}&cosd={start}"
    out = {}
    for line in http_get(url, ua=None).decode("utf-8", "ignore").splitlines()[1:]:
        parts = line.strip().split(",")
        if len(parts) == 2 and parts[1] not in (".", ""):
            try:
                out[parts[0]] = float(parts[1])
            except ValueError:
                pass
    return out


def fetch_naver_json(category: str, code: str, start: str) -> dict:
    """m.stock.naver.com front-api (bond/exchange). closePrice=수익률 or 매매기준율."""
    out = {}
    for page in range(1, 30):
        url = (f"https://m.stock.naver.com/front-api/marketIndex/prices"
               f"?category={category}&reutersCode={urllib.parse.quote(code)}"
               f"&page={page}&pageSize=20")
        data = json.loads(http_get(url, ua=UA["User-Agent"]).decode("utf-8", "ignore"))
        rows = data.get("result") or []
        if isinstance(rows, dict):
            rows = rows.get("priceInfos") or rows.get("prices") or []
        if not rows:
            break
        stop = False
        for it in rows:
            if not isinstance(it, dict):
                continue
            d = str(it.get("localTradedAt", ""))[:10]
            v = str(it.get("closePrice", "")).replace(",", "")
            if not d or not v:
                continue
            out[d] = float(v)
            if d < start:
                stop = True
        if stop:
            break
        time.sleep(0.25)
    return {d: v for d, v in out.items() if d >= start}


def fetch_naver_daily_quote(code: str, start: str, kind: str = "interest") -> dict:
    """finance.naver.com 일별시세 HTML (EUC-KR). kind: interest|exchange"""
    page_url = ("https://finance.naver.com/marketindex/%sDailyQuote.naver"
                "?marketindexCd=%s&page=%d")
    out = {}
    for page in range(1, 60):
        html = http_get(page_url % (kind, code, page), ua=UA["User-Agent"]).decode("euc-kr", "ignore")
        rows = re.findall(
            r'<td class="date">\s*([\d.]+)\s*</td>\s*<td class="num">\s*([\d,.]+)\s*</td>', html)
        if not rows:
            break
        stop = False
        for d, v in rows:
            iso = d.strip(".").replace(".", "-")
            out[iso] = float(v.replace(",", ""))
            if iso < start:
                stop = True
        if stop:
            break
        time.sleep(0.25)
    return {d: v for d, v in out.items() if d >= start}


# ── 병합·파생 ──────────────────────────────────────────────


def step_value(changes, iso_date, default):
    """RATE_CHANGES 테이블에서 해당 일자의 기준금리 스텝 값."""
    v = default
    for d, r in changes:
        if d <= iso_date:
            v = r
        else:
            break
    return v


def month_label(iso: str) -> str:          # 2026-07 → '26.7
    y, m = iso[:7].split("-")
    return f"'{y[2:]}.{int(m)}"


def day_label(iso: str) -> str:            # 2026-07-14 → '26.7.14
    y, m, d = iso.split("-")
    return f"'{y[2:]}.{int(m)}.{int(d)}"


def trend_month_map(trend):
    return {row["m"]: row for row in trend}


def build_daily(master: dict, series: dict, start: str) -> list:
    """series: {key: {iso: val}} → 날짜 정렬·forward-fill 된 DAILY 배열."""
    dates = sorted({d for s in series.values() for d in s})
    dates = [d for d in dates if d >= start]
    rc = master.get("RATE_CHANGES", {})
    trend = master["TREND_ALL"]
    tmap = trend_month_map(trend)
    kr_default = trend[-1]["base"]
    us_default = trend[-1]["us"]

    rows, last = [], {}
    for d in dates:
        row = {"d": d, "m": day_label(d)}
        for k in ("kt3", "kt10", "cd", "fx", "ust2", "ust10"):
            v = series.get(k, {}).get(d)
            if v is None:
                v = last.get(k)
            if v is not None:
                row[k] = round(v, 3 if k != "fx" else 1)
                last[k] = row[k]
        row["base"] = step_value(rc.get("kr", []), d, kr_default)
        row["us"] = step_value(rc.get("us", []), d, us_default)
        mrow = tmap.get(month_label(d))
        if mrow:  # 월별 지표는 해당 월 값으로 스텝필(발표 최신치 유지 개념)
            for k in ("krcpi", "uscpi", "pce", "unemp"):
                if mrow.get(k) is not None:
                    row[k] = mrow[k]
                    last["_m" + k] = mrow[k]
        else:
            for k in ("krcpi", "uscpi", "pce", "unemp"):
                if "_m" + k in last:
                    row[k] = last["_m" + k]
        rows.append(row)
    return rows


def latest_two(daily: list, key: str):
    vals = [r[key] for r in daily if r.get(key) is not None]
    if not vals:
        return None, None
    return vals[-1], (vals[-2] if len(vals) > 1 else vals[-1])


MONTHLY_KEYS = ("kt3", "kt10", "cd", "ust2", "ust10", "fx")


def month_avgs(daily: list) -> dict:
    """DAILY → {월라벨: {key: 그 달 일별 평균}}.

    롤링 윈도(380일)는 달 중간에서 시작할 수 있어 **첫 달은 부분월**이므로 제외한다
    (예: DAILY가 7/7부터면 그 달 평균에 7/1~7/4가 빠져 실제와 어긋난다).
    """
    groups, order = {}, []
    for r in daily:
        ml = month_label(r["d"])
        if ml not in groups:
            groups[ml] = {}
            order.append(ml)
        for k in MONTHLY_KEYS:
            v = r.get(k)
            if v is not None:
                groups[ml].setdefault(k, []).append(v)
    if order:
        groups.pop(order[0], None)
    return {m: {k: sum(vs) / len(vs) for k, vs in d.items()} for m, d in groups.items()}


def update_current_and_trend(master: dict):
    """DAILY 최신치로 CURRENT 카드(국채·CD·환율)와 TREND_ALL 당월 행을 갱신."""
    daily = master["DAILY"]
    if not daily:
        return
    cur = master["CURRENT"]
    for card, key, dec in (("kt3", "kt3", 2), ("kt10", "kt10", 2), ("ust2", "ust2", 2),
                            ("ust10", "ust10", 2), ("cd", "cd", 2), ("fx", "fx", 0)):
        v, prev = latest_two(daily, key)
        if v is None:
            continue
        cur[card]["value"] = round(v, dec)
        cur[card]["chg"] = round(v - prev, dec)

    last_iso = daily[-1]["d"]
    mlabel = month_label(last_iso)
    trend = master["TREND_ALL"]
    avgs = month_avgs(daily)
    if trend[-1]["m"] != mlabel:            # 월 롤오버: 직전 행 복사 후 추가
        new = dict(trend[-1])
        new["m"] = mlabel
        trend.append(new)

    # 월별 행 = 그 달 일별 **평균**(완결 월은 마감 확정, 당월은 진행분 평균).
    # 예전에는 당월 행에 '최신 일별 스팟값'을 넣고 월이 바뀌면 그대로 굳어,
    # 지나간 달이 그 달 평균이 아닌 말일 근처 값으로 남는 문제가 있었다
    # (2026-07 실측: '26.5 kt3가 3.53으로 굳어 실제 월평균 3.68과 0.15%p 어긋남).
    # DAILY가 덮는 달은 매 실행마다 평균으로 다시 확정되므로 자동 교정된다.
    for row in trend:
        a = avgs.get(row["m"])
        if not a:
            continue
        for k in ("kt3", "kt10", "cd", "ust2", "ust10"):
            if k in a:
                row[k] = round(a[k], 2)
        if "fx" in a:
            row["fx"] = round(a["fx"])


def refresh_caprate(master: dict, xlsx_path: str):
    """오피스마켓DB 'Cap. Rate' 탭(연도/분기/서울전체.../국고채3년) → CAP_RATE 교체."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["Cap. Rate"]
    out, year = [], None
    for r in ws.iter_rows(min_row=2, max_col=10, values_only=True):
        if r[0] is not None and isinstance(r[0], (int, float)):
            year = int(r[0])
        q, cap, kt3 = r[1], r[2], r[8]
        if year is None or year < 2016 or q is None:
            continue
        if not isinstance(cap, (int, float)) or not isinstance(kt3, (int, float)):
            continue
        cap, kt3 = round(cap * 100, 2), round(kt3 * 100, 2)
        out.append({"m": f"'{str(year)[2:]}Q{int(q)}", "cap": cap, "kt3": kt3,
                     "spread": round(cap - kt3, 2)})
    wb.close()
    if not out:
        print("[경고] Cap. Rate 탭에서 유효 행을 찾지 못해 기존 값 유지")
        return
    old = {r["m"]: r for r in master["CAP_RATE"]}
    for r in out:
        o = old.get(r["m"])
        if o and (o["cap"] != r["cap"] or o["kt3"] != r["kt3"]):
            print(f"  변경 {r['m']}: cap {o['cap']}→{r['cap']}, kt3 {o['kt3']}→{r['kt3']}")
    added = [r["m"] for r in out if r["m"] not in old]
    if added:
        print("  신규 분기:", ", ".join(added))
    master["CAP_RATE"] = out
    print(f"CAP_RATE: {len(out)}개 분기 ({out[0]['m']}~{out[-1]['m']})")


def write_outputs(master: dict, master_path: str, out_js: str):
    with open(master_path, "w", encoding="utf-8") as f:
        json.dump(master, f, ensure_ascii=False, indent=1)
    payload = json.dumps(master, ensure_ascii=False, separators=(",", ":"))
    with open(out_js, "w", encoding="utf-8", newline="\n") as f:
        f.write("window.__RATES__ = " + payload + ";\n")
    print("saved:", master_path, "/", out_js)


def main():
    ap = argparse.ArgumentParser()
    here = os.path.dirname(os.path.abspath(__file__))
    ap.add_argument("--master", default=os.path.join(here, "data_master.json"))
    ap.add_argument("--out", default=os.path.join(here, "rates_data.js"))
    # 기본 백필 구간: 롤링 380일 (1년 일별 뷰 + 여유)
    ap.add_argument("--window",
                    default=(datetime.now(KST) - timedelta(days=380)).strftime("%Y-%m-%d"))
    ap.add_argument("--caprate", default=None, help="오피스마켓DB xlsx 경로")
    ap.add_argument("--no-fetch", action="store_true")
    a = ap.parse_args()

    master = json.load(open(a.master, encoding="utf-8"))

    if not a.no_fetch:
        series, fails = {}, []
        jobs = [
            ("ust2", lambda: fetch_fred("DGS2", a.window)),
            ("ust10", lambda: fetch_fred("DGS10", a.window)),
            ("kt3", lambda: fetch_naver_json("bond", "KR3YT=RR", a.window)),
            ("kt10", lambda: fetch_naver_json("bond", "KR10YT=RR", a.window)),
            ("fx", lambda: fetch_naver_json("exchange", "FX_USDKRW", a.window)),
            ("cd", lambda: fetch_naver_daily_quote("IRR_CD91", a.window, "interest")),
        ]
        for key, fn in jobs:
            try:
                series[key] = fn()
                print(f"{key}: {len(series[key])}일 "
                      f"(최신 {max(series[key]) if series[key] else '-'} "
                      f"= {series[key].get(max(series[key])) if series[key] else '-'})")
            except Exception as e:  # 소스 하나 실패해도 나머지는 진행
                fails.append(key)
                print(f"[경고] {key} 수집 실패: {e}")
        ok = {k: v for k, v in series.items() if v}
        if not ok:
            print("[오류] 모든 일별 소스 수집 실패 — 기존 데이터 유지, 종료 1")
            sys.exit(1)
        # 기존 DAILY 를 시드로 사용해 실패 시리즈의 과거값 보존
        seed = {}
        for r in master.get("DAILY", []):
            for k in ("kt3", "kt10", "cd", "fx", "ust2", "ust10"):
                if r.get(k) is not None:
                    seed.setdefault(k, {})[r["d"]] = r[k]
        for k, s in seed.items():
            merged = dict(s)
            merged.update(ok.get(k, {}))
            ok[k] = merged
        master["DAILY"] = build_daily(master, ok, a.window)
        update_current_and_trend(master)
        master["dailyUpdatedAt"] = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
        if fails:
            print("[경고] 실패 시리즈(기존값 유지):", ", ".join(fails))

    if a.caprate:
        refresh_caprate(master, a.caprate)

    master["updatedAt"] = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    write_outputs(master, a.master, a.out)
    d = master["DAILY"]
    if d:
        print(f"DAILY: {len(d)}일 ({d[0]['d']}~{d[-1]['d']})")
        print("최신:", {k: d[-1].get(k) for k in ('kt3', 'kt10', 'cd', 'fx', 'ust2', 'ust10')})


if __name__ == "__main__":
    main()
