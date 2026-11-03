#!/usr/bin/env python3
"""result.xlsx 를 읽어 웹용 data.json 으로 변환합니다."""
import argparse
import json
import os
import sys
import tempfile
from datetime import datetime, timedelta, timezone

import pandas as pd

KST = timezone(timedelta(hours=9))


def parse_args():
    p = argparse.ArgumentParser(description="Excel 결과를 data.json 으로 변환")
    p.add_argument("--input", default="result.xlsx", help="입력 엑셀 파일 (기본: result.xlsx)")
    p.add_argument("--output", default="data.json", help="출력 JSON 파일 (기본: data.json)")
    p.add_argument("--sheet", default="0", help="시트 이름 또는 번호 (기본: 0 = 첫 번째 시트)")
    return p.parse_args()


def read_excel(path, sheet):
    if not os.path.isfile(path):
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {path}")
    sheet_name = int(sheet) if str(sheet).isdigit() else sheet
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    except Exception as e:  # 손상된 파일, 잘못된 시트 이름 등
        raise RuntimeError(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}") from e

    df = df.dropna(how="all").dropna(axis=1, how="all")  # 완전히 빈 행/열 제거
    df.columns = [str(c).strip() for c in df.columns]
    if df.empty:
        raise ValueError("엑셀에 읽을 수 있는 데이터가 없습니다.")
    return df


def build_payload(df, source):
    # to_json 은 NaN→null, numpy 숫자, 날짜를 모두 안전하게 처리합니다.
    split = json.loads(df.to_json(orient="split", force_ascii=False, date_format="iso"))
    return {
        "updated_at": datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST"),
        "source": os.path.basename(source),
        "row_count": len(split["data"]),
        "columns": split["columns"],
        "rows": split["data"],
    }


def write_json(payload, path):
    """임시 파일에 먼저 쓴 뒤 교체하여, 실패해도 기존 data.json 이 깨지지 않게 합니다."""
    directory = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(suffix=".json", dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
        os.replace(tmp, path)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def main():
    args = parse_args()
    try:
        df = read_excel(args.input, args.sheet)
        payload = build_payload(df, args.input)
        write_json(payload, args.output)
    except FileNotFoundError as e:
        print(f"[오류] {e}", file=sys.stderr)
        return 1
    except (RuntimeError, ValueError) as e:
        print(f"[오류] {e}", file=sys.stderr)
        return 2
    except OSError as e:
        print(f"[오류] 파일을 저장하지 못했습니다: {e}", file=sys.stderr)
        return 3
    print(f"[완료] {payload['row_count']:,}행 → {args.output} ({payload['updated_at']})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
