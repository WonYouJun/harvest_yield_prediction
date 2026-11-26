import json
import sqlite3
from pathlib import Path
import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestRegressor

FEATURES = [
    "ma_3week",
    "dli_daily",
    "t_mean",
    "vpd_mean",
    "fruit_load",
    "cum_p_n_8w",
    "tr_ratio",
]
TARGET = "y_actual"


def load_source_data(db_path: str = "smartfarm_yield_swing.db", csv_path: str = "raw_yield_source.csv") -> pd.DataFrame:
    """SQLite DB 우선 로드, 없으면 원천 CSV 로드"""
    if Path(db_path).exists():
        conn = sqlite3.connect(db_path)
        try:
            df = pd.read_sql_query("SELECT * FROM weekly_yield_data", conn)
            return df
        finally:
            conn.close()

    if Path(csv_path).exists():
        return pd.read_csv(csv_path)

    raise FileNotFoundError(
        f"원천 데이터를 찾을 수 없습니다. '{db_path}' 또는 '{csv_path}' 파일을 준비해 주세요."
    )


def prepare_features(df: pd.DataFrame) -> pd.DataFrame:
    """누락된 파생 피처가 있으면 원천 컬럼으로부터 보완 계산"""
    out = df.copy().sort_values(["zone_id", "week_no"]).reset_index(drop=True)

    # 3주 이동평균 수확량
    if "ma_3week" not in out.columns and TARGET in out.columns:
        out["ma_3week"] = (
            out.groupby("zone_id")[TARGET]
            .transform(lambda s: s.shift(1).rolling(3, min_periods=1).mean())
            .fillna(0.0)
        )

    # 일일 광량(DLI) 보완 (파생광합성량/일사량 컬럼이 있을 경우)
    if "dli_daily" not in out.columns:
        for cand in ["dli", "solar_dli", "rad_mj"]:
            if cand in out.columns:
                out["dli_daily"] = out[cand]
                break

    # 8주 누적 순광합성량 및 T/R율 기본 보완값
    if "cum_p_n_8w" not in out.columns:
        base_pn = out.get("dli_daily", pd.Series(20.0, index=out.index)) * 0.65
        out["cum_p_n_8w"] = (
            base_pn.groupby(out["zone_id"]).transform(lambda s: s.rolling(8, min_periods=1).sum())
        )

    if "tr_ratio" not in out.columns:
        out["tr_ratio"] = 0.75

    return out


def train_and_export_json(
    zone_id: str = "Zone_A",
    base_week: int = 20,
    min_weeks: int = 8,
    output_json: str = "ml_pred.json",
    season_start_date: str = "2026-09-01",
) -> dict:
    raw_df = load_source_data()
    df = prepare_features(raw_df)

    zone_df = df[df["zone_id"] == zone_id].sort_values("week_no").reset_index(drop=True)
    train_df = zone_df[(zone_df["week_no"] <= base_week) & (zone_df[TARGET].notna())].copy()

    if len(train_df) < min_weeks:
        raise ValueError(
            f"학습 데이터가 부족합니다 ({len(train_df)}주). 최소 {min_weeks}주 이상 필요합니다."
        )

    X_train = train_df[FEATURES].ffill().bfill().fillna(0.0)
    y_train = train_df[TARGET].astype(float)

    model = RandomForestRegressor(
        n_estimators=150,
        max_depth=6,
        min_samples_leaf=2,
        random_state=42,
    )
    model.fit(X_train, y_train)

    # 전체 주차(또는 기준 주차 이후 롤아웃 주차) 예측 생성
    start_ts = pd.to_datetime(season_start_date)
    predictions = []

    # 기존 주차의 피처 기반 예측 + 이후 주차 순차 롤아웃
    last_row = train_df.iloc[-1].copy()
    history_y = list(train_df[TARGET].astype(float).values)

    for _, row in zone_df.iterrows():
        wk = int(row["week_no"])
        wk_date = (start_ts + pd.Timedelta(weeks=max(wk - 1, 0))).strftime("%Y-%m-%d")

        if wk <= base_week:
            x_row = pd.DataFrame([row[FEATURES]])[FEATURES].ffill().bfill().fillna(0.0)
            y_hat = float(model.predict(x_row)[0])
        else:
            steps = wk - base_week
            decay = float(np.exp(-0.3 * steps))
            ma3 = float(np.mean(history_y[-3:])) if len(history_y) >= 1 else 0.0

            next_feat = {
                "ma_3week": ma3,
                "dli_daily": float(row.get("dli_daily", last_row.get("dli_daily", 20.0))) * decay
                + float(last_row.get("dli_daily", 20.0)) * (1.0 - decay),
                "t_mean": float(row.get("t_mean", last_row.get("t_mean", 21.0))) * decay
                + float(last_row.get("t_mean", 21.0)) * (1.0 - decay),
                "vpd_mean": float(row.get("vpd_mean", last_row.get("vpd_mean", 1.0))) * decay
                + float(last_row.get("vpd_mean", 1.0)) * (1.0 - decay),
                "fruit_load": float(row.get("fruit_load", last_row.get("fruit_load", 1.0))),
                "cum_p_n_8w": float(row.get("cum_p_n_8w", last_row.get("cum_p_n_8w", 100.0))),
                "tr_ratio": float(row.get("tr_ratio", last_row.get("tr_ratio", 0.75))),
            }
            x_next = pd.DataFrame([next_feat])[FEATURES].fillna(0.0)
            y_hat = float(model.predict(x_next)[0])
            history_y.append(y_hat)

        predictions.append({
            "week": wk,
            "week_no": wk,
            "date": wk_date,
            "y_pred": round(y_hat, 3),
            "yield": round(y_hat, 3),
        })

    payload = {
        "zone": zone_id,
        "base_week": base_week,
        "model": "RandomForestRegressor (Tab6)",
        "features": FEATURES,
        "predictions": predictions,
    }

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    return payload


if __name__ == "__main__":
    result = train_and_export_json(
        zone_id="Zone_A",
        base_week=20,
        output_json="ml_pred.json",
    )
    print(f"저장 완료: ml_pred.json (총 {len(result['predictions'])}개 주차)")
