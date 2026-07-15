# -*- coding: utf-8 -*-
"""从新浪源(akshare)抓取全股票池后复权日线 OHLCV + 中证1000指数, 用于重建真实 V46。
校验前提: 新浪后复权与用户 parquet 仅为常量缩放差, MACD 金叉时序不变。
"""
import pandas as pd, numpy as np, akshare as ak, time, os, warnings, json
from concurrent.futures import ThreadPoolExecutor, as_completed
warnings.filterwarnings("ignore")

DATA = "/root/.codebuddy/artifact/84d009eb-3049-4b97-afde-74135ba25f00/strategy/data"
OUT  = "/workspace/data"
os.makedirs(OUT, exist_ok=True)

# 股票池代码
pool = pd.read_parquet(f"{DATA}/股票池开盘价_后复权_v3.parquet")
codes = [str(c) for c in pool.columns]

def to_sina(code):
    num, ex = code.split('.')
    return ("sh" if ex == "SH" else "sz") + num

def fetch_one(code):
    for attempt in range(4):
        try:
            h = ak.stock_zh_a_daily(symbol=to_sina(code), adjust="hfq")
            if h is None or len(h) == 0:
                return code, None
            h = h.copy()
            h['date'] = pd.to_datetime(h['date'])
            h = h.set_index('date').sort_index()
            h = h[['open','high','low','close','volume']].astype(float)
            return code, h
        except Exception as e:
            if attempt < 3:
                time.sleep(1.5 * (attempt + 1))
            else:
                return code, f"ERR:{repr(e)[:60]}"

print(f"开始抓取 {len(codes)} 只股票 (新浪源, 多线程)...", flush=True)
t0 = time.time()
results = {}
errors = []
with ThreadPoolExecutor(max_workers=12) as ex:
    futs = {ex.submit(fetch_one, c): c for c in codes}
    done = 0
    for f in as_completed(futs):
        c, r = f.result()
        done += 1
        if isinstance(r, pd.DataFrame):
            results[c] = r
        else:
            errors.append((c, r))
        if done % 200 == 0:
            print(f"  进度 {done}/{len(codes)} 成功{len(results)} 失败{len(errors)} 用时{time.time()-t0:.0f}s", flush=True)

print(f"抓取完成: 成功 {len(results)} 失败 {len(errors)} 用时 {time.time()-t0:.0f}s", flush=True)
if errors:
    print("失败样例:", errors[:10], flush=True)

# 组装宽表 (index=日期, columns=code)
fields = ['open','high','low','close','volume']
panels = {f: {} for f in fields}
all_dates = None
for c, dfc in results.items():
    if all_dates is None:
        all_dates = dfc.index
    else:
        all_dates = all_dates.union(dfc.index)
for f in fields:
    panels[f] = pd.DataFrame(index=all_dates.sort_values())
for c, dfc in results.items():
    for f in fields:
        panels[f][c] = dfc[f]
for f in fields:
    panels[f] = panels[f].sort_index()
    panels[f].to_parquet(f"{OUT}/stock_{f}_hfq.parquet")
    print(f"  保存 stock_{f}_hfq.parquet: {panels[f].shape}", flush=True)

# 中证1000指数
print("抓取中证1000指数(sh000852)...", flush=True)
idx = ak.stock_zh_index_daily(symbol="sh000852")
idx['date'] = pd.to_datetime(idx['date'])
idx = idx.set_index('date').sort_index()[['open','high','low','close','volume']].astype(float)
idx.to_parquet(f"{OUT}/zz1000_hfq.parquet")
print(f"  保存 zz1000_hfq.parquet: {idx.shape} 范围 {idx.index.min().date()}->{idx.index.max().date()}", flush=True)

# 校验: 新浪open 与 parquet open 的常量缩放一致性
par_open = pool
common = [c for c in codes if c in results]
chk = []
for c in common[:30]:
    ak_o = panels['open'][c].dropna()
    par_o = par_open[c].dropna()
    ak_o.index = pd.to_datetime(ak_o.index); par_o.index = pd.to_datetime(par_o.index)
    m = pd.concat([ak_o.rename('a'), par_o.rename('p')], axis=1).dropna()
    if len(m) > 50:
        ratio = (m['a']/m['p'])
        chk.append(ratio.std()/ratio.mean())
chk = np.array(chk)
print(f"校验 前30只 缩放因子变异系数 均值={chk.mean():.4f} 最大={chk.max():.4f} (越小越接近常量缩放)", flush=True)
print("ALL DONE", flush=True)
