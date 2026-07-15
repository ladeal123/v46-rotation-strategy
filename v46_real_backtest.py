# -*- coding: utf-8 -*-
"""
V46 + 行业共振 回测引擎 (真实口径: 收盘价信号 -> T+1开盘成交, 中证1000=真实指数)
===========================================================================
数据来源: 新浪(akshare)后复权日线 OHLCV, 全股票池 2256只 + 中证1000指数。
校验: 新浪后复权与用户 parquet 开盘价为常量缩放差(变异系数~0.006), MACD金叉时序不变。
对比: 与开盘价代理版(v46_resonance_backtest.py)同口径 A/B, 验证"收盘->开盘跳空"是否为V46 alpha。

修复: 卖出条件链式 reason is None (杜绝现金重复入账bug)。
"""
import math, json, time
from collections import defaultdict
import pandas as pd

# ============ 配置 ============
FAST, SLOW, SIGNAL = 10, 20, 9
START_DATE = '2020-06-01'
END_DATE   = '2026-05-14'
INIT_CASH  = 10_000_000
POS_SIZE   = 0.02
MIN_HOLD   = 3
STOP_LOSS  = -0.15
PROFIT_THRESHOLD = 0.50
COMMISSION_RATE  = 0.0001
STAMP_TAX_RATE   = 0.0005
PENALTY = {0: 1.0, 1: 0.5, 2: 0.5, 3: 0.5}

SRC = "/workspace/data"
DATA_DIR = '/root/.codebuddy/artifact/84d009eb-3049-4b97-afde-74135ba25f00/strategy/data'

def get_board(code):
    if code.startswith('30') or code.startswith('688'): return 'chinext'
    if code.startswith('8'): return 'bse'
    return 'main'

def get_limit_price(prev_close, board):
    pct = 0.10 if board == 'main' else (0.20 if board == 'chinext' else 0.30)
    return round(prev_close * (1 + pct), 2)

print("加载数据(真实口径: 收盘价信号/开盘价成交/真实中证1000)...", flush=True)
close_p = pd.read_parquet(f'{SRC}/stock_close_hfq.parquet')
open_p  = pd.read_parquet(f'{SRC}/stock_open_hfq.parquet')
zz1000  = pd.read_parquet(f'{SRC}/zz1000_hfq.parquet')

# 行业映射
import openpyxl
code2ind = {}
wb = openpyxl.load_workbook(f'{DATA_DIR}/股票池.xlsx', read_only=True, data_only=True)
ws = wb['股票池']
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] and row[2]:
        code2ind[str(row[0]).strip()] = str(row[2]).strip()
wb.close()

# 构造每只股票的 收盘/开盘 字典 (按日期) —— 向量化
close_p.index = pd.to_datetime(close_p.index)
open_p.index  = pd.to_datetime(open_p.index)
cdate = close_p.index.strftime('%Y-%m-%d').values
odate = open_p.index.strftime('%Y-%m-%d').values
all_dates_set = set(cdate) & set(odate)
zz_dates = set(zz1000.index.strftime('%Y-%m-%d').values)
dates_valid = sorted(d for d in all_dates_set if START_DATE <= d <= END_DATE and d in zz_dates)
date_pos = {d: i for i, d in enumerate(dates_valid)}
dates_valid_set = set(dates_valid)
print(f"  交易日: {len(dates_valid)} | 范围 {dates_valid[0]}~{dates_valid[-1]}", flush=True)

close_data, open_data = {}, {}
for code in close_p.columns:
    cs = str(code).strip()
    cv = close_p[code].values; ov = open_p[code].values
    d = {}
    for k in range(len(cdate)):
        dd = cdate[k]
        if dd in dates_valid_set and not pd.isna(cv[k]) and not pd.isna(ov[k]) and cv[k] > 0 and ov[k] > 0:
            d[dd] = (float(cv[k]), float(ov[k]))
    if len(d) >= 120:
        close_data[cs] = {x: v[0] for x, v in d.items()}
        open_data[cs]  = {x: v[1] for x, v in d.items()}
print(f"  同时有收盘+开盘+行业: {sum(1 for c in close_data if c in code2ind)}只", flush=True)
codes = [c for c in close_data if c in code2ind]

# ============ 指标预计算 (基于收盘价) ============
def ema_s(pl, p):
    n = len(pl); a = 2.0/(p+1); out = [float('nan')]*n
    fv = next((i for i, v in enumerate(pl) if v > 0), None)
    if fv is None: return out
    out[fv] = pl[fv]
    for i in range(fv+1, n):
        out[i] = a*pl[i] + (1-a)*out[i-1] if pl[i] > 0 else out[i-1]
    return out

print("预计算个股指标(收盘价) + 行业动量...", flush=True)
IND = {}; gc_flag = {}; dif_pct_mat = {}
for code in codes:
    d_list = sorted(close_data[code].keys())
    cl = [close_data[code][d] for d in d_list]   # 收盘价序列(信号)
    ef = ema_s(cl, FAST); es = ema_s(cl, SLOW)
    dif = [ef[i]-es[i] if not(math.isnan(ef[i]) or math.isnan(es[i])) else float('nan') for i in range(len(cl))]
    dea = ema_s(dif, SIGNAL)
    bar = [(dif[i]-dea[i])*2.0 if not(math.isnan(dif[i]) or math.isnan(dea[i])) else float('nan') for i in range(len(cl))]
    ma5 = ema_s(cl, 5); ma144 = ema_s(cl, 144)
    dif_pct = [dif[i]/cl[i]*100 if (not math.isnan(dif[i])) and cl[i] > 0 else 0.0 for i in range(len(cl))]
    gcf = [False]*len(cl)
    for i in range(1, len(cl)):
        if not(math.isnan(dif[i-1]) or math.isnan(dea[i-1]) or math.isnan(dif[i]) or math.isnan(dea[i])):
            if dif[i-1] <= dea[i-1] and dif[i] > dea[i]:
                gcf[i] = True
    IND[code] = {'dates': d_list, 'close': cl, 'dif': dif, 'dea': dea, 'bar': bar,
                 'ma5': ma5, 'ma144': ma144, 'dif_pct': dif_pct, 'board': get_board(code),
                 'date_to_idx': {d: i for i, d in enumerate(d_list)}}
    gc_flag[code] = gcf; dif_pct_mat[code] = dif_pct

# ---- 真实中证1000 -> V4状态机 ----
print("计算真实中证1000 -> V4状态机...", flush=True)
zz_idx = zz1000.index.strftime('%Y-%m-%d').values
zz_map = {}
for k in range(len(zz_idx)):
    v = zz1000['close'].values[k]
    if not pd.isna(v): zz_map[zz_idx[k]] = float(v)
zz_close = {d: zz_map[d] for d in dates_valid if d in zz_map}
idx_vals = [zz_close.get(d, float('nan')) for d in dates_valid]
ef = ema_s(idx_vals, 10); es = ema_s(idx_vals, 20)
dif = [ef[i]-es[i] if not(math.isnan(ef[i]) or math.isnan(es[i])) else float('nan') for i in range(len(idx_vals))]
dea = ema_s(dif, 9)
bar = [(dif[i]-dea[i])*2.0 if not(math.isnan(dif[i]) or math.isnan(dea[i])) else float('nan') for i in range(len(dif))]
dif_pct = [dif[i]/idx_vals[i]*100 if (not math.isnan(dif[i])) and idx_vals[i] and idx_vals[i] > 0 else 0.0 for i in range(len(idx_vals))]
def dif_slope(dif_pct, win=2):
    out = [0.0]*len(dif_pct)
    for i in range(win, len(dif_pct)):
        y = [dif_pct[i-j] for j in range(win-1, -1, -1)]
        xm = [j-(win-1)/2 for j in range(win)]
        ym = [yi-sum(y)/win for yi in y]
        num = sum(xm[j]*ym[j] for j in range(win)); den = sum(xm[j]**2 for j in range(win))
        out[i] = num/den if den != 0 else 0
    return out
slope = dif_slope(dif_pct, 2)
def calc_pos10(dif_pct, slope, dea, bar, th_short, bottom):
    pos = 0; result = []; prev_bar = None
    for i in range(1, len(dif_pct)):
        rd = dif[i] if not math.isnan(dif[i]) else 0
        rdea = dea[i] if not math.isnan(dea[i]) else 0
        pd_ = dif[i-1] if not math.isnan(dif[i-1]) else 0
        pdea = dea[i-1] if not math.isnan(dea[i-1]) else 0
        rdp = dif_pct[i]; rsl = slope[i]
        rbar = bar[i] if (bar is not None and i < len(bar) and not math.isnan(bar[i])) else 0
        gcu = pd_ <= pdea and rd > rdea; gcd = pd_ >= pdea and rd < rdea
        if pos == 0:
            if gcu: pos = 1
            elif rdp > 0 and rsl > 0.02: pos = 1
            elif rdp < bottom and rsl > 0.03: pos = 1
            elif rdp <= th_short: pos = -1
        elif pos == 1:
            if gcd:
                pos = 0
                if rdp <= th_short: pos = -1
            elif prev_bar is not None and prev_bar < 0:
                if rdp > 0 and rsl < -0.02: pos = 0
        elif pos == -1:
            if gcu: pos = 1
            elif rdp < bottom and rsl > 0.03: pos = 1
        result.append(pos); prev_bar = rbar
    return result
v4p = calc_pos10(dif_pct, slope, dea, bar, 0.0, -3.0)
v4_state = {}
for i in range(len(dates_valid)-1):
    v4_state[dates_valid[i+1]] = 'bull' if v4p[i] == 1 else ('short' if v4p[i] == -1 else 'flat')

# ---- 行业动量 ----
print("计算行业动量矩阵...", flush=True)
industries = sorted(set(code2ind[c] for c in codes))
ind_members = defaultdict(list)
for c in codes: ind_members[code2ind[c]].append(c)
ind_avg_dif = {s: [] for s in industries}; ind_gc_frac = {s: [] for s in industries}
for di, d in enumerate(dates_valid):
    for s in industries:
        dp = []; gc = []
        for c in ind_members[s]:
            i = IND[c]['date_to_idx'].get(d)
            if i is None: continue
            dp.append(dif_pct_mat[c][i]); gc.append(1.0 if gc_flag[c][i] else 0.0)
        ind_avg_dif[s].append(sum(dp)/len(dp) if dp else 0.0)
        ind_gc_frac[s].append(sum(gc)/len(gc) if gc else 0.0)
def industry_resonance_pass(s, di, mode, param):
    if mode is None: return True
    if mode == 'difpos': return ind_avg_dif[s][di] > param
    if mode == 'topk':
        vals = [ind_avg_dif[ss][di] for ss in industries]
        vals_sorted = sorted(vals, reverse=True)
        k = max(1, int(len(vals_sorted)*param)); th = vals_sorted[k-1]
        return ind_avg_dif[s][di] >= th
    if mode == 'gcfrac': return ind_gc_frac[s][di] >= param
    if mode == 'exclude_weak': return ind_avg_dif[s][di] >= param  # param为负数, 仅剔除行业动量最差
    return True

# ============ 回测 ============
def backtest(resonance_mode=None, resonance_param=0.0, label='baseline'):
    portfolio = {}; cash = INIT_CASH; trades_log = []
    prev_state = 'flat'; peak_nav = INIT_CASH; max_dd = 0.0
    yearly = {}; sell_reason = defaultdict(int); ytb_skip = 0
    total_comm = 0.0; total_stamp = 0.0
    n = len(dates_valid); t0 = time.time()
    for date_idx in range(n):
        cur = dates_valid[date_idx]; yr = cur[:4]
        sig = dates_valid[date_idx-1] if date_idx >= 1 else None
        idx_state = v4_state.get(cur, 'flat')
        if idx_state == 'flat' and prev_state == 'short': idx_state = 'short'
        prev_state = idx_state
        max_pos = 50 if idx_state == 'bull' else (5 if idx_state == 'short' else 20)
        # 年初 NAV (按收盘价估值)
        if yr not in yearly:
            nav = cash
            for c, p in portfolio.items():
                i = IND[c]['date_to_idx'].get(cur)
                if i is not None and IND[c]['close'][i] > 0:
                    nav += IND[c]['close'][i]*p['shares']
            yearly[yr] = {'start': nav, 'end': None, 'buy': 0, 'sell': 0, 'force_sell': 0}
        # 强卖(超仓位) - 成交用开盘价
        if len(portfolio) > max_pos:
            excess = len(portfolio) - max_pos
            for code, pos in sorted(portfolio.items(), key=lambda x: x[1]['score'])[:excess]:
                j = IND[code]['date_to_idx'].get(cur)
                if j is None: continue
                o = open_data[code].get(cur, IND[code]['close'][j])
                if o <= 0: continue
                cost_tax = o*pos['shares']*(COMMISSION_RATE+STAMP_TAX_RATE)
                cash += o*pos['shares'] - cost_tax
                trades_log.append({'date': cur, 'code': code, 'dir': '卖出', 'price': o, 'shares': pos['shares'], 'pnl': (o-pos['entry_price'])*pos['shares']-cost_tax, 'reason': '超仓位', 'yr': yr})
                del portfolio[code]; yearly[yr]['force_sell'] += 1; sell_reason['超仓位'] += 1
        # 卖出(T日收盘信号 -> T+1开盘成交)
        sells = []
        for code, pos in list(portfolio.items()):
            if sig is None: continue
            i = IND[code]['date_to_idx'].get(sig)
            if i is None: continue
            dif_i = IND[code]['dif'][i]; dea_i = IND[code]['dea'][i]; ma5_i = IND[code]['ma5'][i]
            c_i = IND[code]['close'][i]   # 信号日用收盘价
            if math.isnan(dif_i) or math.isnan(dea_i) or math.isnan(c_i) or c_i <= 0: continue
            j = IND[code]['date_to_idx'].get(cur)
            if j is None: continue
            o_j = open_data[code].get(cur, c_i)  # T+1开盘成交
            if o_j <= 0 or math.isnan(o_j): o_j = c_i
            pos['hold_days'] += 1
            if c_i > pos['max_price']: pos['max_price'] = c_i
            pnl_pct = (o_j-pos['entry_price'])/pos['entry_price']
            reason = None
            if reason is None and pos['hold_days'] >= MIN_HOLD and pnl_pct <= STOP_LOSS: reason = '止损'
            if reason is None and pos['hold_days'] >= MIN_HOLD and i >= 1:
                dp = IND[code]['dif'][i-1]; ep = IND[code]['dea'][i-1]
                if not(math.isnan(dp) or math.isnan(ep)):
                    if dp > ep and dif_i <= dea_i: reason = '死叉'
            if reason is None and pnl_pct > PROFIT_THRESHOLD and not math.isnan(ma5_i) and c_i < ma5_i: reason = '动态止盈'
            if reason is not None: sells.append((code, pos, o_j, reason))
        seen = set(); sells_u = []
        for s in sells:
            if s[0] not in seen: seen.add(s[0]); sells_u.append(s)
        for code, pos, price, reason in sorted(sells_u, key=lambda x: x[1]['score']):
            cost_tax = price*pos['shares']*(COMMISSION_RATE+STAMP_TAX_RATE)
            cash += price*pos['shares'] - cost_tax
            trades_log.append({'date': cur, 'code': code, 'dir': '卖出', 'price': price, 'shares': pos['shares'], 'pnl': (price-pos['entry_price'])*pos['shares']-cost_tax, 'reason': reason, 'yr': yr})
            del portfolio[code]; yearly[yr]['sell'] += 1; sell_reason[reason] += 1
        # 买入(T日收盘信号 -> T+1开盘成交 + 涨停过滤 + 行业共振)
        if len(portfolio) < max_pos and cash > 50000 and sig is not None:
            nav = cash
            for c, p in portfolio.items():
                i = IND[c]['date_to_idx'].get(cur)
                if i is not None and IND[c]['close'][i] > 0: nav += IND[c]['close'][i]*p['shares']
            total_nav = nav
            candidates = []; sig_di = date_pos[sig]
            for code in IND:
                if code in portfolio: continue
                i = IND[code]['date_to_idx'].get(sig)
                if i is None: continue
                j = IND[code]['date_to_idx'].get(cur)
                if j is None: continue
                c_i = IND[code]['close'][i]   # 信号日收盘
                if c_i <= 0 or math.isnan(c_i): continue
                o_j = open_data[code].get(cur, c_i)  # T+1开盘
                if o_j <= 0 or math.isnan(o_j): o_j = c_i
                dif_i = IND[code]['dif'][i]; dea_i = IND[code]['dea'][i]; ma144_i = IND[code]['ma144'][i]; bar_i = IND[code]['bar'][i]
                if math.isnan(dif_i) or math.isnan(dea_i) or math.isnan(ma144_i): continue
                if bar_i < 0: continue
                if c_i < ma144_i: continue
                gc_days = -1
                for off in range(3):
                    if i-off < 0: continue
                    dn = IND[code]['dif'][i-off]; dp = IND[code]['dif'][i-off-1] if i-off-1 >= 0 else float('nan')
                    en = IND[code]['dea'][i-off]; ep = IND[code]['dea'][i-off-1] if i-off-1 >= 0 else float('nan')
                    if not(math.isnan(dn) or math.isnan(dp) or math.isnan(en) or math.isnan(ep)):
                        if dp <= ep and dn > en: gc_days = off; break
                if gc_days < 0: continue
                base_score = (dif_i/c_i*100)*0.3 + (bar_i/c_i*100)*0.7
                score = base_score * PENALTY.get(gc_days, 0.2)
                is_zt = False
                limit_p = get_limit_price(c_i, IND[code]['board'])  # 以信号日收盘算涨停
                if o_j >= limit_p: is_zt = True
                s = code2ind.get(code)
                if s is None or not industry_resonance_pass(s, sig_di, resonance_mode, resonance_param): continue
                candidates.append((score, code, o_j, is_zt))
            candidates.sort(reverse=True)
            for score, code, o, is_zt in candidates:
                if len(portfolio) >= max_pos or cash <= 50000: break
                if is_zt: ytb_skip += 1; continue
                shares = max(int(POS_SIZE*total_nav/o/100)*100, 100)
                cost = shares*o; total_cost = cost*(1+COMMISSION_RATE)
                if total_cost > cash*0.95: continue
                cash -= total_cost
                portfolio[code] = {'shares': shares, 'entry_price': o, 'hold_days': 0, 'score': score, 'max_price': o}
                trades_log.append({'date': cur, 'code': code, 'dir': '买入', 'price': o, 'shares': shares, 'pnl': 0, 'reason': f'金叉gc{gc_days}', 'yr': yr})
                yearly[yr]['buy'] += 1
        # 日终 NAV(按收盘价估值)
        nav = cash
        for c, p in portfolio.items():
            i = IND[c]['date_to_idx'].get(cur)
            if i is not None and IND[c]['close'][i] > 0: nav += IND[c]['close'][i]*p['shares']
        if nav > peak_nav: peak_nav = nav
        dd = (peak_nav-nav)/peak_nav*100
        if dd > max_dd: max_dd = dd
        yearly[yr]['end'] = nav
    final_nav = yearly[list(yearly.keys())[-1]]['end']
    total_return = (final_nav/INIT_CASH - 1)*100
    for y in yearly:
        s = yearly[y]['start']; e = yearly[y]['end']
        yearly[y]['return'] = (e/s-1)*100 if s > 0 else 0
    return {'label': label, 'final_nav': final_nav, 'total_return': total_return, 'max_dd': max_dd,
            'ratio': total_return/max_dd if max_dd > 0 else 0,
            'yearly': {y: yearly[y]['return'] for y in sorted(yearly.keys())},
            'buy': sum(yearly[y]['buy'] for y in yearly), 'sell': sum(yearly[y]['sell'] for y in yearly),
            'force_sell': sum(yearly[y]['force_sell'] for y in yearly), 'ytb_skip': ytb_skip,
            'sell_reason': dict(sell_reason), 'time': time.time()-t0}

if __name__ == '__main__':
    runs = []
    runs.append(backtest(None, 0.0, 'baseline(无行业过滤)'))
    runs.append(backtest('difpos', 0.0, '行业共振-dif>0'))
    runs.append(backtest('topk', 0.3, '行业共振-前30%行业'))
    runs.append(backtest('topk', 0.5, '行业共振-前50%行业'))
    runs.append(backtest('gcfrac', 0.10, '行业共振-金叉占比>=10%'))
    runs.append(backtest('gcfrac', 0.20, '行业共振-金叉占比>=20%'))
    runs.append(backtest('exclude_weak', -1.0, '共振-剔除最弱行业(avgDIF<-1)'))
    runs.append(backtest('exclude_weak', -2.0, '共振-剔除较弱行业(avgDIF<-2)'))

    print("\n" + "=" * 100)
    print(" V46 + 行业共振 回测对比 (真实口径: 收盘价信号+T+1开盘成交, 中证1000=真实指数, 数据=新浪后复权)")
    print("=" * 100)
    hdr = f"{'策略':<24}{'总收益%':>10}{'回撤%':>9}{'收益/回撤':>11}{'买入':>7}{'卖出':>7}{'强卖':>7}{'涨停跳过':>9}"
    print(hdr); print("-" * 100)
    for r in runs:
        print(f"{r['label']:<24}{r['total_return']:>10.2f}{r['max_dd']:>9.2f}{r['ratio']:>11.2f}{r['buy']:>7}{r['sell']:>7}{r['force_sell']:>7}{r['ytb_skip']:>9}")
    print("-" * 100)
    print("逐年收益%:")
    yrk = sorted(runs[0]['yearly'].keys())
    print(f"{'策略':<24}" + "".join(f"{y:>8}" for y in yrk))
    for r in runs:
        print(f"{r['label']:<24}" + "".join(f"{r['yearly'][y]:>8.2f}" for y in yrk))
    print("-" * 100)
    for r in runs:
        print(f"[{r['label']}] 卖出原因: {r['sell_reason']} | 耗时{r['time']:.1f}s")
    with open('/workspace/v46_real_results.json', 'w') as f:
        json.dump(runs, f, ensure_ascii=False, indent=2, default=str)
    print("已保存 /workspace/v46_real_results.json")
