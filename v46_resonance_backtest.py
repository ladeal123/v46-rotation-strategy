# -*- coding: utf-8 -*-
"""
V46 + 行业共振 回测引擎 (开盘价代理版)
=====================================
重要口径说明(必须如实告知):
- 共享盘当前缺失「收盘价长格式」与「成交量」与「中证1000指数」原始文件。
- 仅有的全周期数据为: 开盘价后复权 parquet (1543天 × 2261只, 2020-01 ~ 2026-05)
  以及 股票池.xlsx 的申万一级行业映射(2261只全覆盖)。
- 因此本引擎用【开盘价】同时计算信号与成交(T日盘后看开盘价信号 -> T+1开盘价成交),
  用【池内等权指数】代理中证1000做V4状态机regime。这与原V46(收盘价信号)口径不同,
  绝对收益不可直接与历史文档(+48%/+579%)比较; 但 baseline 与 行业共振 在同一口径下
  的 A/B 对比是有效的、真实的。

比旧代码修正: 卖出条件用链式 reason is None, 杜绝现金重复入账bug。
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

DATA_DIR = '/root/.codebuddy/artifact/84d009eb-3049-4b97-afde-74135ba25f00/strategy/data'

def get_board(code):
    if code.startswith('30') or code.startswith('688'): return 'chinext'
    if code.startswith('8'): return 'bse'
    return 'main'

def get_limit_price(prev_close, board):
    pct = 0.10 if board == 'main' else (0.20 if board == 'chinext' else 0.30)
    return round(prev_close * (1 + pct), 2)

print("加载数据...", flush=True)
# ---- 开盘价 (宽表: index=日期, columns=代码) ----
df = pd.read_parquet(f'{DATA_DIR}/股票池开盘价_后复权_v3.parquet')
all_dates = [str(d)[:10] for d in df.index]
dates_valid = sorted(d for d in all_dates if START_DATE <= d <= END_DATE)
date_pos = {d: i for i, d in enumerate(dates_valid)}
open_data = {}
for code in df.columns:
    cs = str(code).strip()
    series = df[code]
    d = {dates_valid[i]: float(series.iloc[i]) for i in range(len(dates_valid))
         if i < len(series) and pd.notna(series.iloc[i]) and float(series.iloc[i]) > 0}
    if len(d) >= 30:
        open_data[cs] = d
print(f"  开盘价: {len(open_data)}只 | {len(dates_valid)}天", flush=True)

# ---- 行业映射 (申万一级) ----
import openpyxl
code2ind = {}
wb = openpyxl.load_workbook(f'{DATA_DIR}/股票池.xlsx', read_only=True, data_only=True)
ws = wb['股票池']
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] and row[2]:
        code2ind[str(row[0]).strip()] = str(row[2]).strip()
wb.close()
# 只保留有开盘价且有行业的股票
codes = [c for c in open_data if c in code2ind]
print(f"  同时有开盘价+行业: {len(codes)}只", flush=True)

# ============ 指标预计算 ============
def ema_s(pl, p):
    n = len(pl); a = 2.0/(p+1); out = [float('nan')]*n
    fv = next((i for i, v in enumerate(pl) if v > 0), None)
    if fv is None: return out
    out[fv] = pl[fv]
    for i in range(fv+1, n):
        if pl[i] == 0:
            out[i] = out[i-1] if not math.isnan(out[i-1]) else float('nan')
        else:
            out[i] = a*pl[i] + (1-a)*out[i-1]
    return out

print("预计算个股指标 + 行业动量...", flush=True)
IND = {}
gc_flag = {}          # code -> list(bool) 金叉发生日
dif_pct_mat = {}      # code -> list(float) DIF%
for code in codes:
    d_list = sorted(open_data[code].keys())
    pl = [open_data[code][d] for d in d_list]
    ef = ema_s(pl, FAST); es = ema_s(pl, SLOW)
    dif = [ef[i]-es[i] if not(math.isnan(ef[i]) or math.isnan(es[i])) else float('nan') for i in range(len(pl))]
    dea = ema_s(dif, SIGNAL)
    bar = [(dif[i]-dea[i])*2.0 if not(math.isnan(dif[i]) or math.isnan(dea[i])) else float('nan') for i in range(len(pl))]
    ma5  = ema_s(pl, 5); ma144 = ema_s(pl, 144)
    dif_pct = [dif[i]/pl[i]*100 if (not math.isnan(dif[i])) and pl[i] > 0 else 0.0 for i in range(len(pl))]
    gcf = [False]*len(pl)
    for i in range(1, len(pl)):
        if not(math.isnan(dif[i-1]) or math.isnan(dea[i-1]) or math.isnan(dif[i]) or math.isnan(dea[i])):
            if dif[i-1] <= dea[i-1] and dif[i] > dea[i]:
                gcf[i] = True
    IND[code] = {'dates': d_list, 'prices': pl, 'dif': dif, 'dea': dea, 'bar': bar,
                 'ma5': ma5, 'ma144': ma144, 'dif_pct': dif_pct, 'board': get_board(code),
                 'date_to_idx': {d: i for i, d in enumerate(d_list)}}
    gc_flag[code] = gcf
    dif_pct_mat[code] = dif_pct

# ---- 池内等权指数 (代理中证1000) + V4状态机 ----
print("计算池内等权指数 -> V4状态机...", flush=True)
idx_vals = []
for di, d in enumerate(dates_valid):
    vals = [open_data[c][d] for c in codes if d in open_data[c]]
    idx_vals.append(sum(vals)/len(vals) if vals else float('nan'))
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

# ---- 行业动量 (每日, 申万一级) ----
print("计算行业动量矩阵...", flush=True)
industries = sorted(set(code2ind[c] for c in codes))
ind_by_code = {c: code2ind[c] for c in codes}
# 行业 -> 代码列表
ind_members = defaultdict(list)
for c in codes:
    ind_members[code2ind[c]].append(c)
# 每日行业 avg DIF% 与 金叉占比
ind_avg_dif = {s: [] for s in industries}   # [date] -> avg dif_pct
ind_gc_frac = {s: [] for s in industries}   # [date] -> fraction golden cross
for di, d in enumerate(dates_valid):
    for s in industries:
        mem = ind_members[s]
        dp = []; gc = []
        for c in mem:
            i = IND[c]['date_to_idx'].get(d)
            if i is None: continue
            dp.append(dif_pct_mat[c][i])
            gc.append(1.0 if gc_flag[c][i] else 0.0)
        ind_avg_dif[s].append(sum(dp)/len(dp) if dp else 0.0)
        ind_gc_frac[s].append(sum(gc)/len(gc) if gc else 0.0)

def industry_resonance_pass(s, di, mode, param):
    """返回该行业在 di 日是否通过共振过滤"""
    if mode is None:
        return True
    if mode == 'difpos':
        return ind_avg_dif[s][di] > param
    if mode == 'topk':
        # 取所有行业 avg dif 的降序分位, 仅保留前 param 比例
        vals = [ind_avg_dif[ss][di] for ss in industries]
        vals_sorted = sorted(vals, reverse=True)
        k = max(1, int(len(vals_sorted)*param))
        th = vals_sorted[k-1]
        return ind_avg_dif[s][di] >= th
    if mode == 'gcfrac':
        return ind_gc_frac[s][di] >= param
    return True

# ============ 回测 ============
def backtest(resonance_mode=None, resonance_param=0.0, label='baseline'):
    portfolio = {}; cash = INIT_CASH; trades_log = []
    prev_state = 'flat'; peak_nav = INIT_CASH; max_dd = 0.0
    yearly = {}; sell_reason = defaultdict(int); ytb_skip = 0
    total_comm = 0.0; total_stamp = 0.0
    n = len(dates_valid)
    t0 = time.time()
    for date_idx in range(n):
        cur = dates_valid[date_idx]; yr = cur[:4]
        prev = dates_valid[date_idx-1] if date_idx >= 1 else None
        sig = prev  # T日(信号日)
        idx_state = v4_state.get(cur, 'flat')
        if idx_state == 'flat' and prev_state == 'short': idx_state = 'short'
        prev_state = idx_state
        max_pos = 50 if idx_state == 'bull' else (5 if idx_state == 'short' else 20)

        if yr not in yearly:
            nav = cash
            for c, p in portfolio.items():
                i = IND[c]['date_to_idx'].get(cur)
                if i is not None and IND[c]['prices'][i] > 0:
                    nav += IND[c]['prices'][i]*p['shares']
            yearly[yr] = {'start': nav, 'end': None, 'buy': 0, 'sell': 0, 'force_sell': 0}

        # 强卖 (超仓位)
        if len(portfolio) > max_pos:
            excess = len(portfolio) - max_pos
            for code, pos in sorted(portfolio.items(), key=lambda x: x[1]['score'])[:excess]:
                i = IND[code]['date_to_idx'].get(cur)
                if i is None: continue
                o = IND[code]['prices'][i]
                if o <= 0: continue
                cost_tax = o*pos['shares']*(COMMISSION_RATE+STAMP_TAX_RATE)
                total_comm += o*pos['shares']*COMMISSION_RATE
                total_stamp += o*pos['shares']*STAMP_TAX_RATE
                cash += o*pos['shares'] - cost_tax
                trades_log.append({'date': cur, 'code': code, 'dir': '卖出', 'price': o,
                                   'shares': pos['shares'], 'pnl': (o-pos['entry_price'])*pos['shares']-cost_tax,
                                   'reason': '超仓位', 'yr': yr})
                del portfolio[code]
                yearly[yr]['force_sell'] += 1; sell_reason['超仓位'] += 1

        # 卖出 (T日信号 -> T+1开盘成交)
        sells = []
        for code, pos in list(portfolio.items()):
            if sig is None: continue
            i = IND[code]['date_to_idx'].get(sig)
            if i is None: continue
            dif_i = IND[code]['dif'][i]; dea_i = IND[code]['dea'][i]; ma5_i = IND[code]['ma5'][i]
            c_i = IND[code]['prices'][i]
            if math.isnan(dif_i) or math.isnan(dea_i) or math.isnan(c_i) or c_i <= 0: continue
            j = IND[code]['date_to_idx'].get(cur)
            if j is None: continue
            o_j = IND[code]['prices'][j]
            if o_j <= 0 or math.isnan(o_j): o_j = c_i
            pos['hold_days'] += 1
            if c_i > pos['max_price']: pos['max_price'] = c_i
            max_dd_in_hold = (pos['max_price']-o_j)/pos['max_price'] if pos['max_price'] > 0 else 0
            pnl_pct = (o_j-pos['entry_price'])/pos['entry_price']
            reason = None
            if reason is None and pos['hold_days'] >= MIN_HOLD and pnl_pct <= STOP_LOSS:
                reason = '止损'
            if reason is None and pos['hold_days'] >= MIN_HOLD and i >= 1:
                dp = IND[code]['dif'][i-1]; ep = IND[code]['dea'][i-1]
                if not(math.isnan(dp) or math.isnan(ep)):
                    if dp > ep and dif_i <= dea_i:
                        reason = '死叉'
            if reason is None and pnl_pct > PROFIT_THRESHOLD and not math.isnan(ma5_i) and c_i < ma5_i:
                reason = '动态止盈'
            if reason is not None:
                sells.append((code, pos, o_j, reason))
        seen = set(); sells_u = []
        for s in sells:
            if s[0] not in seen:
                seen.add(s[0]); sells_u.append(s)
        for code, pos, price, reason in sorted(sells_u, key=lambda x: x[1]['score']):
            cost_tax = price*pos['shares']*(COMMISSION_RATE+STAMP_TAX_RATE)
            total_comm += price*pos['shares']*COMMISSION_RATE
            total_stamp += price*pos['shares']*STAMP_TAX_RATE
            cash += price*pos['shares'] - cost_tax
            trades_log.append({'date': cur, 'code': code, 'dir': '卖出', 'price': price,
                               'shares': pos['shares'], 'pnl': (price-pos['entry_price'])*pos['shares']-cost_tax,
                               'reason': reason, 'yr': yr})
            del portfolio[code]
            yearly[yr]['sell'] += 1; sell_reason[reason] += 1

        # 买入 (T日信号 -> T+1开盘成交 + 涨停过滤 + 行业共振过滤)
        if len(portfolio) < max_pos and cash > 50000 and sig is not None:
            nav = cash
            for c, p in portfolio.items():
                i = IND[c]['date_to_idx'].get(cur)
                if i is not None and IND[c]['prices'][i] > 0:
                    nav += IND[c]['prices'][i]*p['shares']
            total_nav = nav
            candidates = []
            sig_di = date_pos[sig]
            for code in IND:
                if code in portfolio: continue
                i = IND[code]['date_to_idx'].get(sig)
                if i is None: continue
                j = IND[code]['date_to_idx'].get(cur)
                if j is None: continue
                c_i = IND[code]['prices'][i]
                if c_i <= 0 or math.isnan(c_i): continue
                o_j = IND[code]['prices'][j]
                if o_j <= 0 or math.isnan(o_j): o_j = c_i
                dif_i = IND[code]['dif'][i]; dea_i = IND[code]['dea'][i]; ma144_i = IND[code]['ma144'][i]; bar_i = IND[code]['bar'][i]
                if math.isnan(dif_i) or math.isnan(dea_i) or math.isnan(ma144_i): continue
                if bar_i < 0: continue
                if c_i < ma144_i: continue
                # 3日内金叉
                gc_days = -1
                for off in range(3):
                    if i-off < 0: continue
                    dn = IND[code]['dif'][i-off]; dp = IND[code]['dif'][i-off-1] if i-off-1 >= 0 else float('nan')
                    en = IND[code]['dea'][i-off]; ep = IND[code]['dea'][i-off-1] if i-off-1 >= 0 else float('nan')
                    if not(math.isnan(dn) or math.isnan(dp) or math.isnan(en) or math.isnan(ep)):
                        if dp <= ep and dn > en:
                            gc_days = off; break
                if gc_days < 0: continue
                base_score = (dif_i/c_i*100)*0.3 + (bar_i/c_i*100)*0.7
                score = base_score * PENALTY.get(gc_days, 0.2)
                # 涨停过滤
                is_zt = False
                limit_p = get_limit_price(c_i, IND[code]['board'])
                if o_j >= limit_p:
                    is_zt = True
                # 行业共振过滤
                s = ind_by_code.get(code)
                if s is None or not industry_resonance_pass(s, sig_di, resonance_mode, resonance_param):
                    continue
                candidates.append((score, code, o_j, is_zt))
            candidates.sort(reverse=True)
            for score, code, o, is_zt in candidates:
                if len(portfolio) >= max_pos or cash <= 50000: break
                if is_zt:
                    ytb_skip += 1; continue
                shares = max(int(POS_SIZE*total_nav/o/100)*100, 100)
                cost = shares*o
                total_cost = cost*(1+COMMISSION_RATE)
                if total_cost > cash*0.95: continue
                total_comm += cost*COMMISSION_RATE
                cash -= total_cost
                portfolio[code] = {'shares': shares, 'entry_price': o, 'hold_days': 0,
                                   'score': score, 'max_price': o}
                trades_log.append({'date': cur, 'code': code, 'dir': '买入', 'price': o,
                                   'shares': shares, 'pnl': 0, 'reason': f'金叉gc{gc_days}', 'yr': yr})
                yearly[yr]['buy'] += 1

        nav = cash
        for c, p in portfolio.items():
            i = IND[c]['date_to_idx'].get(cur)
            if i is not None and IND[c]['prices'][i] > 0:
                nav += IND[c]['prices'][i]*p['shares']
        if nav > peak_nav: peak_nav = nav
        dd = (peak_nav-nav)/peak_nav*100
        if dd > max_dd: max_dd = dd
        yearly[yr]['end'] = nav

    final_nav = yearly[list(yearly.keys())[-1]]['end']
    total_return = (final_nav/INIT_CASH - 1)*100
    for y in yearly:
        s = yearly[y]['start']; e = yearly[y]['end']
        yearly[y]['return'] = (e/s-1)*100 if s > 0 else 0
    t1 = time.time()
    return {
        'label': label, 'final_nav': final_nav, 'total_return': total_return,
        'max_dd': max_dd, 'ratio': total_return/max_dd if max_dd > 0 else 0,
        'yearly': {y: yearly[y]['return'] for y in sorted(yearly.keys())},
        'buy': sum(yearly[y]['buy'] for y in yearly),
        'sell': sum(yearly[y]['sell'] for y in yearly),
        'force_sell': sum(yearly[y]['force_sell'] for y in yearly),
        'ytb_skip': ytb_skip, 'sell_reason': dict(sell_reason),
        'time': t1-t0,
    }

# ============ 运行对比 ============
if __name__ == '__main__':
    runs = []
    runs.append(backtest(None, 0.0, 'baseline(无行业过滤)'))
    runs.append(backtest('difpos', 0.0, '行业共振-dif>0'))
    runs.append(backtest('topk', 0.3, '行业共振-前30%行业'))
    runs.append(backtest('topk', 0.5, '行业共振-前50%行业'))
    runs.append(backtest('gcfrac', 0.10, '行业共振-金叉占比>=10%'))
    runs.append(backtest('gcfrac', 0.20, '行业共振-金叉占比>=20%'))

    print("\n" + "=" * 100)
    print(" V46 + 行业共振 回测对比 (开盘价代理口径, 数据: 开盘价+申万一级行业, 指数=池内等权代理)")
    print("=" * 100)
    hdr = f"{'策略':<26}{'总收益%':>10}{'回撤%':>9}{'收益/回撤':>11}{'买入':>7}{'卖出':>7}{'强卖':>7}{'涨停跳过':>9}"
    print(hdr); print("-" * 100)
    for r in runs:
        print(f"{r['label']:<26}{r['total_return']:>10.2f}{r['max_dd']:>9.2f}{r['ratio']:>11.2f}"
              f"{r['buy']:>7}{r['sell']:>7}{r['force_sell']:>7}{r['ytb_skip']:>9}")
    print("-" * 100)
    print("\n逐年收益%:")
    yrs = sorted({y for r in runs for y in r['yearly']})
    print(f"{'策略':<26}" + "".join(f"{y:>9}" for y in yrs))
    for r in runs:
        print(f"{r['label']:<26}" + "".join(f"{r['yearly'].get(y,0):>9.2f}" for y in yrs))
    print("-" * 100)
    for r in runs:
        print(f"\n[{r['label']}] 卖出原因: {r['sell_reason']} | 耗时{r['time']:.1f}s")

    with open('/workspace/v46_resonance_results.json', 'w') as f:
        json.dump(runs, f, ensure_ascii=False, indent=2, default=str)
    print("\n已保存 /workspace/v46_resonance_results.json")
