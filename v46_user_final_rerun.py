"""V46 v62_ytb_v19 - 实战正确版: V62_fix 信号 (T 日盘后看) + T+1 日开盘价成交

实战逻辑:
- T 日 15:00 收盘后 (盘后): 用 T 日收盘价算 MACD/EMA 信号 (实战真实, 盘后可知)
- T+1 日 09:25 集合竞价: 用开盘价成交 (实战真实)
- 涨停过滤: 候选票 T+1 日开盘价 >= 涨停价 → 跳过, 买下一名
- 卖出: T 日盘后看到死叉/止损/止盈 → T+1 日开盘价成交
- NAV: T+1 日开盘价估值 (跟买入价一致)

对比:
- V62_fix: T 日盘后用 T 日收盘价成交 (未来函数)
- V19: T 日盘后看信号, T+1 日开盘成交 (实战正确)
"""
import csv, math, openpyxl, json, sys, time
from collections import defaultdict
from datetime import datetime

FAST,SLOW,SIGNAL=10,20,9
START_DATE='2020-06-01'; END_DATE='2026-05-14'
INIT_CASH=10_000_000; POS_SIZE=0.02
MIN_HOLD=3; STOP_LOSS=-0.15; PROFIT_THRESHOLD=0.50
COMMISSION_RATE=0.0001; STAMP_TAX_RATE=0.0005
PENALTY={0:1.0, 1:0.5, 2:0.5, 3:0.5}

# === 行业共振配置 (env 驱动, 便于多方案扫描) ===
import os
RMODE = os.environ.get('RMODE', 'none')          # none / difpos
RPARAM = float(os.environ.get('RPARAM', '0.0'))   # 行业 avgDIF 阈值
_RREG = os.environ.get('RREGIME', 'none')         # none / bull / flat / short / nonbull
RESONANCE_REGIME = None if _RREG == 'none' else set(_RREG.split(','))
RESONANCE_MODE = None if RMODE == 'none' else RMODE
RESONANCE_PARAM = RPARAM
_RES_LABEL = f"res={RMODE}({RPARAM})@{_RREG}"
RES_CNT = [0]

def get_board(code):
    if code.startswith('30') or code.startswith('688'): return 'chinext'
    if code.startswith('8'): return 'bse'
    return 'main'

def get_limit_price(prev_close, board):
    pct = 0.10 if board == 'main' else (0.20 if board == 'chinext' else 0.30)
    return round(prev_close * (1 + pct), 2)

print("加载数据(用户GitHub CSV)...", flush=True)
D='/workspace/data_user/data_extract/data'
stock_data=defaultdict(dict)
with open(f'{D}/股票池_收盘价.csv', encoding='utf-8-sig') as f:
    for row in csv.DictReader(f):
        try:
            p=float(row['price'])
            if p>0: stock_data[row['code'].strip()][row['date'].strip()]=p
        except: continue

open_data=defaultdict(dict)
with open(f'{D}/股票池_开盘价.csv', encoding='utf-8-sig') as f:
    for row in csv.DictReader(f):
        try:
            o=float(row['open_qf'])
            if o>0: open_data[row['code'].strip()][row['date'].strip()]=o
        except: continue

vol_data=defaultdict(dict)
with open(f'{D}/成交量.csv', encoding='utf-8') as f:
    for row in csv.DictReader(f):
        try:
            v=float(row['volume'])
            if v>0: vol_data[row['code'].strip()][row['date'].strip()]=v
        except: continue
print(f"  成交量: {len(vol_data)}只", flush=True)

all_dates=sorted({d for code in stock_data for d in stock_data[code]})
dates_valid=[d for d in all_dates if START_DATE<=d<=END_DATE]; n=len(dates_valid)
print(f"  {len(stock_data)}只 | {n}天", flush=True)

# 指数 V4 状态机 (用新浪中证1000, 仅收盘, 干净)
import pandas as pd
zz1000 = pd.read_parquet('/workspace/data/zz1000_hfq.parquet')
idx_data={}
for d in zz1000.index:
    ds = pd.to_datetime(d).strftime('%Y-%m-%d')
    v = zz1000['close'].loc[d]
    if not pd.isna(v): idx_data[ds]=float(v)
idx_dates=sorted(idx_data.keys()); prices=[idx_data[d] for d in idx_dates]; n_idx=len(idx_dates)

def ema(s,n_):
    a=2.0/(n_+1); out=[None]*len(s)
    for i in range(len(s)):
        if s[i] is None: continue
        out[i]=s[i] if i==0 or out[i-1] is None else a*s[i]+(1-a)*out[i-1]
    return out

def dif_slope(dif, win=2):
    out=[0.0]*len(dif)
    for i in range(win,len(dif)):
        y=[dif[i-j] for j in range(win-1,-1,-1)]; xm=[j-(win-1)/2 for j in range(win)]; ym=[yi-sum(y)/win for yi in y]
        num=sum(xm[j]*ym[j] for j in range(win)); den=sum(xm[j]**2 for j in range(win))
        out[i]=num/den if den!=0 else 0
    return out

ef10=ema(prices,10); es20=ema(prices,20)
dif10=[ef10[i]-es20[i] if ef10[i] is not None and es20[i] is not None else None for i in range(n_idx)]
dea10=ema(dif10,9)
bar10=[dif10[i]-dea10[i] if dif10[i] is not None and dea10[i] is not None else None for i in range(n_idx)]
dif_pct10=[dif10[i]/prices[i]*100 if dif10[i] is not None else 0 for i in range(n_idx)]
slope10=dif_slope(dif_pct10,2)

def calc_pos10(dif_pct, slope, dea, bar, th_short, bottom):
    pos=0; result=[]; prev_bar=None
    for i in range(1, len(dif_pct)):
        rd=dif10[i] if dif10[i] is not None else 0
        rdea=dea[i] if dea[i] is not None else 0
        pd_=dif10[i-1] if dif10[i-1] is not None else 0
        pdea=dea[i-1] if dea[i-1] is not None else 0
        rdp=dif_pct[i]; rsl=slope[i]
        rbar=bar[i] if bar is not None and i < len(bar) and bar[i] is not None else 0
        gcu=pd_<=pdea and rd>rdea; gcd=pd_>=pdea and rd<pdea
        if pos==0:
            if gcu: pos=1
            elif rdp>0 and rsl>0.02: pos=1
            elif rdp<bottom and rsl>0.03: pos=1
            elif rdp<=th_short: pos=-1
        elif pos==1:
            if gcd:
                pos=0
                if rdp<=th_short: pos=-1
            elif prev_bar is not None and prev_bar < 0:
                if rdp > 0 and rsl < -0.02: pos=0
        elif pos==-1:
            if gcu: pos=1
            elif rdp<bottom and rsl>0.03: pos=1
        result.append(pos)
        prev_bar=rbar
    return result

v4p=calc_pos10(dif_pct10,slope10,dea10,bar10,0.0,-3.0)
v4_state={}
for i in range(len(idx_dates)-1):
    d=idx_dates[i+1]
    v4_state[d]='bull' if v4p[i]==1 else ('short' if v4p[i]==-1 else 'flat')

print("预计算指标...", flush=True)
IND={}

def calc_ema_s(pl,p_):
    n_=len(pl); a=2.0/(p_+1); out=[float('nan')]*n_
    fv_=next((i for i,v in enumerate(pl) if v>0),None)
    if fv_ is None: return out
    out[fv_]=pl[fv_]
    for i in range(fv_+1,n_):
        if pl[i]==0: out[i]=out[i-1] if not math.isnan(out[i-1]) else float('nan')
        else: out[i]=a*pl[i]+(1-a)*out[i-1]
    return out

for code, dp in stock_data.items():
    d_list=sorted(dp.keys())
    if len(d_list)<30: continue
    pl=[dp.get(d,0) for d in d_list]
    ol=[open_data.get(code, {}).get(d, 0) for d in d_list]
    ef_=calc_ema_s(pl,FAST); es_=calc_ema_s(pl,SLOW)
    dif_=[ef_[i]-es_[i] if not(math.isnan(ef_[i]) or math.isnan(es_[i])) else float('nan') for i in range(len(pl))]
    dea_=calc_ema_s(dif_,SIGNAL)
    bar_=[(dif_[i]-dea_[i])*2.0 if not(math.isnan(dif_[i]) or math.isnan(dea_[i])) else float('nan') for i in range(len(dif_))]
    ma5_=calc_ema_s(pl,5); ma144_=calc_ema_s(pl,144)
    vol10_=[float('nan')]*len(pl)
    for i in range(10, len(pl)):
        window=[(pl[j]/pl[j-1]-1)*100 if pl[j-1]>0 and pl[j]>0 else 0 for j in range(i-9, i+1)]
        if len(window)==10:
            mean_w=sum(window)/10
            var_w=sum((x-mean_w)**2 for x in window)/10
            vol10_[i]=var_w**0.5
    IND[code]={
        'dates':d_list,'prices':pl,'opens':ol,
        'dif':dif_,'dea':dea_,'bar':bar_,'ma5':ma5_,'ma144':ma144_,'vol10':vol10_,
        'board':get_board(code),
        'date_to_idx':{d:i for i,d in enumerate(d_list)}
    }
print(f"  IND: {len(IND)}只", flush=True)

# === 行业共振基础设施 (SW一级行业, 来自 股票池.xlsx) ===
code2ind = {}
_wb = openpyxl.load_workbook(f'{D}/股票池.xlsx', read_only=True, data_only=True)
_ws = _wb['股票池']
for _row in _ws.iter_rows(min_row=2, values_only=True):
    if _row[0] and _row[2]:
        code2ind[str(_row[0]).strip()] = str(_row[2]).strip()
_wb.close()
print(f"  行业映射: {len(code2ind)}只", flush=True)

def _dif_pct_of(code, i):
    dif = IND[code]['dif']; pr = IND[code]['prices']
    if math.isnan(dif[i]) or pr[i] <= 0: return 0.0
    return dif[i]/pr[i]*100
industries = sorted(set(code2ind[c] for c in IND if c in code2ind))
ind_members = defaultdict(list)
for c in IND:
    if c in code2ind: ind_members[code2ind[c]].append(c)
ind_avg_dif = {s: [] for s in industries}
for di2, d2 in enumerate(dates_valid):
    for s in industries:
        dps = []
        for c in ind_members[s]:
            ii = IND[c]['date_to_idx'].get(d2)
            if ii is None: continue
            dps.append(_dif_pct_of(c, ii))
        ind_avg_dif[s].append(sum(dps)/len(dps) if dps else 0.0)
date_pos = {d: i for i, d in enumerate(dates_valid)}
print(f"  行业数: {len(industries)} | 行业动量矩阵就绪", flush=True)

def industry_resonance_pass(s, di, mode, param):
    if mode is None: return True
    if mode == 'difpos': return ind_avg_dif[s][di] > param
    return True

# === 主回测 v19: T 日盘后看信号 + T+1 日开盘成交 ===
portfolio={}; cash=INIT_CASH; trades=[]
prev_state='flat'
peak_nav=INIT_CASH; max_dd=0; daily_nav=[]
yearly={}; trades_log=[]
sell_reason_counter=defaultdict(int)
ytb_skip_count=0
total_commission=0.0; total_stamp_tax=0.0

print("开始主回测 v19 (T 日盘后看信号 + T+1 日开盘价成交 + T+1 日开盘价涨停过滤)...", flush=True)
t0=time.time()

# ⭐ 关键: 我们在 T 日盘后 (15:00) 算信号, 但执行是 T+1 日开盘 (09:30)
# 所以流程是: 用 T 日数据算信号, 然后到 T+1 日去成交
# 但 backtest 是单日循环, 我们用 prev_date 的信号在 current_date 成交

for date_idx in range(n):
    current_date=dates_valid[date_idx]; yr=current_date[:4]
    prev_date=dates_valid[date_idx-1] if date_idx>=1 else None
    
    # 用 T-1 日 (即当前是 T+1 日) 的信号在 current_date 成交
    # 因为 T 日盘后看 T 日信号, T+1 日开盘成交
    signal_date = prev_date  # T 日
    if prev_state is None: prev_state='flat'
    idx_state=v4_state.get(current_date,'flat')
    if idx_state=='flat' and prev_state=='short': idx_state='short'
    prev_state=idx_state
    max_pos=50 if idx_state=='bull' else (5 if idx_state=='short' else 20)
    
    if yr not in yearly:
        nav=cash
        for c,p in portfolio.items():
            i=IND[c]['date_to_idx'].get(current_date)
            if i is not None and IND[c]['opens'][i]>0:
                nav+=IND[c]['opens'][i]*p['shares']
            elif i is not None and IND[c]['prices'][i]>0:
                nav+=IND[c]['prices'][i]*p['shares']
        yearly[yr]={'start':nav,'end':None,'buy':0,'sell':0,'force_sell':0,'ytb_skip':0}
    
    # 强卖 - T+1 日开价成交
    if len(portfolio)>max_pos:
        excess=len(portfolio)-max_pos
        for code,pos in sorted(portfolio.items(),key=lambda x:x[1]['score'])[:excess]:
            i=IND[code]['date_to_idx'].get(current_date)
            if i is None: continue
            o=IND[code]['opens'][i]
            if o<=0 or math.isnan(o): o=IND[code]['prices'][i]
            if o<=0: continue
            pnl=(o-pos['entry_price'])*pos['shares']
            cost_tax=o*pos['shares']*(COMMISSION_RATE+STAMP_TAX_RATE)
            total_commission+=o*pos['shares']*COMMISSION_RATE
            total_stamp_tax+=o*pos['shares']*STAMP_TAX_RATE
            pnl=(o-pos['entry_price'])*pos['shares']-cost_tax
            cash+=o*pos['shares']-cost_tax
            trades_log.append({'date':current_date,'code':code,'dir':'卖出','price':o,'shares':pos['shares'],'pnl':pnl,'reason':'超仓位','yr':yr,'state':idx_state})
            if code in portfolio: del portfolio[code]
            yearly[yr]['force_sell']+=1
            sell_reason_counter['超仓位']+=1
    
    # 卖出 - T 日盘后算信号, T+1 日开价成交
    sells=[]
    for code,pos in list(portfolio.items()):
        if signal_date is None: continue
        i=IND[code]['date_to_idx'].get(signal_date)  # ⭐ 用 T 日的 idx
        if i is None: continue
        # T 日 (signal_date) 的收盘价算 EMA (实战 T 日盘后)
        dif_i=IND[code]['dif'][i]; dea_i=IND[code]['dea'][i]
        ma5_i=IND[code]['ma5'][i]; c_i=IND[code]['prices'][i]
        if math.isnan(dif_i) or math.isnan(dea_i) or math.isnan(c_i) or c_i<=0: continue
        
        # T+1 日开盘价 (current_date)
        j=IND[code]['date_to_idx'].get(current_date)
        if j is None: continue
        o_j=IND[code]['opens'][j]
        if o_j<=0 or math.isnan(o_j): o_j=c_i
        c_j=IND[code]['prices'][j]
        if c_j<=0 or math.isnan(c_j): c_j=o_j
        
        pos['hold_days']+=1
        
        # T 日 (信号日) 收盘时 max_price 更新
        if c_i>pos['max_price']: pos['max_price']=c_i
        # T+1 日开盘时, max_price 用 T 日收价 (无未来)
        if c_i>pos['max_price']: pos['max_price']=c_i
        
        max_dd_in_hold=(pos['max_price']-o_j)/pos['max_price'] if pos['max_price']>0 else 0
        
        # 止损: T+1 日开价 vs entry_price
        pnl_pct=(o_j-pos['entry_price'])/pos['entry_price']
        if pos['hold_days']>=MIN_HOLD and pnl_pct<=STOP_LOSS:
            sells.append((code,pos,o_j,'止损'))
            continue
        
        # 死叉: T 日 DIF <= T 日 DEA (T 日盘后看到)
        if pos['hold_days']>=MIN_HOLD and i>=1:
            dif_prev=IND[code]['dif'][i-1]; dea_prev=IND[code]['dea'][i-1]
            if not(math.isnan(dif_prev) or math.isnan(dea_prev)):
                if dif_prev>dea_prev and dif_i<=dea_i:
                    sells.append((code,pos,o_j,'死叉'))
                    continue
        
        # 动态止盈: 浮盈 > 50% && T 日收价 < MA5
        if pnl_pct>PROFIT_THRESHOLD and not math.isnan(ma5_i) and c_i<ma5_i:
            sells.append((code,pos,o_j,'动态止盈'))
            continue
    
    # 去重
    seen_codes=set(); sells_unique=[]
    for s in sells:
        if s[0] not in seen_codes:
            seen_codes.add(s[0]); sells_unique.append(s)
    sells=sells_unique
    
    for code,pos,price,reason in sorted(sells,key=lambda x:x[1]['score']):
        pnl=(price-pos['entry_price'])*pos['shares']
        cost_tax=price*pos['shares']*(COMMISSION_RATE+STAMP_TAX_RATE)
        total_commission+=price*pos['shares']*COMMISSION_RATE
        total_stamp_tax+=price*pos['shares']*STAMP_TAX_RATE
        pnl=(price-pos['entry_price'])*pos['shares']-cost_tax
        cash+=price*pos['shares']-cost_tax
        trades_log.append({'date':current_date,'code':code,'dir':'卖出','price':price,'shares':pos['shares'],'pnl':pnl,'reason':reason,'yr':yr,'state':idx_state})
        if code in portfolio: del portfolio[code]
        yearly[yr]['sell']+=1
        sell_reason_counter[reason]+=1
    
    # 买入 - T 日盘后算金叉信号, T+1 日开价成交 + T+1 日开盘价涨停过滤
    if len(portfolio)<max_pos and cash>50000 and signal_date is not None:
        nav=cash
        for c,p in portfolio.items():
            i=IND[c]['date_to_idx'].get(current_date)
            if i is not None and IND[c]['opens'][i]>0:
                nav+=IND[c]['opens'][i]*p['shares']
            elif i is not None and IND[c]['prices'][i]>0:
                nav+=IND[c]['prices'][i]*p['shares']
        total_nav=nav
        candidates=[]
        for code in IND:
            if code in portfolio: continue
            i=IND[code]['date_to_idx'].get(signal_date)  # ⭐ T 日 idx 算信号
            if i is None: continue
            j=IND[code]['date_to_idx'].get(current_date)  # T+1 日 idx 成交
            if j is None: continue
            c_i=IND[code]['prices'][i]  # T 日收价 (算 EMA)
            if c_i<=0 or math.isnan(c_i): continue
            o_j=IND[code]['opens'][j]  # T+1 日开价 (成交)
            if o_j<=0 or math.isnan(o_j): o_j=c_i
            
            dif_i=IND[code]['dif'][i]; dea_i=IND[code]['dea'][i]; ma144_i=IND[code]['ma144'][i]; bar_i=IND[code]['bar'][i]
            if math.isnan(dif_i) or math.isnan(dea_i) or math.isnan(ma144_i): continue
            if bar_i<0: continue
            if math.isnan(ma144_i) or c_i<ma144_i: continue
            
            # 3日金叉 (T 日信号, T 日盘后可知)
            def bar_gc_3d_days(ind, idx, lookback=3):
                for off in range(lookback):
                    if idx-off<0: continue
                    dn=ind['dif'][idx-off]; dp=ind['dif'][idx-off-1] if idx-off-1>=0 else float('nan')
                    en=ind['dea'][idx-off]; ep=ind['dea'][idx-off-1] if idx-off-1>=0 else float('nan')
                    if not(math.isnan(dn) or math.isnan(dp) or math.isnan(en) or math.isnan(ep)):
                        if dp<=ep and dn>en: return off
                return -1
            gc_days=bar_gc_3d_days(IND[code], i)
            if gc_days<0: continue

            # 行业共振过滤 (regime 条件 + 行业 avgDIF 阈值)
            ind = code2ind.get(code)
            if ind is not None:
                sig_di = date_pos.get(signal_date)
                filter_active = (RESONANCE_REGIME is None) or (idx_state in RESONANCE_REGIME)
                if filter_active and not industry_resonance_pass(ind, sig_di, RESONANCE_MODE, RESONANCE_PARAM):
                    RES_CNT[0]+=1
                    continue

            _v=IND[code].get('vol10')
            vol10_i=_v[i] if _v is not None and i<len(_v) and not math.isnan(_v[i]) else 0
            
            # 量比加分: 近5日平均量比(放量加分, 不缩量不减分)
            vol_bonus=0.0
            vol_dict=vol_data.get(code)
            if vol_dict and signal_date:
                ratios=[]
                for off in range(5):
                    idx_v=date_idx-1-off
                    if idx_v>=0:
                        v_t=vol_dict.get(dates_valid[idx_v])
                        sum5=0; cnt5=0
                        for off2 in range(1,6):
                            idx5=idx_v-off2
                            if idx5>=0:
                                v5=vol_dict.get(dates_valid[idx5])
                                if v5 and v5>0: sum5+=v5; cnt5+=1
                        if cnt5>=3 and sum5/cnt5>0 and v_t and v_t>0:
                            ratios.append(v_t/(sum5/cnt5))
                if ratios:
                    avg_r=sum(ratios)/len(ratios)
                    if avg_r>=1.5: vol_bonus=0.35
                    elif avg_r>=1.2: vol_bonus=0.20
                    elif avg_r>=1.0: vol_bonus=0.10
            
            base_score=(dif_i/c_i*100)*0.3+(bar_i/c_i*100)*0.7 - 0.2*vol10_i
            score=base_score*PENALTY.get(gc_days, 0.2)*(1+vol_bonus)
            
            # 涨停过滤: T+1 日开价 >= T 日收价 × 涨停幅度
            is_zt=False
            limit_p=get_limit_price(c_i, IND[code]['board'])
            if o_j>=limit_p:
                is_zt=True
            candidates.append((score, code, o_j, gc_days, is_zt))
        
        candidates.sort(reverse=True)
        for score, code, o, gc_days, is_zt in candidates:
            if len(portfolio)>=max_pos or cash<=50000: break
            if is_zt:
                ytb_skip_count+=1; yearly[yr]['ytb_skip']+=1
                continue
            shares=max(int(POS_SIZE*total_nav/o/100)*100,100)
            cost=shares*o
            total_cost=cost*(1+COMMISSION_RATE)
            if total_cost>cash*0.95: continue
            total_commission+=cost*COMMISSION_RATE
            cash-=total_cost
            portfolio[code]={'shares':shares,'entry_price':o,'hold_days':0,'score':score,'max_price':o}
            trades_log.append({'date':current_date,'code':code,'dir':'买入','price':o,'shares':shares,'pnl':0,'reason':f'金叉gc{gc_days}','yr':yr,'state':idx_state})
            yearly[yr]['buy']+=1
    
    # NAV 估值 - T+1 日开价 (跟买入价一致)
    nav=cash
    for c,p in portfolio.items():
        i=IND[c]['date_to_idx'].get(current_date)
        if i is not None and IND[c]['opens'][i]>0:
            nav+=IND[c]['opens'][i]*p['shares']
        elif i is not None and IND[c]['prices'][i]>0:
            nav+=IND[c]['prices'][i]*p['shares']
    yearly[yr]['end']=nav
    if nav>peak_nav: peak_nav=nav
    dd=(peak_nav-nav)/peak_nav*100
    if dd>max_dd: max_dd=dd
    daily_nav.append({'date': current_date, 'nav': nav, 'yr': yr})
    
    if date_idx%200==0:
        print(f"  {date_idx}/{n} | {current_date} | nav={nav:,.0f} | cash={cash:,.0f} | pos={len(portfolio)}", flush=True)

t1=time.time()
print(f"主回测耗时: {t1-t0:.1f}秒", flush=True)

final_nav=yearly[list(yearly.keys())[-1]]['end']
total_return=(final_nav/INIT_CASH-1)*100

with open('/workspace/v62_ytb_v19_daily_nav.csv','w',newline='') as f:
    w=csv.DictWriter(f, fieldnames=['date','nav','yr'])
    w.writeheader()
    for r in daily_nav: w.writerow(r)
for y in yearly:
    s=yearly[y]['start']; e=yearly[y]['end']
    yearly[y]['return']=(e/s-1)*100 if s>0 else 0

print("\n"+"="*70)
print(f" V46 v62_v20_final + 行业共振 [{_RES_LABEL}]")
print("="*70)
print(f"{'年份':<6}{'收益%':>10}{'年初':>14}{'年末':>14}{'买入':>6}{'卖出':>6}{'强卖':>6}{'涨停跳过':>10}")
print("-"*80)
for y in sorted(yearly.keys()):
    info=yearly[y]
    print(f"{y:<6}{info['return']:>10.2f}{info['start']:>14,.0f}{info['end']:>14,.0f}{info['buy']:>6}{info['sell']:>6}{info['force_sell']:>6}{info.get('ytb_skip',0):>10}")
print("-"*80)
print(f"\n总收益: {total_return:.2f}% | max_dd: {max_dd:.2f}% | ratio: {total_return/max_dd if max_dd>0 else 0:.2f}")
print(f"final_nav: {final_nav:,.0f}")
tc=total_commission+total_stamp_tax
print(f"交易成本: 佣金={total_commission:,.0f} | 印花税={total_stamp_tax:,.0f} | 合计={tc:,.0f} | 占终值比={tc/final_nav*100:.2f}%")
print(f"买入: {sum(yearly[y]['buy'] for y in yearly)} | 卖出: {sum(yearly[y]['sell'] for y in yearly)} | 强卖: {sum(yearly[y]['force_sell'] for y in yearly)} | 涨停跳过: {ytb_skip_count}")
print(f"行业共振过滤剔除候选数: {RES_CNT[0]} (regime={_RREG}, mode={RMODE}, param={RPARAM})")
print(f"\n卖出原因分布: {dict(sell_reason_counter)}")

result={
    'label':f'V46量比+行业共振[{_RES_LABEL}]',
    'final_nav':final_nav,
    'total_return':total_return,
    'max_dd':max_dd,
    'peak_nav':peak_nav,
    'yearly':yearly,
    'buy_count':sum(yearly[y]['buy'] for y in yearly),
    'sell_count':sum(yearly[y]['sell'] for y in yearly),
    'force_sell_count':sum(yearly[y]['force_sell'] for y in yearly),
    'ytb_skip_count':ytb_skip_count,
    'sell_reason_counter':dict(sell_reason_counter),
}
_OUT=f'/workspace/v46_vol_res_{RMODE}_{str(RPARAM).replace(".","")}_{_RREG}.json'
with open(_OUT,'w') as f:
    json.dump(result, f, ensure_ascii=False, indent=2, default=str)
print(f"\n已保存到 {_OUT}")

with open('/workspace/V46_v62_v20_final_trades.csv','w', newline='') as f:
    w=csv.DictWriter(f, fieldnames=['date','code','dir','price','shares','pnl','reason','yr','state'])
    w.writeheader()
    for t in trades_log: w.writerow(t)
print(f"交易记录: /workspace/V46_v62_v20_final_trades.csv ({len(trades_log)}笔)")