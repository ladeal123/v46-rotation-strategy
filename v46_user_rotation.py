"""V46 行业轮动(rotation) 扫描 — 跑在真实量比基线上

相对 V49/V50 的「行业共振过滤」(硬剔除弱行业, 已证伪无增益):
本脚本改为「行业轮动」: 不删候选, 只通过 (a) 软加权 TILT 或 (b) 仓位配额 SLOT
把资金向强势行业集中, 资金始终在场。

机制:
- TILT: score *= 1 + k * z(行业动量) ; 动量信号: avgdif(行业avgDIF横截面z) / volratio(行业平均量比) / breadth(行业金叉广度)
- SLOT: 把 max_pos 个名额按行业动量 softmax 分配, 强势行业占更多名额, 余量溢出

基线(mode=none) 应精确复现 +557.79 / maxdd 19.43 / ratio 28.71, 作为口径校验。
"""
import csv, math, openpyxl, json, time, os
from collections import defaultdict
from datetime import datetime

FAST,SLOW,SIGNAL=10,20,9
START_DATE='2020-06-01'; END_DATE='2026-05-14'
INIT_CASH=10_000_000; POS_SIZE=0.02
MIN_HOLD=3; STOP_LOSS=-0.15; PROFIT_THRESHOLD=0.50
COMMISSION_RATE=0.0001; STAMP_TAX_RATE=0.0005
PENALTY={0:1.0, 1:0.5, 2:0.5, 3:0.5}
# 中证1000成分股(匹配后933只)
try:
    with open('/workspace/zz1000_codes.json') as _f: ZZ1000_CODES=set(json.load(_f))
except: ZZ1000_CODES=set()
try:
    with open('/workspace/zz500_codes.json') as _f: ZZ500_CODES=set(json.load(_f))
except: ZZ500_CODES=set()
print(f"  ZZ1000成分股: {len(ZZ1000_CODES)}只  ZZ500成分股: {len(ZZ500_CODES)}只", flush=True)

def get_board(code):
    if code.startswith('30') or code.startswith('688'): return 'chinext'
    if code.startswith('8'): return 'bse'
    return 'main'
def get_limit_price(prev_close, board):
    pct = 0.10 if board=='main' else (0.20 if board=='chinext' else 0.30)
    return round(prev_close*(1+pct),2)
def clip(x,lo,hi): return lo if x<lo else (hi if x>hi else x)

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

# === 指数 V4 状态机 (新浪中证1000) ===
import pandas as pd
zz1000 = pd.read_parquet('/workspace/data/zz1000_hfq.parquet')
idx_data={}
for d in zz1000.index:
    ds=pd.to_datetime(d).strftime('%Y-%m-%d'); v=zz1000['close'].loc[d]
    if not pd.isna(v): idx_data[ds]=float(v)
idx_dates=sorted(idx_data.keys()); prices=[idx_data[d] for d in idx_dates]; n_idx=len(idx_dates)
def ema(s,n_):
    a=2.0/(n_+1); out=[None]*len(s)
    for i in range(len(s)):
        if s[i] is None: continue
        out[i]=s[i] if i==0 or out[i-1] is None else a*s[i]+(1-a)*out[i-1]
    return out
def dif_slope(dif,win=2):
    out=[0.0]*len(dif)
    for i in range(win,len(dif)):
        y=[dif[i-j] for j in range(win-1,-1,-1)]; xm=[j-(win-1)/2 for j in range(win)]; ym=[yi-sum(y)/win for yi in y]
        num=sum(xm[j]*ym[j] for j in range(win)); den=sum(xm[j]**2 for j in range(win))
        out[i]=num/den if den!=0 else 0
    return out
ef10=ema(prices,10); es20=ema(prices,20)
dif10=[ef10[i]-es20[i] if ef10[i] is not None and es20[i] is not None else None for i in range(n_idx)]
dea10=ema(dif10,9); bar10=[dif10[i]-dea10[i] if dif10[i] is not None and dea10[i] is not None else None for i in range(n_idx)]
dif_pct10=[dif10[i]/prices[i]*100 if dif10[i] is not None else 0 for i in range(n_idx)]
slope10=dif_slope(dif_pct10,2)
def calc_pos10(dif_pct,slope,dea,bar,th_short,bottom):
    pos=0; result=[]; prev_bar=None
    for i in range(1,len(dif_pct)):
        rd=dif10[i] if dif10[i] is not None else 0; rdea=dea[i] if dea[i] is not None else 0
        pd_=dif10[i-1] if dif10[i-1] is not None else 0; pdea=dea[i-1] if dea[i-1] is not None else 0
        rdp=dif_pct[i]; rsl=slope[i]
        rbar=bar[i] if bar is not None and i<len(bar) and bar[i] is not None else 0
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
            elif prev_bar is not None and prev_bar<0:
                if rdp>0 and rsl<-0.02: pos=0
        elif pos==-1:
            if gcu: pos=1
            elif rdp<bottom and rsl>0.03: pos=1
        result.append(pos); prev_bar=rbar
    return result
v4p=calc_pos10(dif_pct10,slope10,dea10,bar10,0.0,-3.0)
v4_state={}
for i in range(len(idx_dates)-1):
    d=idx_dates[i+1]; v4_state[d]='bull' if v4p[i]==1 else ('short' if v4p[i]==-1 else 'flat')
if os.environ.get('DEBUG','0')=='1':
    print(f"  [DEBUG] idx_dates数={len(idx_dates)} 前5={idx_dates[:5]}", flush=True)
    for dd in idx_dates[:5]:
        print(f"    v4_state[{dd}]={v4_state.get(dd)}", flush=True)
    print(f"    v4_state[2020-06-02]={v4_state.get('2020-06-02')}", flush=True)

# === v6 状态机: index_timing V4延续 — MA250分环境 + 双V4阈值 + 权重映射 ===
# 上方: V4_A(th=0.25,bot=-2.0) + 权重(1.0,1.0,0) → bull/flat=50, short=5
# 下方: V4_B(th=-1.0,bot=-4.0) + 权重(1.0,0.0,0) → bull=50, flat/short=5
_ma250_v6=[sum(prices[max(0,i-249):i+1])/min(i+1,250) for i in range(n_idx)]
_v4a=calc_pos10(dif_pct10,slope10,dea10,bar10,0.25,-2.0)
_v4b=calc_pos10(dif_pct10,slope10,dea10,bar10,-1.0,-4.0)
_v6_state={}
for i in range(len(idx_dates)-1):
    d=idx_dates[i+1]; _v4=_v4a[i] if prices[i]>_ma250_v6[i] else _v4b[i]  # 上方用V4_A, 下方用V4_B
    if prices[i]>_ma250_v6[i]:
        # 上方: 权重(1.0,0.4,0) → bull=50, flat=20, short=5
        _v6_state[d]='bull' if _v4==1 else ('flat' if _v4==0 else 'short')
    else:
        # 下方: 权重(1.0,0.0,0) → bull=50, flat/short=5
        _v6_state[d]='bull' if _v4==1 else 'short'
print(f"  v6状态: bull={sum(1 for v in _v6_state.values() if v=='bull')} "
      f"flat={sum(1 for v in _v6_state.values() if v=='flat')} "
      f"short={sum(1 for v in _v6_state.values() if v=='short')} 天", flush=True)

# === v5 择时信号: MACD(40,80,32) DIF%五区 + MA250_5dOLS分状态 ===
if len(prices)>250:
    _ef40=ema(prices,40); _es80=ema(prices,80)
    _dif_v5=[_ef40[i]-_es80[i] if (_ef40[i] is not None and _es80[i] is not None) else float('nan') for i in range(n_idx)]
    _difpct_v5=[_dif_v5[i]/prices[i]*100 if (not math.isnan(_dif_v5[i]) and prices[i]>0) else 0.0 for i in range(n_idx)]
    _slp_v5=dif_slope(_difpct_v5,2)
    _ma250=[sum(prices[max(0,i-249):i+1])/min(i+1,250) for i in range(n_idx)]
    _ma5s=dif_slope(_ma250,5)
    _b1,_b2,_b3,_b4=-3.9,-0.4,0.7,3.9
    _zones=[(-1e9,_b1),(_b1,_b2),(_b2,_b3),(_b3,_b4),(_b4,1e9)]
    _BULL_ENTRY=[0.023,0.110,0.069,0.014,0.039]
    _BULL_EXIT=[-0.026,0.023,0.011,-0.044,-0.096]
    _BEAR_ENTRY=[0.052,0.072,0.056,0.038,0.029]
    _BEAR_EXIT=[-0.016,0.031,-0.011,-0.023,-0.053]
    _v5pos=0; _v5sig={}
    for i in range(1,n_idx):
        _rdp=_difpct_v5[i]; _rs=_slp_v5[i]
        _bull=1 if _ma5s[i]>=0 else 0
        _entry=_BULL_ENTRY if _bull else _BEAR_ENTRY
        _exit_=_BULL_EXIT if _bull else _BEAR_EXIT
        if _v5pos==0:
            for _k,(_lo,_hi) in enumerate(_zones):
                if _lo<=_rdp<_hi and _rs>_entry[_k]: _v5pos=1; break
        else:
            for _k,(_lo,_hi) in enumerate(_zones):
                if _lo<=_rdp<_hi and _rs<_exit_[_k]: _v5pos=0; break
        _v5sig[idx_dates[i]]=1 if _v5pos>0 else 0
    print(f"  v5信号: 持仓天数={sum(_v5sig.values())}/{n_idx-1} ({sum(_v5sig.values())/(n_idx-1)*100:.0f}%)", flush=True)
else:
    _v5sig={}
    print("  v5信号: 数据不足跳过", flush=True)

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
for code,dp in stock_data.items():
    d_list=sorted(dp.keys())
    if len(d_list)<30: continue
    pl=[dp.get(d,0) for d in d_list]
    ol=[open_data.get(code,{}).get(d,0) for d in d_list]
    ef_=calc_ema_s(pl,FAST); es_=calc_ema_s(pl,SLOW)
    dif_=[ef_[i]-es_[i] if not(math.isnan(ef_[i]) or math.isnan(es_[i])) else float('nan') for i in range(len(pl))]
    dea_=calc_ema_s(dif_,SIGNAL)
    bar_=[(dif_[i]-dea_[i])*2.0 if not(math.isnan(dif_[i]) or math.isnan(dea_[i])) else float('nan') for i in range(len(dif_))]
    ma5_=calc_ema_s(pl,5); ma144_=calc_ema_s(pl,144)
    vol10_=[float('nan')]*len(pl)
    for i in range(10,len(pl)):
        window=[(pl[j]/pl[j-1]-1)*100 if pl[j-1]>0 and pl[j]>0 else 0 for j in range(i-9,i+1)]
        if len(window)==10:
            mw=sum(window)/10; vol10_[i]=(sum((x-mw)**2 for x in window)/10)**0.5
    IND[code]={'dates':d_list,'prices':pl,'opens':ol,'dif':dif_,'dea':dea_,'bar':bar_,'ma5':ma5_,'ma144':ma144_,'vol10':vol10_,'board':get_board(code),'date_to_idx':{d:i for i,d in enumerate(d_list)}}
print(f"  IND: {len(IND)}只", flush=True)

# === 行业映射 (SW一级, 31行业) ===
code2ind={}
_wb=openpyxl.load_workbook(f'{D}/股票池.xlsx', read_only=True, data_only=True); _ws=_wb['股票池']
for _row in _ws.iter_rows(min_row=2, values_only=True):
    if _row[0] and _row[2]: code2ind[str(_row[0]).strip()]=str(_row[2]).strip()
_wb.close()
print(f"  行业映射: {len(code2ind)}只", flush=True)
industries=sorted(set(code2ind[c] for c in IND if c in code2ind))
ind_members=defaultdict(list)
for c in IND:
    if c in code2ind: ind_members[code2ind[c]].append(c)
date_pos={d:i for i,d in enumerate(dates_valid)}

def _dif_pct_of(code,i):
    dif=IND[code]['dif']; pr=IND[code]['prices']
    if math.isnan(dif[i]) or pr[i]<=0: return 0.0
    return dif[i]/pr[i]*100

# 行业 avgDIF 矩阵
ind_avg_dif={s:[] for s in industries}
for di2,d2 in enumerate(dates_valid):
    for s in industries:
        dps=[]
        for c in ind_members[s]:
            ii=IND[c]['date_to_idx'].get(d2)
            if ii is None: continue
            dps.append(_dif_pct_of(c,ii))
        ind_avg_dif[s].append(sum(dps)/len(dps) if dps else 0.0)

# 每只股票的 5日均量比 序列 (用于行业平均量比)
print("预计算行业动量信号 (avgDIF / 量比 / 广度)...", flush=True)
vr5={}  # code -> list over dates_valid
gc_flag={}  # code -> list over dates_valid (1/0)
for code in IND:
    d_list=IND[code]['dates']; vd=vol_data.get(code); arr=[float('nan')]*n; gf=[0]*n
    for di,d in enumerate(dates_valid):
        i=IND[code]['date_to_idx'].get(d)
        if i is None: continue
        # 5日均量比
        if vd:
            v_t=vd.get(d)
            if v_t and v_t>0:
                sum5=0; cnt5=0
                for off2 in range(1,6):
                    dd=dates_valid[di-off2] if di-off2>=0 else None
                    if dd is None: break
                    v5=vd.get(dd)
                    if v5 and v5>0: sum5+=v5; cnt5+=1
                if cnt5>=3 and sum5/cnt5>0: arr[di]=v_t/(sum5/cnt5)
        # 金叉标记
        if i>=1 and not math.isnan(IND[code]['dif'][i]) and not math.isnan(IND[code]['dif'][i-1]) \
           and not math.isnan(IND[code]['dea'][i]) and not math.isnan(IND[code]['dea'][i-1]):
            if IND[code]['dif'][i-1]<=IND[code]['dea'][i-1] and IND[code]['dif'][i]>IND[code]['dea'][i]:
                gf[di]=1
    vr5[code]=arr; gc_flag[code]=gf

# 行业平均量比 / 行业金叉广度
ind_avg_vr={s:[] for s in industries}; ind_breadth={s:[] for s in industries}
for di in range(n):
    for s in industries:
        vrs=[]; gcs=[]
        for c in ind_members[s]:
            v=vr5[c][di]
            if not math.isnan(v): vrs.append(v)
            gcs.append(gc_flag[c][di])
        ind_avg_vr[s].append(sum(vrs)/len(vrs) if vrs else 0.0)
        ind_breadth[s].append(sum(gcs)/len(gcs) if gcs else 0.0)

# 预计算候选 (金叉 + ma144 + bar>0 + 开价>0), 含量比加分, 与引擎口径一致
print("预计算候选池(金叉+量比)...", flush=True)
def bar_gc_3d_days(ind, idx, lookback=3):
    for off in range(lookback):
        if idx-off<0: continue
        dn=ind['dif'][idx-off]; dp=ind['dif'][idx-off-1] if idx-off-1>=0 else float('nan')
        en=ind['dea'][idx-off]; ep=ind['dea'][idx-off-1] if idx-off-1>=0 else float('nan')
        if not(math.isnan(dn) or math.isnan(dp) or math.isnan(en) or math.isnan(ep)):
            if dp<=ep and dn>en: return off
    return -1

cand_by_dateidx=[None]*n
for date_idx in range(1,n):
    prev_date=dates_valid[date_idx-1]; current_date=dates_valid[date_idx]
    cands=[]
    for code in IND:
        i=IND[code]['date_to_idx'].get(prev_date)
        if i is None: continue
        j=IND[code]['date_to_idx'].get(current_date)
        if j is None: continue
        c_i=IND[code]['prices'][i]
        if c_i<=0 or math.isnan(c_i): continue
        o_j=IND[code]['opens'][j]
        if o_j<=0 or math.isnan(o_j): o_j=c_i
        dif_i=IND[code]['dif'][i]; dea_i=IND[code]['dea'][i]; ma144_i=IND[code]['ma144'][i]; bar_i=IND[code]['bar'][i]
        if math.isnan(dif_i) or math.isnan(dea_i) or math.isnan(ma144_i): continue
        if bar_i<0: continue
        if c_i<ma144_i: continue
        gc_days=bar_gc_3d_days(IND[code], i)
        if gc_days<0: continue
        v=IND[code].get('vol10'); vol10_i=v[i] if v is not None and not math.isnan(v[i]) else 0
        # 量比加分 (与引擎一致)
        vol_bonus=0.0; vd=vol_data.get(code)
        if vd:
            ratios=[]
            for off in range(5):
                idx_v=date_idx-1-off
                if idx_v>=0:
                    v_t=vd.get(dates_valid[idx_v]); sum5=0; cnt5=0
                    for off2 in range(1,6):
                        idx5=idx_v-off2
                        if idx5>=0:
                            v5=vd.get(dates_valid[idx5])
                            if v5 and v5>0: sum5+=v5; cnt5+=1
                    if cnt5>=3 and sum5/cnt5>0 and v_t and v_t>0: ratios.append(v_t/(sum5/cnt5))
            if ratios:
                avg_r=sum(ratios)/len(ratios)
                if avg_r>=1.5: vol_bonus=0.35
                elif avg_r>=1.2: vol_bonus=0.20
                elif avg_r>=1.0: vol_bonus=0.10
        base_score=(dif_i/c_i*100)*0.3+(bar_i/c_i*100)*0.7 - 0.2*vol10_i
        score_vol=base_score*PENALTY.get(gc_days,0.2)*(1+vol_bonus)
        # consec 三因子: 超跌/MA144斜率/连涨
        rev_bonus=0.0; rev_raw=0.0
        if i>=10:
            look_prices=IND[code]['prices'][i-9:i+1]
            valid_p=[p for p in look_prices if not math.isnan(p) and p>0]
            if len(valid_p)>=5:
                peak=max(valid_p); cur=valid_p[-1]; drawdown=(peak-cur)/peak
                rev_raw=drawdown
                if drawdown>0.15: rev_bonus=0.15
                elif drawdown>0.08: rev_bonus=0.08
        slope_bonus=0.0; slope_raw=0.0
        if not math.isnan(ma144_i):
            if i>=5:
                ma144_5=IND[code]['ma144'][i-5]
                if not math.isnan(ma144_5) and ma144_5>0:
                    slope=(ma144_i-ma144_5)/ma144_5; slope_raw=slope
                    if slope>0.005: slope_bonus=0.10
                    elif slope>0.001: slope_bonus=0.05
                    elif slope<-0.001: slope_bonus=-0.05
            elif ma144_i>0:
                slope=(ma144_i-c_i)/c_i; slope_raw=slope
                if slope>0: slope_bonus=0.10
        consec_bonus=0.0; consec_raw=0
        if i>=5:
            consec_up=0
            for off in range(1,11):
                if i-off<0: break
                p_today=IND[code]['prices'][i-off+1]; p_yest=IND[code]['prices'][i-off]
                if not (math.isnan(p_today) or math.isnan(p_yest) or p_today<=0 or p_yest<=0):
                    if p_today>p_yest: consec_up+=1
                    else: break
            consec_raw=consec_up
            if consec_up>=7: consec_bonus=-0.10
            elif consec_up>=5: consec_bonus=-0.05
            elif consec_up>=3: consec_bonus=-0.02
        score_consec=base_score*PENALTY.get(gc_days,0.2)*(1+vol_bonus+rev_bonus+slope_bonus+consec_bonus)
        limit_p=get_limit_price(c_i, IND[code]['board'])
        is_zt = o_j>=limit_p
        ind=code2ind.get(code)
        cands.append({'code':code,'ind':ind,'score_vol':score_vol,'score_consec':score_consec,
                      'base_score':base_score,'gc':gc_days,'vol_bonus':vol_bonus,
                      'rev_bonus':rev_bonus,'slope_bonus':slope_bonus,'consec_bonus':consec_bonus,
                      'rev_raw':rev_raw,'slope_raw':slope_raw,'consec_raw':consec_raw,
                      'o':o_j,'is_zt':is_zt})
    cand_by_dateidx[date_idx]=cands

# 行业动量 z值 (avgDIF) 与 softmax 权重 预计算
ind_z={s:[] for s in industries}
for di in range(n):
    vals=[ind_avg_dif[s][di] for s in industries]
    m=sum(vals)/len(vals); sd=(sum((x-m)**2 for x in vals)/len(vals))**0.5
    sd=sd if sd>1e-9 else 1.0
    for s in industries: ind_z[s].append((ind_avg_dif[s][di]-m)/sd)

def industry_signal_at(sig_di, signal):
    if signal=='avgdif': return {s:ind_z[s][sig_di] for s in industries}
    if signal=='indmacd': return {s:ind_z_macd[s][sig_di] for s in industries}
    if signal=='fwdret': return {s:ind_z_fwd[s][sig_di] for s in industries}
    if signal=='amt': return {s:ind_z_amt[s][sig_di] for s in industries}
    if signal=='amtr': return {s:ind_z_amtr[s][sig_di] for s in industries}
    if signal=='breadth': return {s:ind_z_breadth[s][sig_di] for s in industries}
    if signal=='breadth_d': return {s:ind_z_breadth_d[s][sig_di] for s in industries}
    if signal=='abn': return {s:ind_z_abn[s][sig_di] for s in industries}
    if signal=='conc': return {s:ind_z_conc[s][sig_di] for s in industries}
    if signal=='abn_conc': return {s:ind_abn_conc[s][sig_di] for s in industries}
    if signal=='conc10': return {s:ind_z_conc10[s][sig_di] for s in industries}
    if signal=='conc20': return {s:ind_z_conc20[s][sig_di] for s in industries}
    if signal=='avgdif_conc10': return {s:ind_avgdif_conc10[s][sig_di] for s in industries}
    if signal=='avgdif_conc20': return {s:ind_avgdif_conc20[s][sig_di] for s in industries}
    if signal=='volratio': return {s:ind_avg_vr[s][sig_di] for s in industries}
    return {}

# === 替代行业信号源 (V54): 行业指数MACD / 前瞻行业收益(诊断) ===
# 行业等权指数(成员收盘价均值) 的 MACD(10,20,9) DIF, 横截面z
ind_price={s:[] for s in industries}
for _di2,_d2 in enumerate(dates_valid):
    for s in industries:
        _ps=[]
        for c in ind_members[s]:
            _ii=IND[c]['date_to_idx'].get(_d2)
            if _ii is not None and IND[c]['prices'][_ii]>0: _ps.append(IND[c]['prices'][_ii])
        ind_price[s].append(sum(_ps)/len(_ps) if _ps else float('nan'))
ind_dif_macd={}
for s in industries:
    _pl=ind_price[s]
    _ef=calc_ema_s(_pl,FAST); _es=calc_ema_s(_pl,SLOW)
    ind_dif_macd[s]=[(_ef[i]-_es[i]) if not(math.isnan(_ef[i]) or math.isnan(_es[i])) else float('nan') for i in range(n)]
ind_z_macd={s:[] for s in industries}
for _di in range(n):
    _vals=[ind_dif_macd[s][_di] for s in industries]
    _clean=[v for v in _vals if not math.isnan(v)]
    _m=sum(_clean)/len(_clean) if _clean else 0.0
    _sd=(sum((x-_m)**2 for x in _clean)/len(_clean))**0.5 if _clean else 1.0
    _sd=_sd if _sd>1e-9 else 1.0
    for s in industries:
        v=ind_dif_macd[s][_di]
        ind_z_macd[s].append((v-_m)/_sd if not math.isnan(v) else 0.0)
# === 行业长周期共振(硬过滤用, V60) — MA30>MA60 / MA45>MA90 ===
def _reson_ma(arr, n):
    return [sum(arr[max(0,i-n+1):i+1])/min(i+1,n) for i in range(len(arr))]
ind_reson60={s:[] for s in industries}
ind_reson90={s:[] for s in industries}
for s in industries:
    _p=ind_price[s]
    _ma30=_reson_ma(_p,30); _ma60=_reson_ma(_p,60)
    _ma45=_reson_ma(_p,45); _ma90=_reson_ma(_p,90)
    for di in range(n):
        ind_reson60[s].append(1.0 if not math.isnan(_ma30[di]) and not math.isnan(_ma60[di]) and _ma30[di]>_ma60[di] else 0.0)
        ind_reson90[s].append(1.0 if not math.isnan(_ma45[di]) and not math.isnan(_ma90[di]) and _ma45[di]>_ma90[di] else 0.0)
RESON_MAP={60:ind_reson60, 90:ind_reson90}
# 前瞻行业收益(含未来1日, 仅作上界诊断, 非可交易信号)
ind_fwd_ret={s:[] for s in industries}
for _di in range(n-1):
    for s in industries:
        _rs=[]
        for c in ind_members[s]:
            _ii=IND[c]['date_to_idx'].get(dates_valid[_di]); _jj=IND[c]['date_to_idx'].get(dates_valid[_di+1])
            if _ii is not None and _jj is not None and IND[c]['prices'][_ii]>0 and IND[c]['prices'][_jj]>0:
                _rs.append(IND[c]['prices'][_jj]/IND[c]['prices'][_ii]-1)
        ind_fwd_ret[s].append(sum(_rs)/len(_rs) if _rs else 0.0)
for s in industries: ind_fwd_ret[s].append(0.0)
ind_z_fwd={s:[] for s in industries}
for _di in range(n):
    _vals=[ind_fwd_ret[s][_di] for s in industries]
    _m=sum(_vals)/len(_vals); _sd=(sum((x-_m)**2 for x in _vals)/len(_vals))**0.5
    _sd=_sd if _sd>1e-9 else 1.0
    for s in industries: ind_z_fwd[s].append((ind_fwd_ret[s][_di]-_m)/_sd)

# === 替代行业信号源 (V55): 资金强度 / 赚钱效应(广度) — 均从现有 OHLCV 派生 ===
# 行业成交额(资金强度) = Σ(成交量 × 收盘) 各成员, 横截面 z
ind_amt={s:[] for s in industries}
for _di2,_d2 in enumerate(dates_valid):
    for s in industries:
        _a=0.0
        for c in ind_members[s]:
            _ii=IND[c]['date_to_idx'].get(_d2)
            if _ii is not None and IND[c]['prices'][_ii]>0:
                _v=vol_data.get(c,{}).get(_d2)
                if _v and _v>0: _a += _v*IND[c]['prices'][_ii]
        ind_amt[s].append(_a)
ind_z_amt={s:[] for s in industries}
for _di in range(n):
    _vals=[ind_amt[s][_di] for s in industries]
    _m=sum(_vals)/len(_vals); _sd=(sum((x-_m)**2 for x in _vals)/len(_vals))**0.5
    _sd=_sd if _sd>1e-9 else 1.0
    for s in industries: ind_z_amt[s].append((ind_amt[s][_di]-_m)/_sd)
# 资金流入加速 = 行业成交额 5日变化率, 横截面 z
ind_amtr={s:[] for s in industries}
for _di in range(n):
    for s in industries:
        _a0=ind_amt[s][_di-5] if _di>=5 else ind_amt[s][0]
        _a1=ind_amt[s][_di]
        ind_amtr[s].append((_a1/_a0-1) if _a0>0 else 0.0)
ind_z_amtr={s:[] for s in industries}
for _di in range(n):
    _vals=[ind_amtr[s][_di] for s in industries]
    _m=sum(_vals)/len(_vals); _sd=(sum((x-_m)**2 for x in _vals)/len(_vals))**0.5
    _sd=_sd if _sd>1e-9 else 1.0
    for s in industries: ind_z_amtr[s].append((ind_amtr[s][_di]-_m)/_sd)
# 行业赚钱效应 = 上涨股占比(breadth), 及 改善(Δbreadth, 5日)
ind_breadth={s:[] for s in industries}
for _di in range(1,n):
    for s in industries:
        _up=0; _tot=0
        for c in ind_members[s]:
            _ii=IND[c]['date_to_idx'].get(dates_valid[_di]); _jj=IND[c]['date_to_idx'].get(dates_valid[_di-1])
            if _ii is not None and _jj is not None and IND[c]['prices'][_ii]>0 and IND[c]['prices'][_jj]>0:
                _tot+=1
                if IND[c]['prices'][_ii]>IND[c]['prices'][_jj]: _up+=1
        ind_breadth[s].append(_up/_tot if _tot>0 else 0.0)
for s in industries: ind_breadth[s].insert(0,0.0)
ind_breadth_d={s:[] for s in industries}
for _di in range(n):
    for s in industries:
        _b0=ind_breadth[s][_di-5] if _di>=5 else ind_breadth[s][0]
        ind_breadth_d[s].append(ind_breadth[s][_di]-_b0)
ind_z_breadth={s:[] for s in industries}
for _di in range(n):
    _vals=[ind_breadth[s][_di] for s in industries]
    _m=sum(_vals)/len(_vals); _sd=(sum((x-_m)**2 for x in _vals)/len(_vals))**0.5
    _sd=_sd if _sd>1e-9 else 1.0
    for s in industries: ind_z_breadth[s].append((ind_breadth[s][_di]-_m)/_sd)
ind_z_breadth_d={s:[] for s in industries}
for _di in range(n):
    _vals=[ind_breadth_d[s][_di] for s in industries]
    _m=sum(_vals)/len(_vals); _sd=(sum((x-_m)**2 for x in _vals)/len(_vals))**0.5
    _sd=_sd if _sd>1e-9 else 1.0
    for s in industries: ind_z_breadth_d[s].append((ind_breadth_d[s][_di]-_m)/_sd)

# === 替代行业信号源 (V56): 异常换手(去规模偏) + 成交集中度(资金聚焦) — 修复 V55 总量规模偏/无集中度 缺陷 ===
# 异常换手 = 行业成交额 / 自身 MA20 - 1 (比自身常态更热, 去行业规模偏), 横截面 z
ind_amt_ma20={s:[] for s in industries}
for s in industries:
    _arr=ind_amt[s]
    for _di in range(n):
        _a0=max(0,_di-20+1); _win=_arr[_a0:_di+1]
        ind_amt_ma20[s].append(sum(_win)/len(_win))
ind_abn={s:[] for s in industries}
for s in industries:
    for _di in range(n):
        _m=ind_amt_ma20[s][_di]
        ind_abn[s].append((ind_amt[s][_di]/_m-1) if _m>0 else 0.0)
ind_z_abn={s:[] for s in industries}
for _di in range(n):
    _vals=[ind_abn[s][_di] for s in industries]
    _m=sum(_vals)/len(_vals); _sd=(sum((x-_m)**2 for x in _vals)/len(_vals))**0.5
    _sd=_sd if _sd>1e-9 else 1.0
    for s in industries: ind_z_abn[s].append((ind_abn[s][_di]-_m)/_sd)
# 成交集中度 = HHI归一化 (Σ(个股权额占比)^2 - 1/N)/(1 - 1/N) ∈[0,1], 资金聚焦龙头, 去行业成员数偏, 横截面 z
ind_conc={s:[] for s in industries}
for _di in range(n):
    for s in industries:
        _ts=[]
        for c in ind_members[s]:
            _ii=IND[c]['date_to_idx'].get(dates_valid[_di])
            if _ii is not None and IND[c]['prices'][_ii]>0:
                _v=vol_data.get(c,{}).get(dates_valid[_di])
                if _v and _v>0: _ts.append(_v*IND[c]['prices'][_ii])
        _tot=sum(_ts)
        if _tot>0 and _ts:
            _N=len(_ts); _hhi=sum((x/_tot)**2 for x in _ts)
            ind_conc[s].append((_hhi-1.0/_N)/(1.0-1.0/_N) if _N>1 else 0.0)
        else:
            ind_conc[s].append(0.0)
ind_z_conc={s:[] for s in industries}
for _di in range(n):
    _vals=[ind_conc[s][_di] for s in industries]
    _m=sum(_vals)/len(_vals); _sd=(sum((x-_m)**2 for x in _vals)/len(_vals))**0.5
    _sd=_sd if _sd>1e-9 else 1.0
    for s in industries: ind_z_conc[s].append((ind_conc[s][_di]-_m)/_sd)
# 组合: 异常换手 z + 集中度 z (正交信号相加, 不双重计数) — 对齐"资金聚焦成交上升的热点行业"
ind_abn_conc={s:[] for s in industries}
for _di in range(n):
    for s in industries: ind_abn_conc[s].append(ind_z_abn[s][_di]+ind_z_conc[s][_di])

# === V57: conc 滚动版(降噪, 抓"阶段性热度") + 与 avgdif 正交组合 ===
# 10日滚动HHI均值, z
ind_conc10={s:[] for s in industries}
for s in industries:
    for _di in range(n):
        _a0=max(0,_di-10+1); _win=ind_conc[s][_a0:_di+1]
        ind_conc10[s].append(sum(_win)/len(_win))
ind_z_conc10={s:[] for s in industries}
for _di in range(n):
    _vals=[ind_conc10[s][_di] for s in industries]
    _m=sum(_vals)/len(_vals); _sd=(sum((x-_m)**2 for x in _vals)/len(_vals))**0.5
    _sd=_sd if _sd>1e-9 else 1.0
    for s in industries: ind_z_conc10[s].append((ind_conc10[s][_di]-_m)/_sd)
# 20日滚动HHI均值, z
ind_conc20={s:[] for s in industries}
for s in industries:
    for _di in range(n):
        _a0=max(0,_di-20+1); _win=ind_conc[s][_a0:_di+1]
        ind_conc20[s].append(sum(_win)/len(_win))
ind_z_conc20={s:[] for s in industries}
for _di in range(n):
    _vals=[ind_conc20[s][_di] for s in industries]
    _m=sum(_vals)/len(_vals); _sd=(sum((x-_m)**2 for x in _vals)/len(_vals))**0.5
    _sd=_sd if _sd>1e-9 else 1.0
    for s in industries: ind_z_conc20[s].append((ind_conc20[s][_di]-_m)/_sd)
# 组合: avgdif z + conc z (正交信号相加, 不双重计数)
ind_avgdif_conc10={s:[] for s in industries}
for _di in range(n):
    for s in industries: ind_avgdif_conc10[s].append(ind_z[s][_di]+ind_z_conc10[s][_di])
ind_avgdif_conc20={s:[] for s in industries}
for _di in range(n):
    for s in industries: ind_avgdif_conc20[s].append(ind_z[s][_di]+ind_z_conc20[s][_di])

# === 机制诊断: 行业倾斜信号(avgdif横截面z) 与 各 consec 因子的横截面相关性 ===
# 目的: 解释「为何 MA144斜率 与 行业倾斜 冲突」。若二者高度相关=对同一趋势信号的双重计数。
def _pearson(a,b):
    n=len(a)
    if n<2: return 0.0
    ma=sum(a)/n; mb=sum(b)/n
    cov=sum((a[i]-ma)*(b[i]-mb) for i in range(n))
    va=sum((x-ma)**2 for x in a); vb=sum((x-mb)**2 for x in b)
    return cov/((va*vb)**0.5) if va>0 and vb>0 else 0.0
_diag={'ind_z':[], 'slope':[], 'rev':[], 'consec':[], 'vol':[], 'amt':[], 'amtr':[], 'breadth':[], 'breadth_d':[], 'abn':[], 'conc':[], 'abn_conc':[], 'conc10':[], 'conc20':[], 'avgdif_conc10':[], 'avgdif_conc20':[]}
for _di in range(1,n):
    if cand_by_dateidx[_di] is None: continue
    _pd=dates_valid[_di-1]; _sig=date_pos[_pd]
    for _cd in cand_by_dateidx[_di]:
        if _cd['ind'] is None: continue
        _z=ind_z[_cd['ind']][_sig]
        _diag['ind_z'].append(_z); _diag['slope'].append(_cd['slope_bonus'])
        _diag['rev'].append(_cd['rev_bonus']); _diag['consec'].append(_cd['consec_bonus'])
        _diag['vol'].append(_cd['vol_bonus'])
        _diag['amt'].append(ind_z_amt[_cd['ind']][_sig])
        _diag['amtr'].append(ind_z_amtr[_cd['ind']][_sig])
        _diag['breadth'].append(ind_z_breadth[_cd['ind']][_sig])
        _diag['breadth_d'].append(ind_z_breadth_d[_cd['ind']][_sig])
        _diag['abn'].append(ind_z_abn[_cd['ind']][_sig])
        _diag['conc'].append(ind_z_conc[_cd['ind']][_sig])
        _diag['abn_conc'].append(ind_abn_conc[_cd['ind']][_sig])
        _diag['conc10'].append(ind_z_conc10[_cd['ind']][_sig])
        _diag['conc20'].append(ind_z_conc20[_cd['ind']][_sig])
        _diag['avgdif_conc10'].append(ind_avgdif_conc10[_cd['ind']][_sig])
        _diag['avgdif_conc20'].append(ind_avgdif_conc20[_cd['ind']][_sig])
print("\n机制诊断: 行业倾斜信号(avgdif-z) × consec 因子 (候选池横截面 n=%d)"%len(_diag['ind_z']), flush=True)
_N=len(_diag['ind_z'])
for _k in ['slope','rev','consec','vol']:
    _r=_pearson(_diag['ind_z'], _diag[_k])
    _grp_p=[_diag['ind_z'][i] for i in range(_N) if _diag[_k][i]>0]
    _grp_n=[_diag['ind_z'][i] for i in range(_N) if _diag[_k][i]<0]
    _grp_z=[_diag['ind_z'][i] for i in range(_N) if _diag[_k][i]==0]
    _mp=sum(_grp_p)/len(_grp_p) if _grp_p else float('nan')
    _mn=sum(_grp_n)/len(_grp_n) if _grp_n else float('nan')
    _act=100.0*(len(_grp_p)+len(_grp_n))/_N
    print(f"  {_k:<6} corr(z)={_r:+.3f} | 触发占比={_act:5.1f}% | 加分样本均值z={_mp:+.2f} | 减分样本均值z={_mn:+.2f}", flush=True)
# V55/V56 正交性诊断: 新信号 × (slope同维冲突 / avgdif冗余)
print("\nV55/V56 正交性诊断: 新信号 × (slope=同维冲突源 / avgdif=当前倾斜)", flush=True)
for _k in ['amt','amtr','breadth','breadth_d','abn','conc','abn_conc','conc10','conc20','avgdif_conc10','avgdif_conc20']:
    _rs=_pearson(_diag[_k], _diag['slope'])
    _ra=_pearson(_diag[_k], _diag['ind_z'])
    print(f"  {_k:<10} corr(信号,slope)={_rs:+.3f} | corr(信号,avgdif)={_ra:+.3f}", flush=True)

# === 单配置回测 ===
GLOBAL_DUMP=[]
def simulate(cfg, start_idx=0, end_idx=None):
    global GLOBAL_DUMP, TILT_W, BETA_W
    GLOBAL_DUMP=[]
    # 自设置全局倾斜/softmax权重, 使直接调用 simulate()(如 WALK 块) 不依赖残留的全局 TILT_W
    TILT_W=cfg.get('tilt',0.0); BETA_W=cfg.get('beta',0.0)
    mode=cfg['mode']; signal=cfg.get('signal'); tilt=cfg.get('tilt',0.0); beta=cfg.get('beta',0.0)
    signal2=cfg.get('signal2'); tilt2=cfg.get('tilt2',0.0)
    reson=cfg.get('reson',0); reson_k=cfg.get('reson_k',0.0)  # 行业共振过滤: 0=关闭, 60/90=周期
    # V4 状态依赖参数(动态持有期/止损/止盈)
    _v4_hold={'bull':cfg.get('hold_bull',3),'flat':cfg.get('hold_flat',3),'short':cfg.get('hold_short',3)}
    _v4_stop={'bull':cfg.get('stop_bull',-0.15),'flat':cfg.get('stop_flat',-0.15),'short':cfg.get('stop_short',-0.15)}
    _v4_profit={'bull':cfg.get('profit_bull',0.50),'flat':cfg.get('profit_flat',0.50),'short':cfg.get('profit_short',0.50)}
    use_consec=cfg.get('consec',False)
    use_rev=cfg.get('use_rev',use_consec); use_slope=cfg.get('use_slope',use_consec); use_consec_b=cfg.get('use_consec_b',use_consec)
    # consec 因子阈值(可在 cfg 覆盖, 默认=V53 值, 保证默认配置精确复现)
    RT=cfg.get('rev_t'); RM=cfg.get('rev_m',(0.08,0.15))
    ST=cfg.get('slope_t'); SM=cfg.get('slope_m',(0.05,0.10)); SDT=cfg.get('slope_dnt',0.001); SDM=cfg.get('slope_dnm',0.05)
    CT=cfg.get('consec_t'); CM=cfg.get('consec_m',(0.02,0.05,0.10))
    if end_idx is None: end_idx=n
    portfolio={}; cash=INIT_CASH; trades_log=[]; prev_state='flat'
    nav_arr=[]
    peak_nav=INIT_CASH; max_dd=0; daily_nav=[]; yearly={}
    sell_reason_counter=defaultdict(int); ytb_skip_count=0
    total_commission=0.0; total_stamp_tax=0.0
    ind_share_sum=0.0; ind_share_cnt=0  # 集中度: 每日最大单一行业占持仓比
    for date_idx in range(start_idx, end_idx):
        current_date=dates_valid[date_idx]; yr=current_date[:4]
        prev_date=dates_valid[date_idx-1] if date_idx>=1 else None
        idx_state=v4_state.get(current_date,'flat')
        if cfg.get('v6_state'):
            idx_state=_v6_state.get(current_date,'short')
        if cfg.get('v5_state'):
            idx_state='bull' if _v5sig.get(current_date,0) else 'short'
        if idx_state=='flat' and prev_state=='short': idx_state='short'
        prev_state=idx_state
        max_pos=50 if idx_state=='bull' else (cfg.get('flat_pos',20) if idx_state=='flat' else 5)
        # v5风险开关(只空不多): v5=0时极致缩仓, v5=1时V4正常工作
        if cfg.get('v5_short_riskoff') and not _v5sig.get(current_date,1):
            max_pos=cfg.get('v5_short_max',1)
            idx_state='short'
        if yr not in yearly:
            nav=cash
            for c,p in portfolio.items():
                i=IND[c]['date_to_idx'].get(current_date)
                if i is not None and IND[c]['opens'][i]>0: nav+=IND[c]['opens'][i]*p['shares']
                elif i is not None and IND[c]['prices'][i]>0: nav+=IND[c]['prices'][i]*p['shares']
            yearly[yr]={'start':nav,'end':None,'buy':0,'sell':0,'force_sell':0,'ytb_skip':0}
        # 强卖
        if len(portfolio)>max_pos:
            excess=len(portfolio)-max_pos
            for code,pos in sorted(portfolio.items(),key=lambda x:x[1]['score'])[:excess]:
                i=IND[code]['date_to_idx'].get(current_date)
                if i is None: continue
                o=IND[code]['opens'][i]
                if o<=0 or math.isnan(o): o=IND[code]['prices'][i]
                if o<=0: continue
                cost_tax=o*pos['shares']*(COMMISSION_RATE+STAMP_TAX_RATE)
                total_commission+=o*pos['shares']*COMMISSION_RATE; total_stamp_tax+=o*pos['shares']*STAMP_TAX_RATE
                cash+=o*pos['shares']-cost_tax
                trades_log.append({'date':current_date,'code':code,'dir':'卖出','price':o,'shares':pos['shares'],'pnl':(o-pos['entry_price'])*pos['shares']-cost_tax,'reason':'超仓位','yr':yr,'state':idx_state})
                if code in portfolio: del portfolio[code]
                yearly[yr]['force_sell']+=1; sell_reason_counter['超仓位']+=1
        # 卖出 (与引擎一致)
        sells=[]
        for code,pos in list(portfolio.items()):
            if prev_date is None: continue
            i=IND[code]['date_to_idx'].get(prev_date)
            if i is None: continue
            dif_i=IND[code]['dif'][i]; dea_i=IND[code]['dea'][i]; ma5_i=IND[code]['ma5'][i]; c_i=IND[code]['prices'][i]
            if math.isnan(dif_i) or math.isnan(dea_i) or math.isnan(c_i) or c_i<=0: continue
            j=IND[code]['date_to_idx'].get(current_date)
            if j is None: continue
            o_j=IND[code]['opens'][j]
            if o_j<=0 or math.isnan(o_j): o_j=c_i
            c_j=IND[code]['prices'][j]
            if c_j<=0 or math.isnan(c_j): c_j=o_j
            pos['hold_days']+=1
            if c_i>pos['max_price']: pos['max_price']=c_i
            if c_i>pos['max_price']: pos['max_price']=c_i
            pnl_pct=(o_j-pos['entry_price'])/pos['entry_price']
            _eff_hold=_v4_hold.get(idx_state,_v4_hold['flat'])
            _eff_stop=_v4_stop.get(idx_state,_v4_stop['flat'])
            _eff_profit=_v4_profit.get(idx_state,_v4_profit['flat'])
            if pos['hold_days']>=_eff_hold and pnl_pct<=_eff_stop:
                sells.append((code,pos,o_j,'止损')); continue
            if pos['hold_days']>=_eff_hold and i>=1:
                dp=IND[code]['dif'][i-1]; ep=IND[code]['dea'][i-1]
                if not(math.isnan(dp) or math.isnan(ep)):
                    if dp>ep and dif_i<=dea_i: sells.append((code,pos,o_j,'死叉')); continue
            if pnl_pct>_eff_profit and not math.isnan(ma5_i) and c_i<ma5_i:
                sells.append((code,pos,o_j,'动态止盈')); continue
        seen=set(); su=[]
        for s in sells:
            if s[0] not in seen: seen.add(s[0]); su.append(s)
        sells=su
        for code,pos,price,reason in sorted(sells,key=lambda x:x[1]['score']):
            cost_tax=price*pos['shares']*(COMMISSION_RATE+STAMP_TAX_RATE)
            total_commission+=price*pos['shares']*COMMISSION_RATE; total_stamp_tax+=price*pos['shares']*STAMP_TAX_RATE
            cash+=price*pos['shares']-cost_tax
            trades_log.append({'date':current_date,'code':code,'dir':'卖出','price':price,'shares':pos['shares'],'pnl':(price-pos['entry_price'])*pos['shares']-cost_tax,'reason':reason,'yr':yr,'state':idx_state})
            if code in portfolio: del portfolio[code]
            yearly[yr]['sell']+=1; sell_reason_counter[reason]+=1
            GLOBAL_DUMP.append(f"卖 {current_date} {code} @ {price:.2f} r={reason} pnl={(price-pos['entry_price'])*pos['shares']:.0f}")
        # 买入 (轮动)
        if len(portfolio)<max_pos and cash>50000 and prev_date is not None:
            nav=cash
            for c,p in portfolio.items():
                i=IND[c]['date_to_idx'].get(current_date)
                if i is not None and IND[c]['opens'][i]>0: nav+=IND[c]['opens'][i]*p['shares']
                elif i is not None and IND[c]['prices'][i]>0: nav+=IND[c]['prices'][i]*p['shares']
            total_nav=nav
            cands=cand_by_dateidx[date_idx]
            if cands:
                sig_di=date_pos[prev_date]
                sig_map=industry_signal_at(sig_di, signal) if mode!='none' and signal else {}
                sig_map2=industry_signal_at(sig_di, signal2) if mode!='none' and signal2 else {}
                # 成分股过滤
                if cfg.get('zz1000_only',False) and ZZ1000_CODES:
                    cands=[c for c in cands if c['code'] in ZZ1000_CODES]
                if cfg.get('zz500_only',False) and ZZ500_CODES:
                    cands=[c for c in cands if c['code'] in ZZ500_CODES]
                # 组合分数: 量比必选 + 各项 consec 加分按flag开关 + 倾斜
                def compose(cd):
                    mult=1.0+cd['vol_bonus']
                    if reson and cd['ind'] is not None and cd['ind'] in RESON_MAP.get(reson,{}):
                        mult+=reson_k*RESON_MAP[reson][cd['ind']][sig_di]
                    if use_rev:
                        if RT is not None:
                            d=cd['rev_raw']
                            if d>RT[1]: mult+=RM[1]
                            elif d>RT[0]: mult+=RM[0]
                        else:
                            mult+=cd['rev_bonus']
                    if use_slope:
                        if ST is not None:
                            s=cd['slope_raw']
                            if s>ST[1]: mult+=SM[1]
                            elif s>ST[0]: mult+=SM[0]
                            elif s<-SDT: mult+=-SDM
                        else:
                            mult+=cd['slope_bonus']
                    if use_consec_b:
                        if CT is not None:
                            c=cd['consec_raw']
                            if c>=CT[2]: mult+=-CM[2]
                            elif c>=CT[1]: mult+=-CM[1]
                            elif c>=CT[0]: mult+=-CM[0]
                        else:
                            mult+=cd['consec_bonus']
                    sc=cd['base_score']*PENALTY.get(cd['gc'],0.2)*mult
                    if mode=='tilt' and cd['ind'] is not None:
                        _zt=0.0
                        if signal:
                            _zt+=TILT_W*clip(sig_map.get(cd['ind'],0.0),-2,2)
                        if signal2:
                            _zt+=tilt2*clip(sig_map2.get(cd['ind'],0.0),-2,2)
                        sc*=(1.0+_zt)
                    return sc
                if mode=='slot':
                    # softmax 配额
                    keys=list(sig_map.keys()); vals=[sig_map[k] for k in keys]
                    mx=max(vals) if vals else 0
                    ex=[math.exp(BETA_W*(v-mx)) for v in vals]; ssum=sum(ex)
                    w={keys[t]:ex[t]/ssum for t in range(len(keys))}
                    caps={k:max(0,round(max_pos*w[k])) for k in keys}
                    # 校正使 sum==max_pos
                    diff=max_pos-sum(caps.values())
                    if diff!=0:
                        # 加到权重最大行业
                        big=max(keys,key=lambda k:caps[k]) if keys else None
                        if big is not None: caps[big]+=diff
                    by_ind=defaultdict(list)
                    for cd in cands:
                        if cd['ind'] is not None: by_ind[cd['ind']].append(cd)
                    for indx in by_ind: by_ind[indx].sort(key=lambda x:compose(x), reverse=True)
                    ordered=[]
                    # 按权重降序填各行业的名额
                    for indx in sorted(keys,key=lambda k:w[k],reverse=True):
                        cap=caps.get(indx,0)
                        lst=by_ind.get(indx,[])
                        ordered.extend(lst[:cap])
                    # 余量溢出(保持资金在场)
                    filled=set(id(cd) for cd in ordered)
                    spillover=sorted([cd for cd in cands if id(cd) not in filled],key=lambda x:compose(x),reverse=True)
                    ordered.extend(spillover)
                else:
                    # none / tilt 统一用 compose (compose 内已含倾斜乘子), 保证 consec 因子参与排序
                    ordered=sorted(cands,key=compose,reverse=True)
                for cd in ordered:
                    if len(portfolio)>=max_pos or cash<=50000: break
                    if cd['is_zt']:
                        ytb_skip_count+=1; yearly[yr]['ytb_skip']+=1; continue
                    if cd['code'] in portfolio: continue
                    o=cd['o']
                    shares=max(int(POS_SIZE*total_nav/o/100)*100,100)
                    cost=shares*o; total_cost=cost*(1+COMMISSION_RATE)
                    if total_cost>cash*0.95: continue
                    total_commission+=cost*COMMISSION_RATE; cash-=total_cost
                    portfolio[cd['code']]={'shares':shares,'entry_price':o,'hold_days':0,'score':compose(cd),'max_price':o}
                    trades_log.append({'date':current_date,'code':cd['code'],'dir':'买入','price':o,'shares':shares,'pnl':0,'reason':f"轮动{cfg.get('name','?')}",'yr':yr,'state':idx_state})
                    GLOBAL_DUMP.append(f"买 {current_date} {cd['code']} @ {o:.2f} sh={shares}")
                    yearly[yr]['buy']+=1
        nav=cash
        for c,p in portfolio.items():
            i=IND[c]['date_to_idx'].get(current_date)
            if i is not None and IND[c]['opens'][i]>0: nav+=IND[c]['opens'][i]*p['shares']
            elif i is not None and IND[c]['prices'][i]>0: nav+=IND[c]['prices'][i]*p['shares']
        yearly[yr]['end']=nav
        if nav>peak_nav: peak_nav=nav
        dd=(peak_nav-nav)/peak_nav*100
        if dd>max_dd: max_dd=dd
        daily_nav.append({'date':current_date,'nav':nav,'yr':yr})
        nav_arr.append(nav)
        # 集中度: 当日最大单一行业占持仓数比重
        if portfolio:
            cnt=defaultdict(int)
            for c in portfolio:
                ind=code2ind.get(c)
                if ind: cnt[ind]+=1
            tot=sum(cnt.values())
            if tot>0:
                ind_share_sum+=max(cnt.values())/tot; ind_share_cnt+=1
    final_nav=yearly[list(yearly.keys())[-1]]['end']
    total_return=(final_nav/INIT_CASH-1)*100
    ratio=total_return/max_dd if max_dd>0 else 0
    if os.environ.get('DEBUG','0')=='1':
        with open('/workspace/v46_rot_my_trades.csv','w',newline='') as ff:
            w=csv.DictWriter(ff, fieldnames=['date','code','dir','price','shares','pnl','reason','yr','state'])
            w.writeheader(); 
            for t in trades_log: w.writerow(t)
    return {'name':cfg.get('name','?'),'final_nav':final_nav,'total_return':total_return,'max_dd':max_dd,'ratio':ratio,'nav_arr':nav_arr,
            'buy':sum(yearly[y]['buy'] for y in yearly),'sell':sum(yearly[y]['sell'] for y in yearly),
            'force_sell':sum(yearly[y]['force_sell'] for y in yearly),'ytb_skip':ytb_skip_count,
            'ind_share':(ind_share_sum/ind_share_cnt*100 if ind_share_cnt>0 else 0),
            'yearly':{y:{'start':yearly[y]['start'],'end':yearly[y]['end'],'buy':yearly[y]['buy'],'sell':yearly[y]['sell'],'force_sell':yearly[y]['force_sell'],'return':(yearly[y]['end']/yearly[y]['start']-1)*100 if yearly[y]['start']>0 else 0} for y in yearly}}

# === 实验选择器 (V54): EXP=scan 阈值稳定性 / EXP=signal 信号源替换 / 默认=V53消融 ===
EXP=os.environ.get('EXP','ablation')
if EXP=='scan':
    # 任务1: consec全因子 因子阈值稳定性扫描(全样本比率敏感性)
    CONFIGS=[
        {'name':'consec全因子(default)','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True},
        # rev 阈值(回撤%)扫描, 其余默认
        {'name':'rev 5/10%','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'rev_t':(0.05,0.10),'rev_m':(0.05,0.10)},
        {'name':'rev 8/15%(default)','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True},
        {'name':'rev 10/20%','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'rev_t':(0.10,0.20),'rev_m':(0.10,0.20)},
        {'name':'rev 12/25%','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'rev_t':(0.12,0.25),'rev_m':(0.12,0.25)},
        # slope 阈值(MA144斜率)扫描
        {'name':'slope 0.05/0.3%','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'slope_t':(0.0005,0.003),'slope_m':(0.05,0.10),'slope_dnt':0.0005,'slope_dnm':0.05},
        {'name':'slope 0.1/0.5%(default)','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True},
        {'name':'slope 0.2/1.0%','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'slope_t':(0.002,0.01),'slope_m':(0.05,0.10),'slope_dnt':0.002,'slope_dnm':0.05},
        {'name':'slope 0.3/1.5%','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'slope_t':(0.003,0.015),'slope_m':(0.05,0.10),'slope_dnt':0.003,'slope_dnm':0.05},
        # consec 连涨天数扫描
        {'name':'consec 2/4/6','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'consec_t':(2,4,6),'consec_m':(0.02,0.05,0.10)},
        {'name':'consec 3/5/7(default)','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True},
        {'name':'consec 4/6/8','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'consec_t':(4,6,8),'consec_m':(0.02,0.05,0.10)},
        {'name':'consec 5/8/10','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'consec_t':(5,8,10),'consec_m':(0.02,0.05,0.10)},
        # 联合扰动(验证联合敏感性)
        {'name':'all-loose','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'rev_t':(0.05,0.10),'rev_m':(0.05,0.10),'slope_t':(0.0005,0.003),'slope_m':(0.05,0.10),'slope_dnt':0.0005,'slope_dnm':0.05,'consec_t':(2,4,6),'consec_m':(0.02,0.05,0.10)},
        {'name':'all-tight','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'rev_t':(0.12,0.25),'rev_m':(0.12,0.25),'slope_t':(0.003,0.015),'slope_m':(0.05,0.10),'slope_dnt':0.003,'slope_dnm':0.05,'consec_t':(5,8,10),'consec_m':(0.02,0.05,0.10)},
    ]
elif EXP=='signal':
    # 任务2: 行业倾斜信号源替换(avgdif vs 行业指数MACD vs 前瞻行业收益) + k网格找峰值
    CONFIGS=[
        {'name':'倾斜-avgdif-k0.20','mode':'tilt','signal':'avgdif','tilt':0.20},
        {'name':'倾斜-avgdif-k0.26','mode':'tilt','signal':'avgdif','tilt':0.26},
        {'name':'倾斜-avgdif-k0.30','mode':'tilt','signal':'avgdif','tilt':0.30},
        {'name':'倾斜-avgdif-k0.40','mode':'tilt','signal':'avgdif','tilt':0.40},
        {'name':'倾斜-indmacd-k0.20','mode':'tilt','signal':'indmacd','tilt':0.20},
        {'name':'倾斜-indmacd-k0.26','mode':'tilt','signal':'indmacd','tilt':0.26},
        {'name':'倾斜-indmacd-k0.30','mode':'tilt','signal':'indmacd','tilt':0.30},
        {'name':'倾斜-indmacd-k0.40','mode':'tilt','signal':'indmacd','tilt':0.40},
        {'name':'倾斜-fwdret-k0.20(前瞻,非交易)','mode':'tilt','signal':'fwdret','tilt':0.20},
        {'name':'倾斜-fwdret-k0.26(前瞻,非交易)','mode':'tilt','signal':'fwdret','tilt':0.26},
        {'name':'倾斜-fwdret-k0.30(前瞻,非交易)','mode':'tilt','signal':'fwdret','tilt':0.30},
        {'name':'倾斜-fwdret-k0.40(前瞻,非交易)','mode':'tilt','signal':'fwdret','tilt':0.40},
    ]
elif EXP=='signalv2':
    # V55: 资金强度(amt=行业成交额z, amtr=成交额5日变化) + 赚钱效应(breadth=上涨占比, breadth_d=Δ占比) 各扫 k 找峰值
    _KS=[0.20,0.26,0.30,0.40]
    CONFIGS=[]
    for _sig in ['amt','amtr','breadth','breadth_d']:
        for _k in _KS:
            CONFIGS.append({'name':f'倾斜-{_sig}-k{_k}','mode':'tilt','signal':_sig,'tilt':_k})
elif EXP=='signalv3':
    # V56: 异常换手(abn=成交额/自身MA20) + 成交集中度(conc=HHI归一化) + 组合(abn_conc) 各扫 k 找峰值
    _KS=[0.20,0.26,0.30,0.40]
    CONFIGS=[]
    for _sig in ['abn','conc','abn_conc']:
        for _k in _KS:
            CONFIGS.append({'name':f'倾斜-{_sig}-k{_k}','mode':'tilt','signal':_sig,'tilt':_k})
elif EXP=='signalv4':
    # V57: conc滚动版(conc10/20) + avgdif+conc组合 各扫 k 找峰值
    CONFIGS=[]
    for _sig in ['conc10','conc20','avgdif_conc10','avgdif_conc20']:
        for _k in [0.15,0.20,0.26,0.30]:
            CONFIGS.append({'name':f'倾斜-{_sig}-k{_k}','mode':'tilt','signal':_sig,'tilt':_k})
elif EXP=='dual':
    # V58: 双倾斜因子 avgdif(k1) + conc20(k2), 加法叠加
    CONFIGS=[]
    for _k1 in [0.20,0.26]:
        for _k2 in [0.10,0.15,0.20]:
            CONFIGS.append({'name':f'双-avgdif{_k1}-conc20{_k2}','mode':'tilt','signal':'avgdif','tilt':_k1,'signal2':'conc20','tilt2':_k2,'consec':False})
    # 对照: 单因子
    CONFIGS.append({'name':'单-avgdif-k0.26','mode':'tilt','signal':'avgdif','tilt':0.26,'consec':False})
    CONFIGS.append({'name':'单-conc20-k0.15','mode':'tilt','signal':'conc20','tilt':0.15,'consec':False})
elif EXP=='v60':
    CONFIGS=[
        {'name':'V4基线','mode':'none'},
        {'name':'V4+consec','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True},
        {'name':'V6基线(flat30)','mode':'none','v6_state':True,'flat_pos':30},
        {'name':'V6+consec(flat30)','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'v6_state':True,'flat_pos':30},
    ]
else:
    # 默认: V53 消融(11配置)
    CONFIGS=[
        {'name':'量比(基线)','mode':'none'},
        {'name':'量比+倾斜(k0.26)','mode':'tilt','signal':'avgdif','tilt':0.26},
        {'name':'+slope','mode':'none','use_slope':True},
        {'name':'+slope+倾斜','mode':'tilt','signal':'avgdif','tilt':0.26,'use_slope':True},
        {'name':'+rev(超跌)','mode':'none','use_rev':True},
        {'name':'+rev+倾斜','mode':'tilt','signal':'avgdif','tilt':0.26,'use_rev':True},
        {'name':'+consec减分','mode':'none','use_consec_b':True},
        {'name':'+consec减分+倾斜','mode':'tilt','signal':'avgdif','tilt':0.26,'use_consec_b':True},
        {'name':'consec全因子','mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True},
        {'name':'consec全因子+倾斜','mode':'tilt','signal':'avgdif','tilt':0.26,'use_rev':True,'use_slope':True,'use_consec_b':True},
        {'name':'★rev+consec减分+倾斜','mode':'tilt','signal':'avgdif','tilt':0.26,'use_rev':True,'use_consec_b':True},
    ]
TILT_W=0.0; BETA_W=0.0  # 占位, 在 simulate 内按 cfg 读取
import types
def _simulate_wrap(cfg):
    global TILT_W, BETA_W
    TILT_W=cfg.get('tilt',0.0); BETA_W=cfg.get('beta',0.0)
    return simulate(cfg)

print(f"\n开始实验 [EXP={EXP}] ...", flush=True)
if EXP=='signal':
    print("  ⚠️ 倾斜-fwdret 为前瞻行业收益(含未来1日), 属 look-ahead 上界诊断, 非可交易信号!", flush=True)
elif EXP=='signalv2':
    print("  倾斜信号源 V55: 资金强度(amt=行业成交额z / amtr=成交额5日变化) + 赚钱效应(breadth=上涨占比 / breadth_d=Δ占比) — 均 OHLCV 派生, 与趋势同域; 重点看正交性诊断与 OOS。", flush=True)
elif EXP=='signalv3':
    print("  倾斜信号源 V56: 异常换手(abn=成交额/自身MA20, 去规模偏) + 成交集中度(conc=HHI归一化, 资金聚焦龙头) + 组合(abn_conc) — 修复 V55 总量规模偏/无集中度缺陷; 重点看正交性诊断与 OOS。", flush=True)
elif EXP=='signalv4':
    print("  V57: conc滚动版(conc10/conc20=10/20日HHI均值, 抓阶段性热度) + avgdif+conc组合(正交信号相加) — 用L1行业, 看平滑版+组合能否超越纯avgdif。", flush=True)
elif EXP=='dual':
    print("  V58: 双倾斜因子 avgdif(k1) + conc20(k2) 加法叠加 — score*=1+k1*z_avgdif+k2*z_conc20; 扫 k1∈{0.20,0.26}, k2∈{0.10,0.15,0.20} + 单因子对照。", flush=True)
elif EXP=='analysis':
    print("  备选池 Top20 多期持有分析 — 每日按量比基线取top20, 计算3/5/10/20日 open-to-open 胜率与收益。", flush=True)
    _rb=simulate({'mode':'none','consec':False,'name':'base'},0,n)
    _nav=_rb['nav_arr']
    import pandas as pd
    _zz=pd.read_parquet('/workspace/data/zz1000_hfq.parquet')
    _H=[3,5,10,20]; _maxH=max(_H)
    _wr={h:[] for h in _H}; _ar={h:[] for h in _H}; _ir={h:[] for h in _H}
    _dt=[]; _nav_on=[]
    for _di in range(1,n-_maxH):
        _cs=cand_by_dateidx[_di]
        if not _cs: continue
        _sc=[cd['base_score']*0.2*(1+cd['vol_bonus']) for cd in _cs]
        _top=sorted(range(len(_cs)), key=lambda i:_sc[i], reverse=True)[:20]
        _rd={h:[] for h in _H}
        for _i in _top:
            _cd=_cs[_i]; _ei=IND[_cd['code']]['date_to_idx'].get(dates_valid[_di])
            for h in _H:
                _xi=IND[_cd['code']]['date_to_idx'].get(dates_valid[_di+h])
                if _ei is not None and _xi is not None and IND[_cd['code']]['opens'][_ei]>0 and IND[_cd['code']]['opens'][_xi]>0:
                    _rd[h].append(IND[_cd['code']]['opens'][_xi]/IND[_cd['code']]['opens'][_ei]-1)
        if all(len(_rd[h])>=10 for h in _H):
            for h in _H:
                _wr[h].append(sum(1 for r in _rd[h] if r>0)/len(_rd[h]))
                _ar[h].append(sum(_rd[h])/len(_rd[h]))
            _dt.append(dates_valid[_di])
            _ni=_zz.index.get_loc(dates_valid[_di]) if dates_valid[_di] in _zz.index else -1
            if _ni>=0 and _ni+_maxH<len(_zz):
                _nav_on.append(_nav[_di])
                for h in _H:
                    _ir[h].append(_zz.iloc[_ni+h]['close']/_zz.iloc[_ni]['close']-1)
    _N=len(_dt)
    print(f"\n有效天数: {_N}")
    print(f"\n{'持有期':>6} {'平均胜率':>8} {'中位胜率':>8} {'平均收益':>9} {'>50%':>6} {'>60%':>6} {'<40%':>6} {' corr指数':>8} {' corrNAV':>8}")
    print("-"*80)
    _wn=sum(_nav_on)/_N
    for h in _H:
        _wr_h=_wr[h]; _ar_h=_ar[h]; _ir_h=_ir[h]
        _n=len(_wr_h); _wm=sum(_wr_h)/_n; _am=sum(_ar_h)/_n; _im=sum(_ir_h)/_n
        _gt50=sum(1 for w in _wr_h if w>0.5)/_n*100
        _gt60=sum(1 for w in _wr_h if w>0.6)/_n*100
        _lt40=sum(1 for w in _wr_h if w<0.4)/_n*100
        _ci=sum((_wr_h[i]-_wm)*(_ir_h[i]-_im) for i in range(_n))
        _cn=sum((_wr_h[i]-_wm)*(_nav_on[i]-_wn) for i in range(_n))
        _s1=(sum((x-_wm)**2 for x in _wr_h)/_n)**0.5
        _s2=(sum((x-_im)**2 for x in _ir_h)/_n)**0.5
        _s3=(sum((x-_wn)**2 for x in _nav_on)/_n)**0.5
        _corr_i=f'{_ci/_n/_s1/_s2:.3f}' if _s1>0 and _s2>0 else 'NA'
        _corr_n=f'{_cn/_n/_s1/_s3:.3f}' if _s1>0 and _s3>0 else 'NA'
        _med=sorted(_wr_h)[_n//2]*100
        print(f"{h:>4}日 {_wm*100:>7.1f}% {_med:>7.1f}% {_am*100:>8.2f}% {_gt50:>5.1f}% {_gt60:>5.1f}% {_lt40:>5.1f}% {_corr_i:>8} {_corr_n:>8}")
    print()
    for _s,_e,_lb in [(0,_N//3,'初期'),(_N//3,_N*2//3,'中期'),(_N*2//3,_N,'近期')]:
        if _e-_s>0:
            _ss=f"  {_lb}({_dt[_s][:7]}~{_dt[_e-1][:7]}):"
            for h in _H: _ss+=f" {h}日{sum(_wr[h][_s:_e])/(_e-_s)*100:.1f}%"
            print(_ss)
    print(f"\n指数涨跌时 胜率>50% 占比:")
    for h in _H:
        _wr_h=_wr[h]; _ir_h=_ir[h]
        _up=sum(1 for i in range(_N) if _ir_h[i]>0 and _wr_h[i]>0.5)
        _up_t=sum(1 for i in range(_N) if _ir_h[i]>0)
        _dn=sum(1 for i in range(_N) if _ir_h[i]<=0 and _wr_h[i]>0.5)
        _dn_t=sum(1 for i in range(_N) if _ir_h[i]<=0)
        print(f"  {h}日: 指数涨时{_up/_up_t*100:.0f}% / 指数跌时{_dn/_dn_t*100:.0f}% (n_up={_up_t}, n_dn={_dn_t})")
    import sys; sys.exit(0)
elif EXP=='v60':
    print("  V60: V6状态机(50/30/5) vs V4(50/20/5) — 基线+consec对比", flush=True)
t0=time.time()
rows=[]
DEBUG = os.environ.get('DEBUG','0')=='1'
GRID = CONFIGS if not DEBUG else [c for c in CONFIGS if '基线' in c['name'] or '量比+倾斜' in c['name']]
for cfg in GRID:
        r=_simulate_wrap(cfg)
        rows.append(r)
        print(f"  {r['name']:<18} 收益={r['total_return']:>8.2f}%  maxdd={r['max_dd']:>6.2f}%  ratio={r['ratio']:>6.2f}  行业集中度={r['ind_share']:>5.1f}%  买={r['buy']}", flush=True)
        if DEBUG:
            tot_cand=sum(len(c) for c in cand_by_dateidx if c)
            print(f"    [DEBUG] 总候选数={tot_cand}  有候选天数={sum(1 for c in cand_by_dateidx if c)}")
            for y,info in sorted(r['yearly'].items()):
                print(f"      {y} start={info['start']:,.0f} end={info['end']:,.0f} ret={info['return']:.2f}% buy={info['buy']} sell={info['sell']} force={info['force_sell']}")
            print("    [DEBUG] 前200笔交易(含买卖):")
            for t in GLOBAL_DUMP[:200]:
                print(f"      {t}")
print(f"网格耗时: {time.time()-t0:.1f}秒", flush=True)

print("\n"+"="*92)
print(f"{'配置':<18}{'收益%':>10}{'maxdd%':>10}{'ratio':>8}{'行业集中度%':>12}{'买入':>8}{'卖出':>8}")
print("-"*92)
for r in rows:
    print(f"{r['name']:<18}{r['total_return']:>10.2f}{r['max_dd']:>10.2f}{r['ratio']:>8.2f}{r['ind_share']:>12.1f}{r['buy']:>8}{r['sell']:>8}")
print("="*92)
print(f"对照: 量比基线 557.79/19.43/28.71 ; consec全因子 596.94/17.63/33.85")

# === Walk-forward 稳定性检验 (扩展窗口, NAV序列切片, 持仓连续) ===
if os.environ.get('WALK','0')=='1':
    K_GRID=[0.0,0.15,0.20,0.25,0.26,0.30]
    anchors=[0,288,576,864,1152,n]
    test_windows=[(anchors[i], anchors[i+1]) for i in range(1,len(anchors)-1)]
    def win_metrics(nav_arr, a, b):
        n0=nav_arr[a]; ret=nav_arr[b-1]/n0-1
        peak=n0; mdd=0.0
        for idx in range(a,b):
            if nav_arr[idx]>peak: peak=nav_arr[idx]
            dd=(peak-nav_arr[idx])/peak if peak>0 else 0
            if dd>mdd: mdd=dd
        ret_p=ret*100; mdd_p=mdd*100
        return ret_p, mdd_p, (ret_p/mdd_p if mdd_p>0 else 0.0)
    print("\n"+"="*104)
    print("Walk-forward 稳定性检验 (扩展窗口: 训练[0,s)选k*, 测试[s,e)做OOS, 持仓连续切片)")
    print("="*104)
    print(f"{'测试窗':<24}{'训练选k*':>9}{'OOS基线':>10}{'OOS选k*':>10}{'OOS k0.26':>11}{'k*胜基线':>10}{'0.26胜基线':>11}")
    print("-"*104)
    agg={'base':[],'sel':[],'026':[]}
    n_sel_win=0; n_026_win=0
    for (s,e) in test_windows:
        # 训练窗[0,s): 对每个k跑整段[0,e)并切片[0,s)选k*
        best_k=0.0; best_ratio=-1e9
        for k in K_GRID:
            r=simulate({'mode':'tilt','signal':'avgdif','tilt':k,'consec':False,'name':f'k{k}'}, 0, e)
            tr,_tm,tratio=win_metrics(r['nav_arr'], 0, s)
            if tratio>best_ratio: best_ratio=tratio; best_k=k
        rb=simulate({'mode':'none','consec':False,'name':'base'}, 0, e)
        rk=simulate({'mode':'tilt','signal':'avgdif','tilt':best_k,'consec':False,'name':f'k{best_k}'}, 0, e)
        r26=simulate({'mode':'tilt','signal':'avgdif','tilt':0.26,'consec':False,'name':'k0.26'}, 0, e)
        _br,bm,bratio=win_metrics(rb['nav_arr'], s, e)
        _kr,km,kratio=win_metrics(rk['nav_arr'], s, e)
        _rr,rm,r26ratio=win_metrics(r26['nav_arr'], s, e)
        sel_win = 1 if kratio>bratio else 0
        w26 = 1 if r26ratio>bratio else 0
        n_sel_win+=sel_win; n_026_win+=w26
        agg['base'].append(bratio); agg['sel'].append(kratio); agg['026'].append(r26ratio)
        print(f"{(dates_valid[s]+'~'+dates_valid[e-1]).ljust(22)}",
              f"{best_k:>9.2f}", f"{bratio:>10.2f}", f"{kratio:>10.2f}", f"{r26ratio:>11.2f}",
              f"{'是' if sel_win else '否':>10}", f"{'是' if w26 else '否':>11}")
    mb=sum(agg['base'])/len(agg['base']); ms=sum(agg['sel'])/len(agg['sel']); m26=sum(agg['026'])/len(agg['026'])
    print("-"*104)
    print(f"{'OOS均值':<24}{'':>9}{mb:>10.2f}{ms:>10.2f}{m26:>11.2f}   {n_sel_win}/{len(test_windows)}      {n_026_win}/{len(test_windows)}")
    print(f"\n结论: 训练窗选k*在 {n_sel_win}/{len(test_windows)} 个测试窗 OOS胜基线; 固定k=0.26 在 {n_026_win}/{len(test_windows)} 个测试窗 OOS胜基线")
    print(f"      OOS平均比率: 基线={mb:.2f} | 选k*={ms:.2f} | k0.26={m26:.2f}")
    # 全样本实际对照 (修正: 不再引用陈值 36.77)
    _rb=simulate({'mode':'none','consec':False,'name':'base'},0,n)
    _rk=simulate({'mode':'tilt','signal':'avgdif','tilt':0.26,'consec':False,'name':'k0.26'},0,n)
    print(f"      全样本对照: 基线 ratio={_rb['ratio']:.2f}(收益{_rb['total_return']:.2f}%) | 量比+倾斜 ratio={_rk['ratio']:.2f}(收益{_rk['total_return']:.2f}%) 差 {_rk['ratio']-_rb['ratio']:+.2f}")

    # === consec全因子 walk-forward 稳健性检验 (与 k=0.26 同口径) ===
    print("\n"+"="*104)
    print("consec全因子 Walk-forward 稳健性检验 (扩展窗口[0,s)训练/[s,e)OOS, NAV切片, 持仓连续)")
    print("="*104)
    print(f"{'测试窗':<22}{'OOS基线':>9}{'OOS consec':>12}{'OOS k0.26':>12}{'consec胜基线':>14}")
    print("-"*104)
    c_base=[]; c_consec=[]; c_26=[]
    c_win=0
    for (s,e) in test_windows:
        rb=simulate({'mode':'none','consec':False,'name':'base'},0,e)
        rc=simulate({'mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'name':'consec'},0,e)
        r26=simulate({'mode':'tilt','signal':'avgdif','tilt':0.26,'consec':False,'name':'k0.26'},0,e)
        _br,bm,bratio=win_metrics(rb['nav_arr'],s,e)
        _cr,cm,cratio=win_metrics(rc['nav_arr'],s,e)
        _rr,rm,r26ratio=win_metrics(r26['nav_arr'],s,e)
        win=1 if cratio>bratio else 0
        c_win+=win
        c_base.append(bratio); c_consec.append(cratio); c_26.append(r26ratio)
        print(f"{(dates_valid[s]+'~'+dates_valid[e-1]).ljust(20)}",f"{bratio:>9.2f}",f"{cratio:>12.2f}",f"{r26ratio:>12.2f}",f"{'是' if win else '否':>14}")
    mb2=sum(c_base)/len(c_base); mc2=sum(c_consec)/len(c_consec); m26b=sum(c_26)/len(c_26)
    print("-"*104)
    print(f"{'OOS均值':<22}{mb2:>9.2f}{mc2:>12.2f}{m26b:>12.2f}   consec胜基线 {c_win}/{len(test_windows)}")
    _rf=simulate({'mode':'none','consec':False,'name':'base'},0,n)
    _cf=simulate({'mode':'none','use_rev':True,'use_slope':True,'use_consec_b':True,'name':'consec'},0,n)
    print(f"全样本对照: 基线 ratio={_rf['ratio']:.2f}({_rf['total_return']:.2f}%) | consec全因子 ratio={_cf['ratio']:.2f}({_cf['total_return']:.2f}%) 差 {_cf['ratio']-_rf['ratio']:+.2f}")
    print(f"          consec全因子 OOS均值 {mc2:.2f} vs 量比+倾斜OOS均值 {m26b:.2f} → consec {'更优' if mc2>m26b else '更劣'}")

    # === V55 新信号 walk-forward 稳健性 (最佳新信号 vs 基线 & avgdif k0.26) ===
    _v5_sigs=[('breadth',0.20),('amtr',0.20)]
    for _sig,_k in _v5_sigs:
        _col=f'OOS {_sig}'
        print("\n"+"="*104)
        print(f"V55 新行业信号 Walk-forward ({_sig} @k{_k} vs 基线 vs avgdif k0.26, 扩展窗口[0,s)/[s,e))")
        print("="*104)
        print(f"{'测试窗':<22}{'OOS基线':>9}{_col:>13}{'OOS k0.26':>12}{f'{_sig}胜基线':>15}")
        print("-"*104)
        _b=[]; _x=[]; _t=[]; _win=0
        for (s,e) in test_windows:
            rb=simulate({'mode':'none','consec':False,'name':'base'},0,e)
            rx=simulate({'mode':'tilt','signal':_sig,'tilt':_k,'consec':False,'name':_sig},0,e)
            r26=simulate({'mode':'tilt','signal':'avgdif','tilt':0.26,'consec':False,'name':'k0.26'},0,e)
            _br,bm,bratio=win_metrics(rb['nav_arr'],s,e)
            _xr,xm,xratio=win_metrics(rx['nav_arr'],s,e)
            _rr,rm,r26ratio=win_metrics(r26['nav_arr'],s,e)
            win=1 if xratio>bratio else 0
            _win+=win
            _b.append(bratio); _x.append(xratio); _t.append(r26ratio)
            print(f"{(dates_valid[s]+'~'+dates_valid[e-1]).ljust(20)}",f"{bratio:>9.2f}",f"{xratio:>13.2f}",f"{r26ratio:>12.2f}",f"{'是' if win else '否':>15}")
        _mb=sum(_b)/len(_b); _mx=sum(_x)/len(_x); _mt=sum(_t)/len(_t)
        print("-"*104)
        print(f"{'OOS均值':<22}{_mb:>9.2f}{_mx:>13.2f}{_mt:>12.2f}   {_sig}胜基线 {_win}/{len(test_windows)}")
        print(f"      {_sig}@k{_k} OOS均值 {_mx:.2f} vs avgdif k0.26 OOS均值 {_mt:.2f} → 新信号 {'更优' if _mx>_mt else '更劣'}; 均 vs 基线 {_mb:.2f}")

    # === V56 新信号 walk-forward 稳健性 (异常换手/集中度/组合 vs 基线 & avgdif k0.26) ===
    _v6_sigs=[('abn_conc',0.20),('abn',0.20),('conc',0.20)]
    for _sig,_k in _v6_sigs:
        _col=f'OOS {_sig}'
        print("\n"+"="*104)
        print(f"V56 新行业信号 Walk-forward ({_sig} @k{_k} vs 基线 vs avgdif k0.26, 扩展窗口[0,s)/[s,e))")
        print("="*104)
        print(f"{'测试窗':<22}{'OOS基线':>9}{_col:>13}{'OOS k0.26':>12}{f'{_sig}胜基线':>15}")
        print("-"*104)
        _b=[]; _x=[]; _t=[]; _win=0
        for (s,e) in test_windows:
            rb=simulate({'mode':'none','consec':False,'name':'base'},0,e)
            rx=simulate({'mode':'tilt','signal':_sig,'tilt':_k,'consec':False,'name':_sig},0,e)
            r26=simulate({'mode':'tilt','signal':'avgdif','tilt':0.26,'consec':False,'name':'k0.26'},0,e)
            _br,bm,bratio=win_metrics(rb['nav_arr'],s,e)
            _xr,xm,xratio=win_metrics(rx['nav_arr'],s,e)
            _rr,rm,r26ratio=win_metrics(r26['nav_arr'],s,e)
            win=1 if xratio>bratio else 0
            _win+=win
            _b.append(bratio); _x.append(xratio); _t.append(r26ratio)
            print(f"{(dates_valid[s]+'~'+dates_valid[e-1]).ljust(20)}",f"{bratio:>9.2f}",f"{xratio:>13.2f}",f"{r26ratio:>12.2f}",f"{'是' if win else '否':>15}")
        _mb=sum(_b)/len(_b); _mx=sum(_x)/len(_x); _mt=sum(_t)/len(_t)
        print("-"*104)
        print(f"{'OOS均值':<22}{_mb:>9.2f}{_mx:>13.2f}{_mt:>12.2f}   {_sig}胜基线 {_win}/{len(test_windows)}")
        print(f"      {_sig}@k{_k} OOS均值 {_mx:.2f} vs avgdif k0.26 OOS均值 {_mt:.2f} → 新信号 {'更优' if _mx>_mt else '更劣'}; 均 vs 基线 {_mb:.2f}")

    # === V57 conc平滑版 walk-forward (conc20 vs 基线 & avgdif k0.26) ===
    _v7_sigs=[('conc20',0.15),('conc10',0.26)]
    for _sig,_k in _v7_sigs:
        _col=f'OOS {_sig}'
        print("\n"+"="*104)
        print(f"V57 conc平滑版 Walk-forward ({_sig} @k{_k} vs 基线 vs avgdif k0.26, 扩展窗口[0,s)/[s,e))")
        print("="*104)
        print(f"{'测试窗':<22}{'OOS基线':>9}{_col:>13}{'OOS k0.26':>12}{f'{_sig}胜基线':>15}")
        print("-"*104)
        _b=[]; _x=[]; _t=[]; _win=0
        for (s,e) in test_windows:
            rb=simulate({'mode':'none','consec':False,'name':'base'},0,e)
            rx=simulate({'mode':'tilt','signal':_sig,'tilt':_k,'consec':False,'name':_sig},0,e)
            r26=simulate({'mode':'tilt','signal':'avgdif','tilt':0.26,'consec':False,'name':'k0.26'},0,e)
            _br,bm,bratio=win_metrics(rb['nav_arr'],s,e)
            _xr,xm,xratio=win_metrics(rx['nav_arr'],s,e)
            _rr,rm,r26ratio=win_metrics(r26['nav_arr'],s,e)
            win=1 if xratio>bratio else 0
            _win+=win
            _b.append(bratio); _x.append(xratio); _t.append(r26ratio)
            print(f"{(dates_valid[s]+'~'+dates_valid[e-1]).ljust(20)}",f"{bratio:>9.2f}",f"{xratio:>13.2f}",f"{r26ratio:>12.2f}",f"{'是' if win else '否':>15}")
        _mb=sum(_b)/len(_b); _mx=sum(_x)/len(_x); _mt=sum(_t)/len(_t)
        print("-"*104)
        print(f"{'OOS均值':<22}{_mb:>9.2f}{_mx:>13.2f}{_mt:>12.2f}   {_sig}胜基线 {_win}/{len(test_windows)}")
        print(f"      {_sig}@k{_k} OOS均值 {_mx:.2f} vs avgdif k0.26 OOS均值 {_mt:.2f} → 新信号 {'更优' if _mx>_mt else '更劣'}; 均 vs 基线 {_mb:.2f}")

_RES_JSON=f'/workspace/v46_rotation_results_{EXP}.json'
with open(_RES_JSON,'w') as f:
    json.dump(rows, f, ensure_ascii=False, indent=2, default=str)
print(f"\n已保存 {_RES_JSON}")
