# -*- coding: utf-8 -*-
"""
瑞士电子烟市场调研数据看板 v1
------------------------------------------------------------
结构沿用阿联酋看板 vape_ae_dashboard_v2.py，按瑞士数据集特征重构：
  · 数据源：sheet "popular sell" —— 7 个瑞士零售网站热销榜
      E-liquid 前20 / Disposable Vape 前20（nosmoky.ch 一次性仅 12 条）对应的口味明细
  · 与阿联酋版的主要差异：
      1) 瑞士数据集没有独立的 "Disposable Vape" 型号榜 sheet，一次性与烟油的口味数据结构相近，
         因此不再单设「一次性榜单」页：一次性的品牌 / 口数 / 价格分析并入「品牌分析」「产品规格」「价格分析」页，
         并统一受侧边栏筛选联动（阿联酋版该页使用独立数据表，不受筛选影响）；
      2) 币种为 CHF，价格为逗号小数格式（如 "7,50 CHF"）；
      3) 口数字段格式为 0.6k / 1k / 2k4 / 10k；
      4) 法规按瑞士 TabPG / TabPV 重写：尼古丁 ≤ 20 mg/mL、补充装 ≤ 10 mL、一次性 ≤ 2 mL；
         消费税：一次性 CHF 1.00/mL，可重复使用系统（含尼古丁）CHF 0.20/mL；
      5) 新增「容量超限线索」分析（对照 TabPG 第9条容量上限）；消费税在法规页仅展示税率，不做估算。
  · 法规 / 市场背景：瑞士电子烟市场调研概览.docx

运行方式：
    pip install streamlit pandas plotly openpyxl
    streamlit run vape_ch_dashboard_v1.py
将本脚本与 Excel 数据文件放在同一目录。
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from collections import Counter
import re
import os
import glob
import inspect

st.set_page_config(
    page_title="瑞士电子烟市场调研数据看板",
    page_icon="🇨🇭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 常量
# ============================================================
# 汇率：CHF→CNY 固定按 8.2 折算（人工设定，不随市场变动；如需调整直接改这里）
CHF_CNY = 8.2

# 与法规对照用阈值（来源：瑞士电子烟市场调研概览.docx，TabPG 第7、9条及附件2）
NIC_LIMIT_MG_ML = 20          # 含尼古丁液体 ≤ 20 mg/mL
REFILL_LIMIT_ML = 10          # 补充装容器 ≤ 10 mL
DISPOSABLE_LIMIT_ML = 2       # 一次性电子烟及一次性烟弹 ≤ 2 mL
TAX_DISPOSABLE_CHF_ML = 1.00  # 一次性电子烟：CHF 1.00/mL（无论是否含尼古丁）
TAX_REFILL_NIC_CHF_ML = 0.20  # 可重复使用系统：仅含尼古丁的液体/补充装/烟弹 CHF 0.20/mL
TOPN = 20                     # 各榜单最多取前 20 名，用于“榜单得分”
NOSMOKY_DP_EXPECTED = 10      # 调研概览说明：nosmoky.ch 一次性仅取前 10 名

st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #D52B1E 0%, #7f1d1d 100%);
        padding: 16px 20px; border-radius: 12px; color: white;
        text-align: center; box-shadow: 0 4px 15px rgba(213,43,30,0.25);
    }
    .metric-card .value { font-size: 2rem; font-weight: 700; line-height: 1; }
    .metric-card .label { font-size: 0.8rem; opacity: 0.9; margin-top: 6px; }
    .metric-card-alt {
        background: linear-gradient(135deg, #475569 0%, #1e293b 100%);
        padding: 16px 20px; border-radius: 12px; color: white;
        text-align: center; box-shadow: 0 4px 15px rgba(30,41,59,0.25);
    }
    .metric-card-alt .value { font-size: 2rem; font-weight: 700; line-height: 1; }
    .metric-card-alt .label { font-size: 0.8rem; opacity: 0.9; margin-top: 6px; }
    .section-title {
        font-size: 1.05rem; font-weight: 600; color: #7f1d1d;
        margin: 18px 0 10px 0; padding-left: 10px;
        border-left: 4px solid #D52B1E;
    }
    .insight-box {
        background: #fdf1f0; border-left: 4px solid #D52B1E;
        padding: 12px 16px; border-radius: 0 8px 8px 0;
        margin: 8px 0; font-size: 0.92rem; line-height: 1.7;
    }
    .warning-box {
        background: #fef3c7; border-left: 4px solid #d97706;
        padding: 12px 16px; border-radius: 0 8px 8px 0;
        margin: 8px 0; font-size: 0.92rem; line-height: 1.7;
    }
    .danger-box {
        background: #fef2f2; border-left: 4px solid #991b1b;
        padding: 12px 16px; border-radius: 0 8px 8px 0;
        margin: 8px 0; font-size: 0.92rem; line-height: 1.7;
    }
    .action-box {
        background: #f0fdf4; border-left: 4px solid #10b981;
        padding: 12px 16px; border-radius: 0 8px 8px 0;
        margin: 8px 0; font-size: 0.92rem; line-height: 1.7;
    }
    .tag-chip {
        display: inline-block; background: #f8d7d4; color: #7f1d1d;
        padding: 2px 10px; border-radius: 999px;
        font-size: 0.78rem; margin: 2px;
    }
    .flavor-row {
        margin: 3px 0 6px 0; padding: 6px 10px;
        background: #f8fafc; border-radius: 6px;
        font-size: 0.85rem; color: #374151;
    }
    .summary-card {
        background: #fdf1f0; border-radius: 12px; padding: 16px;
        margin: 12px 0; border-left: 6px solid #D52B1E;
    }
    .warning-card {
        background: #fff7ed; border-radius: 12px; padding: 16px;
        margin: 12px 0; border-left: 6px solid #f97316;
    }
    .reg-card {
        background: #f8fafc; border-radius: 10px; padding: 12px 16px;
        margin: 6px 0; border-left: 4px solid #94a3b8;
        font-size: 0.88rem; line-height: 1.6;
    }
    .reg-card-ban { border-left: 4px solid #991b1b; }
    .reg-card-restrict { border-left: 4px solid #d97706; }
    .reg-card-open { border-left: 4px solid #10b981; }
</style>
""", unsafe_allow_html=True)

COLORS = {
    "primary":   "#D52B1E",   # 瑞士国旗红
    "secondary": "#334155",   # 石板灰，用作第二主色
    "gold":      "#b45309",
    "ice":       "#38bdf8",
    "tobacco":   "#92400e",
    "fruit":     "#22c55e",
    "drink":     "#06b6d4",
    "candy":     "#8b5cf6",
    "menthol":   "#34d399",
    "dessert":   "#f97316",
    "other":     "#94a3b8",
}
RED_SCALE = [[0, "#fbdcd9"], [1, COLORS["primary"]]]
SLATE_SCALE = [[0, "#e2e8f0"], [1, COLORS["secondary"]]]
HEAT_SCALE = [[0, "#fdf1f0"], [1, COLORS["primary"]]]

CAT_COLOR_MAP = {
    "水果": COLORS["fruit"], "烟草": COLORS["tobacco"], "甜点": COLORS["dessert"],
    "饮料": COLORS["drink"], "糖果": COLORS["candy"], "薄荷": COLORS["menthol"],
    "其他": COLORS["other"],
}
COUNTRY_COLOR_MAP = {
    "瑞士":   COLORS["primary"],
    "中国":   "#f59e0b",
    "英国":   "#8b5cf6",
    "法国":   "#3b82f6",
    "意大利": "#22c55e",
    "美国":   "#0ea5e9",
    "未知":   COLORS["other"],
}
FORM_EL = "瓶装烟油 E-liquid"
FORM_DP = "一次性电子烟 Disposable"
FORM_COLOR_MAP = {FORM_EL: COLORS["secondary"], FORM_DP: COLORS["primary"]}
ICE_TAGS = {"ice", "menthol", "koolada"}   # 带这些标签即视为含冰/凉感
T = "plotly_white"

# ============================================================
# 数据文件
# ============================================================
_DIR = os.path.dirname(os.path.abspath(__file__))
# ⚠️ 如需更换数据文件，修改下面的文件名即可（与本脚本放同一目录）；找不到时会自动匹配 瑞士*.xlsx
EXCEL_FILE = os.path.join(_DIR, "瑞士热销产品榜单-20261019.xlsx")
DATA_DATE = "2026-10-19"          # 来自数据文件名中的日期
SHEET_POP = "popular sell"


def resolve_excel():
    if os.path.exists(EXCEL_FILE):
        return EXCEL_FILE
    cands = sorted(glob.glob(os.path.join(_DIR, "瑞士*.xlsx")))
    return cands[-1] if cands else None


# ---- 归一化规则（仅处理大小写 / 拼写重复，不做主观合并）----
BRAND_NORM_MAP = {
    "elf bar": "ElfBar", "elfbar": "ElfBar",
    "ske": "SKE", "vozol": "Vozol", "lost mary": "Lost Mary",
    "kiwi vapor": "KIWI VAPOR", "happy hale": "Happy Hale",
    "aroma king": "Aroma King", "ultd salts": "ULTD Salts",
}
# 口味标签里的拼写 / 写法差异（原表中确有的笔误与空格差异）
TAG_NORM_MAP = {
    "passionfruit": "passion fruit",
    "bule raspberry": "blue raspberry",
    "buleberry": "blueberry",
    "eucalyptu": "eucalyptus",
    "dragon fruit": "dragonfruit",
}
DEVICE_NORM_MAP = {"瓶装烟油": FORM_EL, "一次性电子烟": FORM_DP}


# ============================================================
# 通用函数
# ============================================================
_CHART_SEQ = {"n": 0}   # 每次脚本重跑都会重置；按出现顺序给每个图表分配稳定 key，
                         # 避免筛选到极少量数据时两张内容相同的图触发 DuplicateElementId


def _next_key(prefix):
    _CHART_SEQ["n"] += 1
    return f"{prefix}_{_CHART_SEQ['n']}"


def _params(fn):
    try:
        return inspect.signature(fn).parameters
    except (TypeError, ValueError):
        return {}


def _layout_kwargs(fn):
    """按当前 Streamlit 版本选择“撑满容器宽度”的写法：
    新版（width 参数默认 "stretch"）用 width="stretch"；旧版用 use_container_width=True。"""
    pr = _params(fn)
    w = pr.get("width")
    if w is not None and w.default == "stretch":
        return {"width": "stretch"}
    if "use_container_width" in pr:
        return {"use_container_width": True}
    return {}


def show_chart(fig):
    """兼容不同 Streamlit 版本的 plotly 渲染。"""
    kw = _layout_kwargs(st.plotly_chart)
    if "key" in _params(st.plotly_chart):
        kw["key"] = _next_key("chart")
    st.plotly_chart(fig, **kw)


def show_df(df, **kwargs):
    kw = _layout_kwargs(st.dataframe)
    if "key" in _params(st.dataframe):
        kw["key"] = _next_key("df")
    st.dataframe(df, hide_index=True, **kw, **kwargs)


def top_n_with_ties(counts, n=3):
    """counts 为降序的 value_counts Series。取前 n 名；若第 n 名出现并列，把并列项一并纳入（并列值 ≥ 2 次时）。"""
    if len(counts) <= n:
        return counts
    cutoff = counts.iloc[n - 1]
    if cutoff <= 1:           # 并列值只有 1 次时不展开（大量“仅出现一次”的项没有排名意义，展开会让图表失控）
        return counts.head(n)
    return counts[counts >= cutoff]


def tie_cap(shown_len, n):
    """Top-N 图表在第 N 名出现并列、实际显示条数 > N 时，给一行说明。"""
    if shown_len > n:
        st.caption(f"第 {n} 名出现并列，并列项一并显示（共 {shown_len} 项）。")


def vc_df(series, c1, c2):
    vc = series.value_counts()
    return pd.DataFrame({c1: vc.index.tolist(), c2: vc.values.tolist()})


def pct(n, d):
    return f"{n / d:.0%}" if d else "—"


def split_tags(tag_str):
    if pd.isna(tag_str):
        return []
    out = []
    for p in str(tag_str).split(","):
        p = re.sub(r"\s+", " ", p.strip().lower())
        if p and p != "无":
            out.append(TAG_NORM_MAP.get(p, p))
    return out


def flavor_complexity(n):
    if n == 0:
        return "无口味标签"
    if n == 1:
        return "单一口味"
    if n == 2:
        return "双重复合"
    return "三重以上复合"


def nums_in(v):
    """从形如 '5, 10, 20' / '20' / 20 / 10.5 的单元格中取出全部数值。"""
    if pd.isna(v):
        return []
    return [float(x) for x in re.findall(r"\d+\.?\d*", str(v))]


def parse_price(v):
    """瑞士站点价格为逗号小数，如 '7,50 CHF' / '19,90  CHF'。"""
    if pd.isna(v):
        return None
    s = str(v).replace("\xa0", " ").upper().replace("CHF", "").replace("'", "").strip()
    s = s.replace(",", ".")
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def parse_puffs_field(v):
    """puffs口数 字段，格式如 0.6k / 1k / 2k4 / 3k5 / 10k。"""
    if pd.isna(v):
        return None
    s = str(v).strip().lower()
    m = re.fullmatch(r"(\d+(?:\.\d+)?)k(\d)?", s)
    if m:
        base = float(m.group(1)) * 1000
        if m.group(2):
            base += int(m.group(2)) * 100
        return int(round(base))
    return int(s) if re.fullmatch(r"\d+", s) else None


PUFF_BINS = [0, 600, 1000, 3000, 8000, 10 ** 7]
PUFF_LABELS = ["≤600", "601–1K", "1K–3K", "3K–8K", ">8K"]


def puff_tier(series):
    return pd.cut(series, bins=PUFF_BINS, labels=PUFF_LABELS, right=True)


def norm_brand(x):
    x = re.sub(r"\s+", " ", str(x).strip())
    return BRAND_NORM_MAP.get(x.lower(), x)


# ============================================================
# 加载数据
# ============================================================
@st.cache_data
def load_data(path, mtime):
    xl = pd.ExcelFile(path)
    ps = pd.read_excel(xl, sheet_name=SHEET_POP)
    ps.columns = ps.columns.astype(str).str.strip()
    ps = ps.loc[:, [c for c in ps.columns if "Unnamed" not in c]]

    ps["网站名称"] = ps["网站名称"].astype(str).str.strip().str.replace(r"^www\.", "", regex=True)
    ps["产品类型"] = ps["产品类型"].astype(str).str.strip()
    ps["品牌_原始"] = ps["品牌"].astype(str).str.strip()
    ps["品牌"] = ps["品牌_原始"].apply(norm_brand)
    ps["含冰/薄荷"] = ps["含冰/薄荷"].fillna("不确定")
    ps["含烟草"] = ps["含烟草"].fillna("否")
    ps["分类"] = ps["分类"].fillna("其他")
    ps["品牌来源国"] = ps["品牌来源国"].fillna("未知").astype(str).str.strip()
    ps["是否本地代理/本土品牌"] = ps["是否本地代理/本土品牌"].fillna("否")
    ps["口味名称"] = ps["口味名称"].astype(str).str.strip()

    # 口味名归一：忽略大小写 / 多余空格，passionfruit → passion fruit（同一口味在不同站点写法不同，避免被拆成多个口味）
    ps["口味名_key"] = (ps["口味名称"].str.lower().str.replace("passionfruit", "passion fruit", regex=False)
                        .str.replace(r"\s+", " ", regex=True).str.strip())
    ps["口味名_显示"] = ps["口味名_key"].str.title()

    ps["口味标签列表"] = ps["口味标签"].apply(split_tags)
    # 数据校验：原表个别记录的「含冰/薄荷」「含烟草」与口味标签自相矛盾（如 Cola Ice 标“否”）。
    # 以口味标签为准做校正（只把“否→是”），原始值保留在 *_原始 列，并在总览页列出被校正的记录。
    ps["含冰/薄荷_原始"] = ps["含冰/薄荷"]
    ps["含烟草_原始"] = ps["含烟草"]
    fix_ice = (ps["含冰/薄荷"] != "是") & ps["口味标签列表"].apply(lambda l: any(t in ICE_TAGS for t in l))
    fix_tob = (ps["含烟草"] != "是") & ps["口味标签列表"].apply(lambda l: "tobacco" in l)
    ps.loc[fix_ice, "含冰/薄荷"] = "是"
    ps.loc[fix_tob, "含烟草"] = "是"
    ps["数据校正"] = ""
    ps.loc[fix_ice, "数据校正"] = "含冰/薄荷：否→是"
    ps.loc[fix_tob, "数据校正"] = ps.loc[fix_tob, "数据校正"].apply(lambda x: (x + "；" if x else "") + "含烟草：否→是")
    ps["口味标签数"] = ps["口味标签列表"].apply(len)
    ps["口味复杂度"] = ps["口味标签数"].apply(flavor_complexity)
    ps["无口味"] = ps["口味标签数"] == 0          # Nicotine Shot / 基液 / Flavourless 等

    ps["价格_CHF"] = ps["价格"].apply(parse_price)

    ps["尼古丁档位列表"] = ps["尼古丁浓度mg"].apply(nums_in)
    ps["尼古丁最高浓度"] = ps["尼古丁档位列表"].apply(lambda l: max(l) if l else None)

    ps["容量档位列表"] = ps["容量ml"].apply(nums_in)
    ps["容量_主规格ml"] = ps["容量档位列表"].apply(lambda l: l[0] if l else None)
    ps["容量_最大ml"] = ps["容量档位列表"].apply(lambda l: max(l) if l else None)
    ps["容量_单一规格"] = ps["容量档位列表"].apply(lambda l: len(set(l)) == 1)

    ps["设备类别_原始"] = ps["设备类别"].astype(str).str.strip()
    ps["设备类别_显示"] = ps["设备类别_原始"].map(DEVICE_NORM_MAP).fillna(ps["设备类别_原始"])

    ps["口数_数值"] = ps["puffs口数"].apply(parse_puffs_field)
    ps["榜单得分"] = (TOPN + 1 - ps["排名"]).clip(lower=1)   # 第1名=20分…第20名=1分，仅作相对热度参考，非销量
    ps["产品_key"] = (ps["品牌"] + " | " + ps["口味名_key"]).str.strip()

    # 单价（CHF/mL）：仅瓶装烟油且容量为单一规格时计算
    def unit_price(r):
        if r["设备类别_显示"] != FORM_EL or not r["容量_单一规格"]:
            return None
        if r["价格_CHF"] is None or pd.isna(r["价格_CHF"]) or not r["容量_主规格ml"]:
            return None
        return r["价格_CHF"] / r["容量_主规格ml"]
    ps["单价_CHF每ml"] = ps.apply(unit_price, axis=1)

    # 一次性：每百口价格（CHF / 100 puffs）
    ps["每百口价格_CHF"] = ps.apply(
        lambda r: (r["价格_CHF"] / r["口数_数值"] * 100)
        if (r["设备类别_显示"] == FORM_DP and r["价格_CHF"] is not None and not pd.isna(r["价格_CHF"])
            and r["口数_数值"]) else None, axis=1)

    # 容量超限线索：一次性 > 2 mL；瓶装烟油 > 10 mL
    ps["容量超限线索"] = (
        ((ps["设备类别_显示"] == FORM_DP) & (ps["容量_最大ml"] > DISPOSABLE_LIMIT_ML))
        | ((ps["设备类别_显示"] == FORM_EL) & (ps["容量_最大ml"] > REFILL_LIMIT_ML))
    )
    return ps


excel_path = resolve_excel()
if excel_path is None:
    st.error(f"❌ 未找到数据文件。请把 Excel 放到脚本同一目录：\n\n{_DIR}\n\n并确认文件名为：{os.path.basename(EXCEL_FILE)}")
    st.stop()
try:
    ps = load_data(excel_path, os.path.getmtime(excel_path))
except Exception as e:      # noqa
    st.error(f"❌ 读取数据失败：{e}\n\n请确认 Excel 内含 sheet「{SHEET_POP}」。")
    st.stop()

# ============================================================
# 侧边栏
# ============================================================
with st.sidebar:
    st.markdown("## 🔍 筛选（热销口味明细库）")
    st.markdown("---")
    all_devices = sorted(ps["设备类别_显示"].unique().tolist())
    sel_devices = st.multiselect("产品形态", all_devices, default=all_devices)
    all_countries = sorted(ps["品牌来源国"].dropna().unique().tolist())
    sel_countries = st.multiselect("品牌来源国", all_countries, default=all_countries)
    all_cats = sorted(ps["分类"].dropna().unique().tolist())
    sel_cats = st.multiselect("口味分类", all_cats, default=all_cats)
    all_sites = sorted(ps["网站名称"].dropna().unique().tolist())
    sel_sites = st.multiselect("数据来源网站", all_sites, default=all_sites)
    st.markdown("---")
    st.caption(f"💱 汇率固定：1 CHF = {CHF_CNY:g} CNY")
    st.markdown("---")
    st.caption(f"📅 数据文件日期：2026-10-20")
    st.caption("数据看板制作人：Erin")
    st.caption(f"热销口味库：{ps['网站名称'].nunique()} 个网站 · {len(ps)} 条记录")

fdf = ps[
    ps["设备类别_显示"].isin(sel_devices)
    & ps["品牌来源国"].isin(sel_countries)
    & ps["分类"].isin(sel_cats)
    & ps["网站名称"].isin(sel_sites)
].copy()

if fdf.empty:
    st.error("⚠️ 当前筛选条件下无数据，请放宽筛选范围。")
    st.stop()

n_el = int((fdf["设备类别_显示"] == FORM_EL).sum())
n_dp = int((fdf["设备类别_显示"] == FORM_DP).sum())

st.markdown("# 🇨🇭 瑞士电子烟市场调研数据看板")
st.caption(
    f"筛选后：{len(fdf)} 条热销口味记录（瓶装烟油 {n_el} · 一次性 {n_dp}）｜ 品牌 {fdf['品牌'].nunique()} 个 ｜"
    f" 网站 {fdf['网站名称'].nunique()} 个 ｜ 独立口味名 {fdf['口味名_key'].nunique()} 个"
)
st.markdown("---")

tabs = st.tabs([
    "📊 总览",
    "🍭 口味热榜",
    "🔖 口味元素",
    "🏷️ 品牌分析",
    "💉 产品规格",
    "💰 价格分析",
    "⚖️ 法规与税收",
    "📝 总结洞察",
])

# ============================================================
# TAB 0  总览
# ============================================================
with tabs[0]:
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    local_n = int((fdf["是否本地代理/本土品牌"] == "是").sum())
    ice_n = int((fdf["含冰/薄荷"] == "是").sum())
    kpis = [
        ("热销口味记录数", len(fdf), False),
        ("品牌数", fdf["品牌"].nunique(), False),
        ("独立口味名数", fdf["口味名_key"].nunique(), False),
        ("瑞士本土品牌占比", pct(local_n, len(fdf)), True),
        ("含冰/薄荷占比", pct(ice_n, len(fdf)), True),
        ("覆盖网站数", fdf["网站名称"].nunique(), False),
    ]
    for col, (label, val, alt) in zip([c1, c2, c3, c4, c5, c6], kpis):
        cls = "metric-card-alt" if alt else "metric-card"
        col.markdown(
            f'<div class="{cls}"><div class="value">{val}</div><div class="label">{label}</div></div>',
            unsafe_allow_html=True,
        )
    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<p class="section-title">🇨🇭 市场背景速览</p>', unsafe_allow_html=True)
    b1, b2, b3 = st.columns(3)
    with b1:
        st.markdown("""
        <div class="insight-box">
        <b>🌱 成长期市场，一次性在使用者中占比最高</b><br>
        EGL 2025：<b>4.1%</b> 的人口至少每月使用电子烟（15–19岁为 9.2%）；
        15 岁以上使用者约 <b>30–32 万</b>。
        EGL 2024：电子烟是瑞士第二常用的尼古丁产品；<b>54%</b> 的使用者主要使用一次性电子烟（15–24岁为 62%）。
        语区差异明显：德语区 3.1% / 法语区 7.4% / 意语区 5.2%。
        </div>
        """, unsafe_allow_html=True)
    with b2:
        st.markdown("""
        <div class="warning-box">
        <b>💰 消费税按“毫升”计，一次性税负重</b><br>
        一次性电子烟 <b>CHF 1.00/mL</b>（无论是否含尼古丁）；
        可重复使用系统仅对含尼古丁的液体/补充装/烟弹征收 <b>CHF 0.20/mL</b>（不含尼古丁的补充液不征税）。
        </div>
        """, unsafe_allow_html=True)
    with b3:
        st.markdown("""
        <div class="danger-box">
        <b>📢 监管快速收紧</b><br>
        2024-10-01 起全国统一 <b>18 岁</b>销售年龄；议会已于 2025-06-20 通过广告修订（BAG 预计 2027 年初生效，
        将覆盖销售点、网络与社交媒体广告）；联邦层面的一次性电子烟禁令动议已获议会通过，政府尚未提交草案，正评估以条例或修订 TabPG 的方式落实，
        瓦莱州、日内瓦州已先行实施禁令，其他多个州已有动议或草案。
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown('<p class="section-title">口味分类分布</p>', unsafe_allow_html=True)
        cc = vc_df(fdf["分类"], "分类", "数量")
        fig = px.bar(cc, x="分类", y="数量", color="分类", color_discrete_map=CAT_COLOR_MAP, text="数量", template=T)
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False, margin=dict(t=10, b=10), xaxis_title="", yaxis_title="数量")
        show_chart(fig)
        vcc = fdf["分类"].value_counts()
        st.caption(f"{vcc.index[0]}口味占比最高（{vcc.iloc[0] / len(fdf):.0%}）"
                   + (f"，其次为{vcc.index[1]}（{vcc.iloc[1] / len(fdf):.0%}）。" if len(vcc) > 1 else "。")
                   + "“其他”主要是 Nicotine Shot / 基液等无口味产品。")

    with col_r:
        st.markdown('<p class="section-title">产品形态 × 口味分类</p>', unsafe_allow_html=True)
        cross = fdf.groupby(["设备类别_显示", "分类"]).size().reset_index(name="数量")
        fig = px.bar(cross, x="设备类别_显示", y="数量", color="分类", barmode="stack",
                     color_discrete_map=CAT_COLOR_MAP, template=T)
        fig.update_layout(xaxis_title="", yaxis_title="数量", legend_title_text="分类", margin=dict(t=10, b=10))
        show_chart(fig)
        st.caption("样本中瓶装烟油与一次性电子烟的口味分类结构对比；两类样本量由各站榜单长度决定，不代表真实销量占比。")

    st.markdown("---")
    col_o1, col_o2 = st.columns(2)
    with col_o1:
        st.markdown('<p class="section-title">品牌来源国分布</p>', unsafe_allow_html=True)
        oc = vc_df(fdf["品牌来源国"], "来源国", "数量")
        fig_oc = px.pie(oc, values="数量", names="来源国", color="来源国",
                        color_discrete_map=COUNTRY_COLOR_MAP, hole=0.45, template=T)
        fig_oc.update_traces(textinfo="percent+label")
        fig_oc.update_layout(margin=dict(t=20, b=20))
        show_chart(fig_oc)
        voc = fdf["品牌来源国"].value_counts()
        ch_pct = voc.get("瑞士", 0) / len(fdf)
        cn_pct = voc.get("中国", 0) / len(fdf)
        st.markdown(f"""
        <div class="insight-box">
        <b>来源国最大项：{voc.index[0]}（{voc.iloc[0] / len(fdf):.0%}）</b>；
        瑞士本土品牌占 <b>{ch_pct:.0%}</b>，中国品牌占 <b>{cn_pct:.0%}</b>。
        </div>
        """, unsafe_allow_html=True)

    with col_o2:
        st.markdown('<p class="section-title">口味复杂度分布</p>', unsafe_allow_html=True)
        comp = vc_df(fdf["口味复杂度"], "复杂度", "数量")
        fig_comp = px.pie(comp, values="数量", names="复杂度", color="复杂度",
                          color_discrete_map={"无口味标签": "#cbd5e1", "单一口味": "#94a3b8",
                                              "双重复合": COLORS["primary"], "三重以上复合": COLORS["secondary"]},
                          hole=0.45, template=T)
        fig_comp.update_traces(textinfo="percent+label")
        fig_comp.update_layout(margin=dict(t=20, b=20))
        show_chart(fig_comp)
        multi_pct = fdf["口味复杂度"].isin(["双重复合", "三重以上复合"]).sum() / len(fdf)
        st.caption(f"复合口味（≥2个口味标签）占 {multi_pct:.0%}；“无口味标签”为 Nicotine Shot / 基液等。")

    st.markdown("---")
    st.markdown('<p class="section-title">📋 数据来源与调研范围说明（以下站点数/条数均由数据文件实时统计）</p>', unsafe_allow_html=True)
    el_all = ps[ps["设备类别_显示"] == FORM_EL]
    dp_all = ps[ps["设备类别_显示"] == FORM_DP]
    d1, d2, d3 = st.columns(3)
    with d1:
        st.markdown(f"""
        <div class="insight-box">
        <b>🏆 瓶装烟油榜</b><br>
        <b>{el_all['网站名称'].nunique()} 个网站</b>的 E-liquid 热销榜前 <b>{int(el_all['排名'].max())} 名</b>，共 <b>{len(el_all)}</b> 条记录。
        “出现次数”是榜单出现条数，不是销量。
        </div>
        """, unsafe_allow_html=True)
    with d2:
        st.markdown(f"""
        <div class="insight-box">
        <b>🚬 一次性电子烟榜</b><br>
        <b>{dp_all['网站名称'].nunique()} 个网站</b>的 Disposable 热销榜前 <b>{int(dp_all['排名'].max())} 名</b>，共 <b>{len(dp_all)}</b> 条记录。
        数据集作为热门口味参考榜。
        </div>
        """, unsafe_allow_html=True)
    with d3:
        st.markdown("""
        <div class="insight-box">
        <b>⚖️ 法规信息来源</b><br>
        以 TabPG（烟草制品法）/ TabPV、联邦消费税规则及 BAG / BFS 官方调查为主线整理，详见「法规与税收」页。
        </div>
        """, unsafe_allow_html=True)

    # --- 榜单长度与来源 URL（动态）---
    list_len = ps.groupby(["网站名称", "设备类别_显示"]).agg(条数=("排名", "size"), 最大排名=("排名", "max")).reset_index()
    list_src = ps.groupby(["网站名称", "设备类别_显示"])["榜单来源"].first().reset_index()
    list_tbl = list_len.merge(list_src, on=["网站名称", "设备类别_显示"]).rename(
        columns={"设备类别_显示": "产品形态", "榜单来源": "榜单链接"})
    with st.expander("查看各榜单条数与来源链接"):
        show_df(list_tbl)
    corr = ps[ps["数据校正"] != ""]
    if len(corr):
        with st.expander(f"查看数据校验：{len(corr)} 条记录的「含冰/薄荷」「含烟草」标记与口味标签矛盾，已按标签校正"):
            show_df(corr[["网站名称", "设备类别_显示", "品牌", "口味名称", "口味标签", "含冰/薄荷_原始", "含烟草_原始", "数据校正"]].rename(
                columns={"网站名称": "网站", "设备类别_显示": "产品形态", "含冰/薄荷_原始": "原含冰/薄荷", "含烟草_原始": "原含烟草"}))
            st.caption("校正只把“否”改为“是”（标签含 ice / menthol / koolada 或 tobacco）；原始 Excel 未改动，建议同步修正源表。")

    st.markdown('<p class="section-title">⚠️ 调研瓶颈与数据局限性（部分条目由数据文件自动检测）</p>', unsafe_allow_html=True)
    notes = [
        "<b>样本覆盖局限：</b>仅为各网站“热销排序”榜单的头部条目（线上渠道），未涵盖线下门店、烟草店、加油站与社交媒体；榜单排序为 bestseller / popularity，非公开销量。",
    ]
    nos_dp = list_len[(list_len["网站名称"] == "nosmoky.ch") & (list_len["设备类别_显示"] == FORM_DP)]
    if len(nos_dp) and int(nos_dp["条数"].iloc[0]) != NOSMOKY_DP_EXPECTED:
        notes.append(
            f"<b>nosmoky.ch 一次性条数与说明不一致：</b>调研概览写明该站一次性“只取前 {NOSMOKY_DP_EXPECTED} 名”，"
            f"但数据文件中实际有 <b>{int(nos_dp['条数'].iloc[0])}</b> 条（该站一次性总共 12 个产品）。看板按数据文件实际条数统计，如需严格按前 10 名，请删除多出的记录。")
    salt_sites = sorted(ps[(ps["设备类别_显示"] == FORM_EL) & ps["榜单来源"].str.contains("salt|salz", case=False, na=False)]["网站名称"].unique())
    if salt_sites:
        notes.append(f"<b>烟油榜单口径偏向尼古丁盐：</b>{'、'.join(salt_sites)} 的烟油榜单取自“尼古丁盐（nicotine salt）”分类页，"
                     "因此烟油样本偏向尼古丁盐产品，不等同于全部 E-liquid。")
    pod_sites = sorted(ps[(ps["设备类别_显示"] == FORM_DP) & ps["榜单来源"].str.contains("pod", case=False, na=False)]["网站名称"].unique())
    if pod_sites:
        notes.append(f"<b>“一次性”榜单混入预填充 Pod 系统：</b>{'、'.join(pod_sites)} 的“一次性”榜单链接为 prefilled pods 分类页，"
                     "其中大容量（> 2 mL）条目可能属于 Pod 系统而非严格意义上的一次性产品，但口味热度依旧可以参考。")
    if len(corr):
        notes.append(f"<b>数据校验：</b>有 {len(corr)} 条记录的冰感/烟草标记与口味标签自相矛盾（如口味名含 Ice 却标注“否”），看板已按口味标签校正，明细见上方折叠区。")
    notes.append("<b>价格口径：</b>标价为网站展示价（CHF），多浓度/多规格产品只记一个标价；未区分促销价与原价。")
    st.markdown('<div class="warning-box"><ul style="margin:0; padding-left:20px;">'
                + "".join(f"<li>{n}</li>" for n in notes) + "</ul></div>", unsafe_allow_html=True)

# ============================================================
# TAB 1  口味热榜
# ============================================================
with tabs[1]:
    mode = st.radio("统计口径", ["按口味名称", "按 品牌 + 口味名称"], horizontal=True, key="flavor_mode",
                    help="同名口味（如 Grape）可能来自不同品牌；选择“品牌+口味名称”可避免跨品牌合并。")
    fl_df = fdf[~fdf["无口味"]].copy()          # 口味榜不统计 Nicotine Shot / 基液等无口味产品
    n_nf = int(fdf["无口味"].sum())
    if fl_df.empty:
        st.info("当前筛选范围内没有带口味标签的记录。")
    else:
        fl_df["_口味键"] = fl_df["口味名_显示"] if mode == "按口味名称" else (fl_df["品牌"] + " · " + fl_df["口味名_显示"])
        st.caption(f"口味榜已排除 {n_nf} 条无口味标签记录（Nicotine Shot / 基液 / Flavourless 等），以下统计基于 {len(fl_df)} 条带口味记录；口味名称已统一大小写及写法。")

        col_l, col_r = st.columns(2)
        fl_counts = fl_df["_口味键"].value_counts()
        with col_l:
            st.markdown('<p class="section-title">全榜热门口味 Top 15</p>', unsafe_allow_html=True)
            fl_top = top_n_with_ties(fl_counts, 15)
            fl = pd.DataFrame({"口味": fl_top.index.tolist(), "出现次数": fl_top.values.tolist()})
            fl = fl.sort_values("出现次数", ascending=True)
            fig = px.bar(fl, x="出现次数", y="口味", orientation="h", text="出现次数", color="出现次数",
                         color_continuous_scale=RED_SCALE, template=T)
            fig.update_traces(textposition="outside")
            fig.update_layout(height=500, margin=dict(t=10, b=10), xaxis_title="", yaxis_title="",
                              coloraxis_showscale=False, yaxis={"categoryorder": "total ascending"})
            show_chart(fig)
            tie_cap(len(fl_top), 15)
            top3s = top_n_with_ties(fl_counts, 3)
            top3 = " / ".join(f"{k}（{v}）" for k, v in top3s.items())
            tie_note = "（第三名出现并列，并列项一并列出）" if len(top3s) > 3 else ""
            st.caption(f"当前筛选下出现最多的口味：{top3}{tie_note}。出现次数 = 在各站点热销榜单中的记录条数。")

        with col_r:
            st.markdown('<p class="section-title">按产品形态查看 Top 15</p>', unsafe_allow_html=True)
            dtype = st.selectbox("选择产品形态", fl_df["设备类别_显示"].unique(), key="tab1_type")
            sub_vc = top_n_with_ties(fl_df[fl_df["设备类别_显示"] == dtype]["_口味键"].value_counts(), 15)
            sub_fl = pd.DataFrame({"口味": sub_vc.index.tolist(), "出现次数": sub_vc.values.tolist()})
            sub_fl = sub_fl.sort_values("出现次数", ascending=True)
            fig = px.bar(sub_fl, x="出现次数", y="口味", orientation="h", text="出现次数",
                         color_discrete_sequence=[FORM_COLOR_MAP.get(dtype, COLORS["primary"])], template=T)
            fig.update_traces(textposition="outside")
            fig.update_layout(height=500, margin=dict(t=10, b=10), xaxis_title="", yaxis_title="",
                              yaxis={"categoryorder": "total ascending"})
            show_chart(fig)
            tie_cap(len(sub_vc), 15)

        st.markdown('<p class="section-title">Top 15 口味详细成分标签</p>', unsafe_allow_html=True)
        tie_cap(len(fl_top), 15)
        cols2 = st.columns(2)
        for i, (fname, cnt) in enumerate(fl_top.items()):
            sub = fl_df[fl_df["_口味键"] == fname]
            tag_freq = Counter([t for tl in sub["口味标签列表"] for t in tl])
            tags_html = " ".join(f'<span class="tag-chip">{t}</span>' for t, _ in tag_freq.most_common()) \
                or '<span style="color:#9ca3af">无标签</span>'
            sites_n = sub["网站名称"].nunique()
            with cols2[i % 2]:
                st.markdown(
                    f'<div class="flavor-row"><b>#{i + 1} {fname}</b>（{cnt}次 · {sites_n}个站点）<br>{tags_html}</div>',
                    unsafe_allow_html=True)

    st.markdown("---")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('<p class="section-title">含冰/薄荷分布</p>', unsafe_allow_html=True)
        ice_cnt = vc_df(fdf["含冰/薄荷"], "状态", "数量")
        fig_ice = px.pie(ice_cnt, values="数量", names="状态", color="状态",
                         color_discrete_map={"是": COLORS["ice"], "否": "#94a3b8", "不确定": "#e2e8f0"},
                         hole=0.45, template=T)
        fig_ice.update_traces(textinfo="percent+label")
        fig_ice.update_layout(margin=dict(t=20, b=20))
        show_chart(fig_ice)
        ice_yes = int((fdf["含冰/薄荷"] == "是").sum())
        st.caption(f"明确含冰/薄荷 {ice_yes} 条（{ice_yes / len(fdf):.0%}）。")

    with col_b:
        st.markdown('<p class="section-title">含烟草分布</p>', unsafe_allow_html=True)
        tb_cnt = vc_df(fdf["含烟草"], "状态", "数量")
        fig_tb = px.pie(tb_cnt, values="数量", names="状态", color="状态",
                        color_discrete_map={"是": COLORS["tobacco"], "否": "#94a3b8"}, hole=0.45, template=T)
        fig_tb.update_traces(textinfo="percent+label")
        fig_tb.update_layout(margin=dict(t=20, b=20))
        show_chart(fig_tb)
        tb_yes = int((fdf["含烟草"] == "是").sum())
        st.caption(f"含烟草口味 {tb_yes} 条（{tb_yes / len(fdf):.0%}）。")

    st.markdown('<p class="section-title">🔥 口味分类 × 含冰/薄荷热力图</p>', unsafe_allow_html=True)
    cross_ice = pd.crosstab(fdf["分类"], fdf["含冰/薄荷"])
    fig_hi = px.imshow(cross_ice, text_auto=True, aspect="auto", color_continuous_scale=HEAT_SCALE, template=T)
    fig_hi.update_layout(xaxis_title="", yaxis_title="", coloraxis_showscale=False, height=320, margin=dict(t=10, b=10))
    show_chart(fig_hi)
    if "水果" in cross_ice.index and "是" in cross_ice.columns:
        fr = cross_ice.loc["水果"]
        st.caption(f"水果类中明确含冰/薄荷的有 {int(fr['是'])} 条，占该分类 {fr['是'] / fr.sum():.0%}。")

    st.markdown("---")
    st.markdown('<p class="section-title">含烟草口味常见复合元素（不含"tobacco"本身）</p>', unsafe_allow_html=True)
    tb_df = fdf[fdf["含烟草"] == "是"]
    if tb_df.empty:
        st.info("当前筛选范围内暂无含烟草标记的产品。")
    else:
        tb_tags = Counter()
        for tl in tb_df["口味标签列表"]:
            for t in tl:
                if t != "tobacco":
                    tb_tags[t] += 1
        if tb_tags:
            html = " ".join(f'<span class="tag-chip">{t}（{c}）</span>' for t, c in tb_tags.most_common(15))
            st.markdown(f'<div class="flavor-row">{html}</div>', unsafe_allow_html=True)
        pure = int((tb_df["口味标签数"] <= 1).sum())
        st.caption(f"含烟草记录 {len(tb_df)} 条，其中仅标注单一“烟草”元素的有 {pure} 条（{pure / len(tb_df):.0%}）。")

# ============================================================
# TAB 2  口味元素
# ============================================================
with tabs[2]:
    all_tags = [t for tl in fdf["口味标签列表"] for t in tl]
    tag_counts = Counter(all_tags)

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown('<p class="section-title">高频口味元素 Top 20</p>', unsafe_allow_html=True)
        if tag_counts:
            _tag_top = top_n_with_ties(pd.Series(dict(tag_counts.most_common())), 20)
            top_tags_df = pd.DataFrame({"口味元素": _tag_top.index.tolist(), "出现次数": _tag_top.values.tolist()}).sort_values("出现次数")
            fig = px.bar(top_tags_df, x="出现次数", y="口味元素", orientation="h", text="出现次数", color="出现次数",
                         color_continuous_scale=RED_SCALE, template=T)
            fig.update_traces(textposition="outside")
            fig.update_layout(coloraxis_showscale=False, height=560, margin=dict(t=10, b=10),
                              xaxis_title="", yaxis_title="", yaxis={"categoryorder": "total ascending"})
            show_chart(fig)
            tie_cap(len(_tag_top), 20)
            _tc3 = top_n_with_ties(pd.Series(dict(tag_counts.most_common())), 3)
            st.caption("出现最高频的元素：" + " / ".join(f"{t}（{c}）" for t, c in _tc3.items())
                       + "。标签已统一大小写及个别拼写差异（如 passionfruit / passion fruit）。")
        else:
            st.info("暂无口味标签。")

    with col_r:
        st.markdown('<p class="section-title">各口味分类的高频元素</p>', unsafe_allow_html=True)
        for cat in fdf["分类"].value_counts().index.tolist():
            cdf = fdf[fdf["分类"] == cat]
            top6 = Counter([t for tl in cdf["口味标签列表"] for t in tl]).most_common(6)
            tags_html = " ".join(f'<span class="tag-chip">{t}（{c}）</span>' for t, c in top6) \
                or '<span style="color:#9ca3af">无标签</span>'
            st.markdown(f'<div class="flavor-row"><b>{cat}</b>（{len(cdf)}条）&nbsp;&nbsp;{tags_html}</div>',
                        unsafe_allow_html=True)

    st.markdown("---")
    st.markdown('<p class="section-title">🔗 双元素组合频率（Top 12）</p>', unsafe_allow_html=True)
    combos = Counter()
    for tl in fdf["口味标签列表"]:
        uniq = sorted(set(tl))
        for ii in range(len(uniq)):
            for jj in range(ii + 1, len(uniq)):
                combos[(uniq[ii], uniq[jj])] += 1
    _all_combos = combos.most_common()
    if _all_combos:
        _cut = _all_combos[min(11, len(_all_combos) - 1)][1]
        top_combos = [x for x in _all_combos if x[1] >= _cut] if _cut >= 2 else _all_combos[:12]   # 第 12 名并列（≥2 次）时一并纳入
    else:
        top_combos = []
    if top_combos:
        combo_df = pd.DataFrame([(f"{a} + {b}", c) for (a, b), c in top_combos],
                                columns=["组合", "出现次数"]).sort_values("出现次数")
        fig_cb = px.bar(combo_df, x="出现次数", y="组合", orientation="h", text="出现次数", color="出现次数",
                        color_continuous_scale=SLATE_SCALE, template=T)
        fig_cb.update_traces(textposition="outside")
        fig_cb.update_layout(coloraxis_showscale=False, height=420, margin=dict(t=10, b=10),
                             yaxis={"categoryorder": "total ascending"})
        show_chart(fig_cb)
        tie_cap(len(top_combos), 12)
        st.caption(f"最高频双元素组合：{top_combos[0][0][0]} + {top_combos[0][0][1]}（{top_combos[0][1]}次）。")
    else:
        st.info("暂无足够数据生成双元素组合。")

# ============================================================
# TAB 3  品牌分析
# ============================================================
with tabs[3]:
    st.markdown("## 🏷️ 品牌维度分析（热销口味库）")
    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown('<p class="section-title">品牌出现频次 Top 12</p>', unsafe_allow_html=True)
        bvc = top_n_with_ties(fdf["品牌"].value_counts(), 12)
        brand_cnt = pd.DataFrame({"品牌": bvc.index.tolist(), "出现次数": bvc.values.tolist()})
        origin_map = fdf.drop_duplicates("品牌").set_index("品牌")["品牌来源国"].to_dict()
        brand_cnt["来源国"] = brand_cnt["品牌"].map(origin_map)
        brand_cnt = brand_cnt.sort_values("出现次数")
        fig_bc = px.bar(brand_cnt, x="出现次数", y="品牌", orientation="h", text="出现次数", color="来源国",
                        color_discrete_map=COUNTRY_COLOR_MAP, template=T)
        fig_bc.update_traces(textposition="outside")
        fig_bc.update_layout(height=480, margin=dict(t=10, b=10), xaxis_title="", yaxis_title="",
                             legend_title_text="品牌来源国", yaxis={"categoryorder": "total ascending"})
        show_chart(fig_bc)
        tie_cap(len(bvc), 12)
        bc_all = fdf["品牌"].value_counts()
        cr1 = bc_all.iloc[0] / len(fdf)
        bc_top3 = top_n_with_ties(bc_all, 3)
        cr3 = bc_top3.sum() / len(fdf)
        st.caption(f"{bc_all.index[0]} 以 {bc_all.iloc[0]} 次居首（占当前记录 {cr1:.0%}）；前三名{'（含并列）' if len(bc_top3) > 3 else ''}合计 {cr3:.0%}。"
                   "出现次数反映榜单占位，受各站榜单长度和变体拆分影响。")

    with col_r:
        st.markdown('<p class="section-title">瑞士本土品牌占比</p>', unsafe_allow_html=True)
        local_cnt = vc_df(fdf["是否本地代理/本土品牌"], "是否本土", "数量")
        fig_local = px.pie(local_cnt, values="数量", names="是否本土", color="是否本土",
                           color_discrete_map={"是": COLORS["primary"], "否": "#94a3b8"}, hole=0.45, template=T)
        fig_local.update_traces(textinfo="percent+label")
        fig_local.update_layout(margin=dict(t=20, b=20))
        show_chart(fig_local)
        loc_df = fdf[fdf["是否本地代理/本土品牌"] == "是"]
        if len(loc_df):
            loc_desc = "；".join(
                f"{b}（{int((loc_df['品牌'] == b).sum())}条，{'/'.join(sorted(loc_df[loc_df['品牌'] == b]['设备类别_显示'].map(lambda x: '烟油' if x == FORM_EL else '一次性').unique()))}）"
                for b in loc_df["品牌"].value_counts().index)
            st.caption(f"本土品牌记录占 {local_n / len(fdf):.0%}：{loc_desc}。其余为进口品牌。")
        else:
            st.caption("当前筛选下无本土品牌记录。")

    st.markdown('<p class="section-title">网站 × 品牌出现频次矩阵（Top 10 品牌）</p>', unsafe_allow_html=True)
    top_brands_list = top_n_with_ties(fdf["品牌"].value_counts(), 10).index.tolist()
    site_brand = pd.crosstab(fdf["网站名称"], fdf["品牌"])
    site_brand_top = site_brand[[c for c in top_brands_list if c in site_brand.columns]]
    fig_sb = px.imshow(site_brand_top, text_auto=True, aspect="auto", color_continuous_scale=HEAT_SCALE, template=T)
    fig_sb.update_layout(xaxis_title="品牌", yaxis_title="网站", coloraxis_showscale=False, height=380,
                         margin=dict(t=10, b=10))
    show_chart(fig_sb)
    spread = fdf.groupby("品牌")["网站名称"].nunique().sort_values(ascending=False)
    st.caption(f"铺货最广的品牌：{spread.index[0]}（出现在 {spread.iloc[0]} 个站点）；"
               f"仅出现在单一站点的品牌 {int((spread == 1).sum())} 个。")

    st.markdown('<p class="section-title">品牌代表性口味（可选择）</p>', unsafe_allow_html=True)
    brand_sel = st.selectbox("选择品牌", sorted(fdf["品牌"].unique()), key="brand_flavor_sel")
    bvc2 = top_n_with_ties(fdf[fdf["品牌"] == brand_sel]["口味名_显示"].value_counts(), 10)
    if not bvc2.empty:
        bf = pd.DataFrame({"口味名称": bvc2.index.tolist(), "出现次数": bvc2.values.tolist()}).sort_values("出现次数")
        fig_bf = px.bar(bf, x="出现次数", y="口味名称", orientation="h", text="出现次数",
                        color_discrete_sequence=[COLORS["primary"]], template=T)
        fig_bf.update_traces(textposition="outside")
        fig_bf.update_layout(height=380, margin=dict(t=10, b=10), yaxis={"categoryorder": "total ascending"})
        show_chart(fig_bf)
        tie_cap(len(bvc2), 10)

    col_cc1, col_cc2 = st.columns(2)
    with col_cc1:
        st.markdown('<p class="section-title">品牌来源国 × 产品形态</p>', unsafe_allow_html=True)
        ct = fdf.groupby(["品牌来源国", "设备类别_显示"]).size().reset_index(name="数量")
        fig_ct = px.bar(ct, x="品牌来源国", y="数量", color="设备类别_显示", barmode="stack",
                        color_discrete_map=FORM_COLOR_MAP, template=T)
        fig_ct.update_layout(xaxis_title="", yaxis_title="记录数", legend_title_text="", margin=dict(t=10, b=10))
        show_chart(fig_ct)
        comp_parts, loc_parts = [], []
        for form, short in [(FORM_DP, "一次性"), (FORM_EL, "烟油")]:
            sub = fdf[fdf["设备类别_显示"] == form]
            if sub.empty:
                continue
            comp = sub["品牌来源国"].value_counts(normalize=True)
            comp_parts.append(f"{short}榜单（{len(sub)} 条）：" + "、".join(f"{k} {v:.0%}" for k, v in comp.items()))
            loc = sub[sub["品牌来源国"] == "瑞士"]["品牌"].value_counts()
            if len(loc):
                loc_parts.append(f"{short}为 " + "、".join(f"{b}（{n}）" for b, n in loc.items()))
        cap_txt = "各形态内的来源国构成 — " + "；".join(comp_parts) + "。"
        if loc_parts:
            cap_txt += " 瑞士本土品牌：" + "；".join(loc_parts) + "。"
        cap_txt += " 比例按当前筛选后各形态的记录数计算，反映的是榜单构成，不代表市场份额。"
        st.caption(cap_txt)

    with col_cc2:
        st.markdown('<p class="section-title">品牌来源国 × 口味分类</p>', unsafe_allow_html=True)
        ccat = fdf.groupby(["品牌来源国", "分类"]).size().reset_index(name="数量")
        fig_cc = px.bar(ccat, x="品牌来源国", y="数量", color="分类", barmode="stack",
                        color_discrete_map=CAT_COLOR_MAP, template=T)
        fig_cc.update_layout(xaxis_title="", yaxis_title="记录数", legend_title_text="分类", margin=dict(t=10, b=10))
        show_chart(fig_cc)
        st.caption("各来源国的口味分类构成（记录数口径）；样本量小的来源国（如美国）仅供参考。")

    st.markdown('<p class="section-title">🏅 品牌综合热度榜（Top 15）</p>', unsafe_allow_html=True)
    bh = fdf.groupby(["品牌", "品牌来源国"]).agg(上榜次数=("排名", "size"), 覆盖站点数=("网站名称", "nunique"),
                                              平均排名=("排名", "mean"), 榜单得分=("榜单得分", "sum")).reset_index()
    bh["平均排名"] = bh["平均排名"].round(1)
    bh = bh.sort_values(["榜单得分", "上榜次数"], ascending=False)
    if len(bh) > 15:
        bh = bh[bh["榜单得分"] >= bh["榜单得分"].iloc[14]]      # 第 15 名得分并列时一并纳入
    show_df(bh)
    tie_cap(len(bh), 15)
    st.caption("榜单得分 = 各榜单（21 − 排名）之和，同时反映上榜次数与名次，仅作相对热度参考，不是销量；会随左侧「产品形态」等筛选联动。")

    st.markdown('<p class="section-title">🥇 各榜单第 1 名</p>', unsafe_allow_html=True)
    first = fdf[fdf["排名"] == 1][["网站名称", "设备类别_显示", "品牌", "口味名称", "品牌来源国", "puffs口数", "价格_CHF"]].rename(
        columns={"网站名称": "网站", "设备类别_显示": "产品形态", "口味名称": "口味", "puffs口数": "口数", "价格_CHF": "标价(CHF)"})
    if first.empty:
        st.info("当前筛选下没有保留榜首记录的榜单。")
    else:
        show_df(first)
        fv = first["品牌"].value_counts()
        st.caption("榜首品牌分布：" + "、".join(f"{k} {v} 个榜单" for k, v in fv.items()) + "。仅统计筛选后仍保留榜首记录的榜单。")

    st.markdown("---")
    st.markdown('<p class="section-title">🔁 跨站点重复上榜的产品（品牌 + 口味名称，出现在 ≥2 个榜单）</p>', unsafe_allow_html=True)
    rep_src = fdf[~fdf["无口味"]]
    mm = rep_src.groupby(["设备类别_显示", "产品_key"]).agg(
        品牌=("品牌", "first"), 口味=("口味名_显示", "first"), 来源国=("品牌来源国", "first"),
        上榜站点数=("网站名称", "nunique"), 平均排名=("排名", "mean"), 得分=("榜单得分", "sum")).reset_index()
    mm = mm[mm["上榜站点数"] >= 2].sort_values(["上榜站点数", "得分"], ascending=False)
    if mm.empty:
        st.info("当前筛选下没有跨站点重复上榜的产品。")
    else:
        mm_show = mm.rename(columns={"设备类别_显示": "产品形态"})[["产品形态", "品牌", "口味", "来源国", "上榜站点数", "平均排名", "得分"]].head(30).copy()
        mm_show["平均排名"] = mm_show["平均排名"].round(1)
        show_df(mm_show)
        st.caption("按“品牌 + 口味名称（忽略大小写）”精确匹配，写法不同的同款产品不会被合并，重复数可能偏保守；仅显示前 30 条。"
                   "得分 = 各榜单（21 − 排名）之和，仅作相对热度参考。")

# ============================================================
# TAB 4  产品规格
# ============================================================
with tabs[4]:
    st.markdown("## 💉 产品规格分析")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown('<p class="section-title">尼古丁浓度档位分布（mg/mL）</p>', unsafe_allow_html=True)
        nic_flat = [n for nl in fdf["尼古丁档位列表"] for n in nl]
        if nic_flat:
            nic_df = pd.DataFrame(Counter(nic_flat).items(), columns=["浓度(mg)", "出现次数"]).sort_values("浓度(mg)")
            nic_df["对照"] = nic_df["浓度(mg)"].apply(
                lambda x: "> 20 mg（超上限）" if x > NIC_LIMIT_MG_ML else ("= 20 mg（法规上限）" if x == NIC_LIMIT_MG_ML else "< 20 mg"))
            nic_df["浓度(mg)"] = nic_df["浓度(mg)"].apply(lambda x: f"{x:g}")
            fig_nic = px.bar(nic_df, x="浓度(mg)", y="出现次数", text="出现次数", color="对照",
                             color_discrete_map={"< 20 mg": COLORS["secondary"], "= 20 mg（法规上限）": COLORS["primary"],
                                                 "> 20 mg（超上限）": "#7f1d1d"}, template=T)
            fig_nic.update_traces(textposition="outside")
            fig_nic.update_layout(margin=dict(t=10, b=10), xaxis_title="尼古丁浓度（mg/mL）", yaxis_title="出现次数",
                                  legend_title_text="", xaxis={"type": "category"})
            show_chart(fig_nic)
            top_nic = Counter(nic_flat).most_common(1)[0]
            over = sum(1 for n in nic_flat if n > NIC_LIMIT_MG_ML)
            at_cap = sum(1 for n in nic_flat if n == NIC_LIMIT_MG_ML)
            st.caption(f"{top_nic[0]:g} mg 档位出现最多（{top_nic[1]}次）；档位总计 {len(nic_flat)} 个，其中恰为 20 mg 上限的 {at_cap} 个"
                       f"（{at_cap / len(nic_flat):.0%}），> 20 mg 的 {over} 个；同一产品多档位分别计数。")
        else:
            st.info("暂无尼古丁浓度数据。")

        hi = fdf[fdf["尼古丁最高浓度"] > NIC_LIMIT_MG_ML]
        if len(hi):
            st.markdown(f"""
            <div class="danger-box">
            当前筛选下有 <b>{len(hi)}</b> 条记录标注浓度 &gt; {NIC_LIMIT_MG_ML} mg，涉及品牌：{"、".join(sorted(hi['品牌'].unique()))}；
            对照 TabPG：含尼古丁液体须 ≤ {NIC_LIMIT_MG_ML} mg/mL。字段按 mg/mL 理解，需逐条回到详情页核实。
            </div>
            """, unsafe_allow_html=True)
        else:
            top20 = int((fdf["尼古丁最高浓度"] == NIC_LIMIT_MG_ML).sum())
            st.markdown(f"""
            <div class="insight-box">
            <b>合规对照：</b>当前筛选下没有标注浓度 &gt; {NIC_LIMIT_MG_ML} mg 的记录；
            最高档恰为 {NIC_LIMIT_MG_ML} mg（法规上限）的有 <b>{top20}</b> 条（{top20 / len(fdf):.0%}），即多数产品的最高档直接贴着上限。
            </div>
            """, unsafe_allow_html=True)

    with col_b:
        st.markdown('<p class="section-title">尼古丁浓度 × 产品形态</p>', unsafe_allow_html=True)
        rows = []
        for _, r in fdf.iterrows():
            for n in r["尼古丁档位列表"]:
                rows.append({"产品形态": r["设备类别_显示"], "尼古丁浓度(mg)": f"{n:g}", "_n": n})
        ntc = pd.DataFrame(rows)
        if not ntc.empty:
            agg = ntc.groupby(["产品形态", "尼古丁浓度(mg)", "_n"]).size().reset_index(name="出现次数").sort_values("_n")
            fig_ntc = px.bar(agg, x="尼古丁浓度(mg)", y="出现次数", color="产品形态", barmode="group",
                             color_discrete_map=FORM_COLOR_MAP, template=T)
            fig_ntc.update_layout(xaxis_title="尼古丁浓度（mg/mL）", yaxis_title="出现次数", legend_title_text="",
                                  margin=dict(t=10, b=10), xaxis={"type": "category"})
            show_chart(fig_ntc)
            dp_nic = fdf[fdf["设备类别_显示"] == FORM_DP]["尼古丁最高浓度"].dropna()
            if len(dp_nic):
                z = int((fdf[fdf["设备类别_显示"] == FORM_DP]["尼古丁档位列表"].apply(lambda l: 0 in l)).sum())
                st.caption(f"一次性电子烟记录标注的最高浓度：{', '.join(f'{v:g}' for v in sorted(dp_nic.unique()))} mg；"
                           f"其中有 {z} 条同时提供 0 mg（无尼古丁）选项。")

    st.markdown("---")
    cap_a, cap_b = st.columns(2)
    with cap_a:
        st.markdown(f'<p class="section-title">瓶装烟油容量规格（补充装上限 {REFILL_LIMIT_ML} mL）</p>', unsafe_allow_html=True)
        el_df = fdf[fdf["设备类别_显示"] == FORM_EL]
        cap_flat = [c for cl in el_df["容量档位列表"] for c in cl]
        if cap_flat:
            cap_df = pd.DataFrame(Counter(cap_flat).items(), columns=["容量(ml)", "出现次数"]).sort_values("容量(ml)")
            cap_df["容量标签"] = cap_df["容量(ml)"].apply(lambda x: f"{x:g}ml")
            cap_df["对照"] = cap_df["容量(ml)"].apply(lambda x: f"≤{REFILL_LIMIT_ML}ml" if x <= REFILL_LIMIT_ML else f">{REFILL_LIMIT_ML}ml")
            fig_cap = px.bar(cap_df, x="容量标签", y="出现次数", text="出现次数", color="对照",
                             color_discrete_map={f"≤{REFILL_LIMIT_ML}ml": COLORS["secondary"], f">{REFILL_LIMIT_ML}ml": COLORS["gold"]},
                             template=T)
            fig_cap.update_traces(textposition="outside")
            fig_cap.update_layout(margin=dict(t=10, b=10), xaxis_title="容量规格", yaxis_title="出现次数",
                                  legend_title_text="", xaxis={"type": "category"})
            show_chart(fig_cap)
            big_cap = int((el_df["容量_最大ml"] > REFILL_LIMIT_ML).sum())
            st.caption(f"瓶装烟油记录 {len(el_df)} 条，其中容量 > {REFILL_LIMIT_ML} mL 的 {big_cap} 条（{pct(big_cap, len(el_df))}）。")
        else:
            st.info("暂无瓶装烟油容量数据。")

    with cap_b:
        st.markdown(f'<p class="section-title">一次性电子烟容量规格（一次性上限 {DISPOSABLE_LIMIT_ML} mL）</p>', unsafe_allow_html=True)
        dp_df = fdf[fdf["设备类别_显示"] == FORM_DP]
        dcap = dp_df["容量_主规格ml"].dropna()
        if len(dcap):
            dc = dcap.value_counts().reset_index()
            dc.columns = ["容量(ml)", "记录数"]
            dc = dc.sort_values("容量(ml)")
            dc["容量标签"] = dc["容量(ml)"].apply(lambda x: f"{x:g}ml")
            dc["对照"] = dc["容量(ml)"].apply(lambda x: f"≤{DISPOSABLE_LIMIT_ML}ml" if x <= DISPOSABLE_LIMIT_ML else f">{DISPOSABLE_LIMIT_ML}ml")
            fig_dc = px.bar(dc, x="容量标签", y="记录数", text="记录数", color="对照",
                            color_discrete_map={f"≤{DISPOSABLE_LIMIT_ML}ml": COLORS["primary"], f">{DISPOSABLE_LIMIT_ML}ml": COLORS["gold"]},
                            template=T)
            fig_dc.update_traces(textposition="outside")
            fig_dc.update_layout(margin=dict(t=10, b=10), xaxis_title="容量规格", yaxis_title="记录数",
                                 legend_title_text="", xaxis={"type": "category"})
            show_chart(fig_dc)
            bigd = int((dcap > DISPOSABLE_LIMIT_ML).sum())
            st.caption(f"一次性记录 {len(dcap)} 条，其中容量 > {DISPOSABLE_LIMIT_ML} mL 的 {bigd} 条（{bigd / len(dcap):.0%}）。")
        else:
            st.info("当前筛选下暂无一次性容量数据。")

    st.markdown('<p class="section-title">⚠️ 合规线索：容量超过法规上限的记录</p>', unsafe_allow_html=True)
    clue = fdf[fdf["容量超限线索"]].copy()
    if clue.empty:
        st.info("当前筛选范围内暂无容量超限线索。")
    else:
        clue["对照上限(mL)"] = clue["设备类别_显示"].map({FORM_DP: DISPOSABLE_LIMIT_ML, FORM_EL: REFILL_LIMIT_ML})
        grp = clue.groupby(["网站名称", "设备类别_显示", "品牌", "容量_最大ml", "对照上限(mL)"]).agg(
            记录数=("排名", "size"), 口数=("puffs口数", "first"), 尼古丁最高=("尼古丁最高浓度", "max"),
            标价CHF=("价格_CHF", "first")).reset_index().rename(
            columns={"网站名称": "网站", "设备类别_显示": "产品形态", "容量_最大ml": "标注容量(mL)"})
        st.markdown(f"""
        <div class="danger-box">
        共 <b>{len(clue)}</b> 条记录（{pct(len(clue), len(fdf))}）的标注容量超过对应上限：一次性 &gt; {DISPOSABLE_LIMIT_ML} mL 或补充装 &gt; {REFILL_LIMIT_ML} mL
        <br>
        <span style="color:#6b7280">解读须谨慎：① 这些产品可能被站点列在“一次性”分类下，但实际属于可充电 / 可补充设备或预填充 Pod 系统，法规归类和适用上限不同；
        ② 50 mL 的两款为 0 mg 无尼古丁烟油；
        ③ 以上不构成对商家合规性的结论，需回到详情页核实产品类型。</span>
        </div>
        """, unsafe_allow_html=True)
        show_df(grp)

    st.markdown("---")
    st.markdown('<p class="section-title">🔋 一次性电子烟口数分档（puffs口数字段：0.6k=600，2k4=2,400，10k=10,000）</p>', unsafe_allow_html=True)
    dpp = fdf[fdf["设备类别_显示"] == FORM_DP].copy()
    if dpp.empty:
        st.info("当前筛选下暂无一次性电子烟记录。")
    else:
        dpp["口数档位"] = puff_tier(dpp["口数_数值"]).astype(object).where(dpp["口数_数值"].notna(), "未标注")
        pc1, pc2 = st.columns(2)
        with pc1:
            ppc = vc_df(dpp["口数档位"], "口数档位", "条数")
            order = PUFF_LABELS + ["未标注"]
            ppc["_o"] = ppc["口数档位"].apply(lambda x: order.index(x) if x in order else 99)
            ppc = ppc.sort_values("_o")
            fig = px.bar(ppc, x="口数档位", y="条数", text="条数", color_discrete_sequence=[COLORS["primary"]], template=T)
            fig.update_traces(textposition="outside")
            fig.update_layout(margin=dict(t=10, b=10), xaxis_title="口数（puffs）", yaxis_title="条数")
            show_chart(fig)
        with pc2:
            bt = dpp.groupby(["品牌", "口数档位"]).size().reset_index(name="条数")
            brand_order = dpp["品牌"].value_counts().index.tolist()
            fig_bt = px.bar(bt, x="品牌", y="条数", color="口数档位", barmode="stack", template=T,
                            category_orders={"品牌": brand_order, "口数档位": PUFF_LABELS + ["未标注"]},
                            color_discrete_sequence=["#fbdcd9", "#f5a39b", "#e8695b", COLORS["primary"], "#7f1d1d", "#94a3b8"])
            fig_bt.update_layout(xaxis_title="", yaxis_title="条数", legend_title_text="口数", margin=dict(t=10, b=10))
            show_chart(fig_bt)
        pv = dpp["口数_数值"].dropna()
        if len(pv):
            n600 = int((pv <= 600).sum())
            st.caption(f"一次性记录中可解析口数的 {len(pv)} 条：中位数约 {int(pv.median()):,} 口；≤600 口 {n600} 条（{n600 / len(pv):.0%}）。")

# ============================================================
# TAB 5  价格分析
# ============================================================
with tabs[5]:
    st.markdown("## 💰 价格分析（CHF）")
    excl_nf = False
    nf_n = int(fdf["无口味"].sum())
    if nf_n:
        excl_nf = st.checkbox(f"剔除无口味产品（Nicotine Shot / 基液 / Flavourless，当前 {nf_n} 条）", value=True, key="excl_nf",
                              help="尼古丁 Shot 等无口味产品价格明显低于调味烟油，混入会拉低烟油均价；原始数据未改动。")
    pdf_ = fdf.dropna(subset=["价格_CHF"]).copy()
    if excl_nf:
        pdf_ = pdf_[~pdf_["无口味"]]

    if pdf_.empty:
        st.info("暂无价格数据。")
    else:
        dev_types = pdf_["设备类别_显示"].unique()
        price_cols = st.columns(len(dev_types))
        for col, dev in zip(price_cols, dev_types):
            sub = pdf_[pdf_["设备类别_显示"] == dev]["价格_CHF"]
            col.markdown(
                f'<div class="metric-card"><div class="value">CHF {sub.mean():,.2f}</div>'
                f'<div class="label">{dev} 均价（中位数 {sub.median():,.2f}）<br>'
                f'区间 CHF {sub.min():,.2f}—{sub.max():,.2f}<br>≈¥{sub.mean() * CHF_CNY:,.1f}</div></div>',
                unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        col_l, col_r = st.columns(2)
        with col_l:
            st.markdown('<p class="section-title">各产品形态价格分布</p>', unsafe_allow_html=True)
            fig_pb = px.box(pdf_, x="设备类别_显示", y="价格_CHF", color="设备类别_显示", points="all",
                            color_discrete_map=FORM_COLOR_MAP, template=T, hover_data=["品牌", "网站名称"])
            fig_pb.update_layout(showlegend=False, xaxis_title="", yaxis_title="标价（CHF）", margin=dict(t=10, b=10))
            show_chart(fig_pb)
            st.caption("箱线图含全部散点；一次性电子烟价格受口数、油量影响很大，比较时请结合「产品规格」页。")

        with col_r:
            st.markdown('<p class="section-title">各品牌来源国均价对比</p>', unsafe_allow_html=True)
            cp = pdf_.groupby("品牌来源国")["价格_CHF"].agg(["mean", "count"]).reset_index()
            cp.columns = ["品牌来源国", "均价", "记录数"]
            cp = cp.sort_values("均价", ascending=False)
            cp["display"] = cp.apply(lambda r: f"CHF {r['均价']:,.2f}（¥{r['均价'] * CHF_CNY:,.0f}）· n={int(r['记录数'])}", axis=1)
            fig_cp = px.bar(cp, x="品牌来源国", y="均价", text="display", color="品牌来源国",
                            color_discrete_map=COUNTRY_COLOR_MAP, template=T)
            fig_cp.update_traces(texttemplate="%{text}", textposition="outside")
            fig_cp.update_layout(xaxis_title="", yaxis_title="均价（CHF）", showlegend=False, margin=dict(t=10, b=10))
            show_chart(fig_cp)
            st.caption("来源国均价混合了烟油与一次性两种形态，结构不同会影响对比；n 为样本量，样本小的国家仅供参考。")

        st.markdown('<p class="section-title">各网站均价对比（分产品形态）</p>', unsafe_allow_html=True)
        sp = pdf_.groupby(["网站名称", "设备类别_显示"])["价格_CHF"].mean().reset_index()
        fig_sp = px.bar(sp, x="价格_CHF", y="网站名称", color="设备类别_显示", orientation="h", barmode="group",
                        color_discrete_map=FORM_COLOR_MAP, template=T, text=sp["价格_CHF"].round(2))
        fig_sp.update_traces(textposition="outside")
        fig_sp.update_layout(xaxis_title="均价（CHF）", yaxis_title="", legend_title_text="", margin=dict(t=10, b=10))
        show_chart(fig_sp)

        cu1, cu2 = st.columns(2)
        with cu1:
            st.markdown('<p class="section-title">瓶装烟油单价（CHF/mL，单一规格）</p>', unsafe_allow_html=True)
            up = pdf_[pdf_["单价_CHF每ml"].notna()]
            if up.empty:
                st.info("暂无可计算单价的烟油记录。")
            else:
                ups = up.groupby("网站名称")["单价_CHF每ml"].mean().reset_index().sort_values("单价_CHF每ml")
                fig_up = px.bar(ups, x="单价_CHF每ml", y="网站名称", orientation="h", text=ups["单价_CHF每ml"].round(2),
                                color_discrete_sequence=[COLORS["secondary"]], template=T)
                fig_up.update_traces(textposition="outside")
                fig_up.update_layout(xaxis_title="平均单价（CHF/mL）", yaxis_title="", margin=dict(t=10, b=10))
                show_chart(fig_up)
                st.caption("按“标价 ÷ 容量”计算；50 mL 的 0 mg 烟油单价明显更低，已包含在平均值内。")
        with cu2:
            st.markdown('<p class="section-title">一次性：各品牌每百口价格（CHF/100 puffs）</p>', unsafe_allow_html=True)
            pp100 = pdf_[pdf_["每百口价格_CHF"].notna()]
            if pp100.empty:
                st.info("暂无可计算的一次性每百口价格。")
            else:
                pb = pp100.groupby("品牌").agg(每百口价格=("每百口价格_CHF", "mean"), n=("排名", "size")).reset_index().sort_values("每百口价格")
                pb["display"] = pb.apply(lambda r: f"{r['每百口价格']:.2f} · n={int(r['n'])}", axis=1)
                fig_pb2 = px.bar(pb, x="每百口价格", y="品牌", orientation="h", text="display",
                                 color_discrete_sequence=[COLORS["primary"]], template=T)
                fig_pb2.update_traces(textposition="outside")
                fig_pb2.update_layout(xaxis_title="平均每百口价格（CHF）", yaxis_title="", margin=dict(t=10, b=10))
                show_chart(fig_pb2)
                st.caption("高口数产品（如 10K）每口价格更低，但整机标价更高，且大容量产品的消费税也更高（见「法规与税收」页）。")

# ============================================================
# TAB 6  法规与税收
# ============================================================
with tabs[6]:
    st.markdown("## ⚖️ 国家监管框架与消费税政策")
    st.markdown('<p class="section-title">瑞士现行监管框架</p>', unsafe_allow_html=True)
    regs = [
        ("销售年龄",
         "全国统一 <b>18 周岁</b>，自 <b>2024-10-01</b> 起；销售点须张贴禁售提示；自动售货机仅在未成年人无法取得时方可销售。",
         "reg-card-open"),
        ("尼古丁浓度上限",
         "含尼古丁液体 ≤ <b>20 mg/mL</b>。",
         "reg-card"),
        ("容量限制",
         "补充装容器 ≤ <b>10 mL</b>；一次性电子烟及一次性烟弹 ≤ <b>2 mL</b>。",
         "reg-card-restrict"),
        ("广告与促销（现行，2024-10-01 起）",
         "禁止针对未成年人的广告；禁止在公共场所可见的广告牌、影院、公共交通、公共建筑、运动场所及赛事投放广告，以及带比价或赠品承诺的广告；"
         "销售点内广告目前仍允许；禁止免费派发及赠品、奖品式促销；禁止具国际性质、面向未成年人或由公共机构举办活动的赞助；"
         "广告须附警示语；各州可制定更严格规定。仅面向烟草行业从业人员的广告与促销不受上述限制。",
         "reg-card-restrict"),
        ("广告与促销（即将生效）",
         "议会已于 <b>2025-06-20</b> 通过修订（落实“无烟草广告”人民倡议），将禁止销售点广告、公共可达场所广告、网络与社交媒体广告（有年龄核验系统的网站除外）"
         "及多数印刷媒体广告；BAG 预计 <b>2027 年初</b>生效，TabPV 修订稿咨询期已于 2026-03-20 截止。",
         "reg-card-ban"),
        ("一次性电子烟禁令（立法进程）",
         "联邦层面的一次性电子烟禁令动议已获议会通过，政府尚未提交草案，正评估以条例或修订 TabPG 的方式落实；瓦莱州、日内瓦州已先行实施禁令，其他多个州已有动议或草案。"
         "产品结构上（调研判断）：不可充电的一次性产品面临最大不确定性；可充电的预填充 Pod / 可补充系统被动议排除在禁令之外，相对更有确定性。",
         "reg-card-ban"),
        ("健康警示与包装",
         "包装须载明：品名、瑞士境内制造商/进口商名称或 BAZG 分配的 Revers 号、产地、警示语；含尼古丁液体的每个容器须标注 mg/mL。",
         "reg-card"),
    ]
    for topic, detail, cls in regs:
        st.markdown(f'<div class="reg-card {cls}"><b>{topic}</b><br>{detail}</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown('<p class="section-title">💰 消费税：按毫升计税，一次性与可重复使用系统差异大</p>', unsafe_allow_html=True)
    tax_rule = pd.DataFrame({
        "产品类别": ["一次性电子烟", "可重复使用系统：含尼古丁的液体 / 补充装 / 烟弹", "可重复使用系统：不含尼古丁的补充液"],
        "税率": [f"CHF {TAX_DISPOSABLE_CHF_ML:.2f}/mL（无论是否含尼古丁）", f"CHF {TAX_REFILL_NIC_CHF_ML:.2f}/mL", "不征税"],
    })
    show_df(tax_rule)

# ============================================================
# TAB 7  总结洞察
# ============================================================
with tabs[7]:
    st.markdown("## 📝 市场总结与关键洞察")

    fl_sum = fdf[~fdf["无口味"]]
    flavor_top_s = top_n_with_ties(fl_sum["口味名_显示"].value_counts(), 3) if len(fl_sum) else pd.Series(dtype=int)
    flavor_top = flavor_top_s.index.tolist()
    _tag_s = pd.Series(dict(Counter([t for tl in fdf["口味标签列表"] for t in tl]).most_common()), dtype=int)
    tags_top3 = [f"{t}（{c}）" for t, c in top_n_with_ties(_tag_s, 3).items()] if len(_tag_s) else []
    vcat = fdf["分类"].value_counts()
    fruit_pct = (fdf["分类"] == "水果").mean()
    tob_pct = (fdf["分类"] == "烟草").mean()
    mint_pct = (fdf["分类"] == "薄荷").mean()
    candy_pct = (fdf["分类"] == "糖果").mean()
    ice_pct_total = (fdf["含冰/薄荷"] == "是").mean()
    voc2 = fdf["品牌来源国"].value_counts()
    ch_all = voc2.get("瑞士", 0) / len(fdf)
    cn_all = voc2.get("中国", 0) / len(fdf)
    local_all = (fdf["是否本地代理/本土品牌"] == "是").mean()
    brand_top = top_n_with_ties(fdf["品牌"].value_counts(), 3)
    cr3_all = brand_top.sum() / len(fdf)
    el_f = fdf[fdf["设备类别_显示"] == FORM_EL]
    dp_f = fdf[fdf["设备类别_显示"] == FORM_DP]
    pv_all = dp_f["口数_数值"].dropna()
    big_dp = int((dp_f["容量_最大ml"] > DISPOSABLE_LIMIT_ML).sum()) if len(dp_f) else 0
    big_sites = "、".join(sorted(dp_f[dp_f["容量_最大ml"] > DISPOSABLE_LIMIT_ML]["网站名称"].unique())) if big_dp else "—"
    at20_pct = (fdf["尼古丁最高浓度"] == NIC_LIMIT_MG_ML).mean()

    st.markdown('<div class="summary-card">', unsafe_allow_html=True)
    st.markdown("### 🌍 1. 市场背景：成长期、一次性占比最高、监管快速收紧")
    st.markdown("- EGL 2025：4.1% 的人口至少每月使用电子烟，使用者约 30–32 万；EGL 2024：54% 的使用者主要使用一次性电子烟。")
    st.markdown("- 监管：18 岁统一销售年龄（2024-10-01）；含尼古丁液体 ≤ 20 mg/mL、补充装 ≤ 10 mL、一次性 ≤ 2 mL；一次性消费税 CHF 1.00/mL，含尼古丁补充液 CHF 0.20/mL。")
    st.markdown("- 趋势：广告修订预计 2027 年初生效；联邦一次性电子烟禁令动议已获议会通过，政府正评估落实方式（条例或修订 TabPG）；瓦莱州、日内瓦州已先行实施禁令。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="summary-card">', unsafe_allow_html=True)
    st.markdown("### 🍓 2. 口味：水果为主，冰感常见")
    st.markdown(f"- 水果口味占 **{fruit_pct:.0%}**，烟草 **{tob_pct:.0%}**，薄荷 **{mint_pct:.0%}**，糖果 **{candy_pct:.0%}**；分类排序：{' > '.join(f'{k}({v / len(fdf):.0%})' for k, v in vcat.items())}。")
    if flavor_top:
        st.markdown(f"- 出现最多的口味名：**{' / '.join(f'{k}（{v}）' for k, v in flavor_top_s.items())}**；最高频口味元素：**{' / '.join(tags_top3)}**"
                    + ("（第三名并列的一并列出）。" if (len(flavor_top_s) > 3 or len(tags_top3) > 3) else "。"))
    st.markdown(f"- 明确含冰/薄荷的记录占 **{ice_pct_total:.0%}**。")
    st.markdown("- 与官方调查方向基本一致（水果味最常用，薄荷 21%、糖果 15%），但官方为“使用者选择”的多选占比，本看板为“热销榜记录占比”，口径不同，不宜直接对比数值。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="summary-card">', unsafe_allow_html=True)
    st.markdown(f"### 🏷️ 3. 品牌：{brand_top.index[0]} 占 {brand_top.iloc[0] / len(fdf):.0%}，来源国以{voc2.index[0]}为主（{voc2.iloc[0] / len(fdf):.0%}）")
    st.markdown(f"- 品牌来源国：{'、'.join(f'{k} {v / len(fdf):.0%}' for k, v in voc2.items())}；中国 **{cn_all:.0%}**，瑞士本土 **{ch_all:.0%}**。")
    brand_top_txt = "、".join(f"{k}（{v}）" for k, v in brand_top.items())
    tie_tag = "（含并列）" if len(brand_top) > 3 else ""
    st.markdown(f"- 出现最多的品牌：{brand_top_txt}；前三名{tie_tag}合计 **{cr3_all:.0%}**。")
    st.markdown(f"- 瑞士本土品牌/本地代理记录占 **{local_all:.0%}**，集中在店铺自有品牌（尼古丁 Shot / 基液）和少数本土一次性品牌。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="summary-card">', unsafe_allow_html=True)
    st.markdown("### 🔋 4. 一次性电子烟：口数与容量结构")
    if not len(dp_f):
        st.markdown("- 当前筛选下没有一次性电子烟记录。")
    if len(pv_all):
        st.markdown(f"- 一次性记录中可解析口数 {len(pv_all)} 条，中位数约 **{int(pv_all.median()):,}** 口；≤600 口占 **{(pv_all <= 600).mean():.0%}**，最高 {int(pv_all.max()):,} 口。")
    st.markdown("- 一次性数据来自网站所定义的一次性热卖榜单，结论以品牌/口味/口数为主。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="action-box">', unsafe_allow_html=True)
    st.markdown("### 📢 5. 对产品与市场进入的参考（基于样本的描述性建议）")
    st.markdown(f"- **合规先行**：当前样本中最高浓度档恰为 20 mg/mL 上限的记录占 {at20_pct:.0%}；补充装容器 ≤ 10 mL、一次性 ≤ 2 mL；大容量 / 高口数产品的归类与禁令风险最大，预填充 Pod / 可补充系统相对更有确定性。")
    st.markdown("- **口味结构**：以水果为主、搭配冰感/薄荷，并保留烟草口味；同一口味常见 5 / 10 / 20 mg 多浓度并行（尤其尼古丁盐烟油）。")
    st.markdown("- **营销**：现行已禁止公共场所可见广告与赠品促销，2027 年初预计进一步覆盖销售点与网络/社交媒体广告，建议优先布局合规渠道与 B2B 展会（具体边界以主管部门文件为准）。")
    st.markdown('</div>', unsafe_allow_html=True)
