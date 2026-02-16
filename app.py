import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np

# ═══════════════════════════════════════════════════════════════════════════
# Page Configuration
# ═══════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Healthier Humanity · Executive Dashboard",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ═══════════════════════════════════════════════════════════════════════════
# Styling
# ═══════════════════════════════════════════════════════════════════════════

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    * { font-family: 'Inter', sans-serif; }
    .block-container { padding: 1rem 2rem 2rem 2rem; max-width: 1400px; }
    h1 { font-weight: 700 !important; letter-spacing: -0.5px; }
    h2 { font-weight: 600 !important; font-size: 1.3rem !important; color: #1e293b; margin-top: 0.8rem !important; }
    h3 { font-weight: 600 !important; font-size: 1.05rem !important; color: #334155; }

    /* KPI cards */
    [data-testid="metric-container"] {
        background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 20px 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04), 0 4px 12px rgba(0,0,0,0.03);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    [data-testid="metric-container"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 16px rgba(0,0,0,0.08);
    }
    [data-testid="stMetricLabel"] p { font-size: 0.82rem; font-weight: 500; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; }
    [data-testid="stMetricValue"] { font-size: 1.8rem !important; font-weight: 700 !important; color: #0f172a !important; }
    [data-testid="stMetricDelta"] { font-size: 0.78rem; }

    /* Section separators */
    .section-divider { border: none; border-top: 1px solid #e2e8f0; margin: 1.5rem 0 1rem 0; }

    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# Color Palette
# ═══════════════════════════════════════════════════════════════════════════

PALETTE = {
    "yt_long": "#FF0000",
    "yt_shorts": "#FF8A8A",
    "audio_fr": "#7C3AED",
    "audio_en": "#C4B5FD",
    "tiktok": "#00C9C8",
    "instagram": "#E1306C",
    "blog": "#F59E0B",
    "linkedin": "#0A66C2",
    "s0": "#94a3b8",
    "s1": "#3B82F6",
    "s2": "#10B981",
    "s3": "#F59E0B",
    "s4": "#EF4444",
}
SEASON_COLORS_LIST = [PALETTE["s0"], PALETTE["s1"], PALETTE["s2"], PALETTE["s3"], PALETTE["s4"],
                      "#8B5CF6", "#06B6D4", "#F97316", "#84CC16", "#EC4899"]


def get_season_color(s: int) -> str:
    return SEASON_COLORS_LIST[s] if s < len(SEASON_COLORS_LIST) else "#94a3b8"


def get_season_label(s: int) -> str:
    return "Teaser" if s == 0 else f"Season {s}"
CHANNEL_COLORS = {
    "YouTube": PALETTE["yt_long"],
    "Audio": PALETTE["audio_fr"],
    "TikTok": PALETTE["tiktok"],
    "Instagram": PALETTE["instagram"],
    "Blog": PALETTE["blog"],
}

PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", size=12, color="#334155"),
    margin=dict(l=50, r=20, t=40, b=50),
    hoverlabel=dict(bgcolor="white", font_size=12, font_family="Inter"),
)

# ═══════════════════════════════════════════════════════════════════════════
# Data Loading
# ═══════════════════════════════════════════════════════════════════════════

EXPECTED_SHEETS = ["Config", "Episodes", "Summary_By_Season", "Summary_By_Episode",
                    "YouTube_Long", "Audio"]


def fmt(n: float) -> str:
    """Human-friendly number formatting."""
    if n >= 1_000_000:
        return f"{n / 1_000_000:.1f}M"
    if n >= 10_000:
        return f"{n / 1_000:.1f}K"
    if n >= 1_000:
        return f"{n / 1_000:.1f}K"
    return f"{n:,.0f}"


@st.cache_data
def load_all(file_bytes: bytes):
    """Load every sheet into a dict of DataFrames from raw file bytes."""
    from io import BytesIO
    xls = pd.ExcelFile(BytesIO(file_bytes))
    data = {}

    # Config
    cfg = pd.read_excel(xls, "Config")
    data["snapshot_date"] = str(cfg.loc[cfg["Key"] == "SnapshotDate", "Value"].iloc[0])[:10]

    # Episodes master list
    ep = pd.read_excel(xls, "Episodes")
    ep.columns = ["season", "ep_num", "guest"]
    ep["season"] = ep["season"].astype(int)
    ep["ep_num"] = ep["ep_num"].astype(int)
    ep["label"] = ep.apply(
        lambda r: "Teaser" if r["ep_num"] == 0 else f"Ep {int(r['ep_num'])} – {r['guest']}", axis=1
    )
    ep["short_label"] = ep.apply(
        lambda r: "Teaser" if r["ep_num"] == 0 else f"Ep {int(r['ep_num'])}", axis=1
    )
    ep["season_label"] = ep["season"].map(get_season_label)
    data["episodes"] = ep

    # Summary by episode
    se = pd.read_excel(xls, "Summary_By_Episode")
    se = se.iloc[:len(ep)]  # trim empty rows
    se.columns = [
        "season", "ep_num", "guest",
        "yt_long_views", "yt_long_imp", "yt_long_ctr",
        "n_shorts", "yt_shorts_views", "yt_shorts_imp",
        "yt_total_views", "audio_total",
        "tt_posts", "tt_views",
        "ig_posts", "ig_views",
        "blog_views",
        "li_posts", "li_imp", "li_likes", "li_comments", "li_shares",
    ]
    se["season"] = se["season"].astype(int)
    se["ep_num"] = se["ep_num"].astype(int)
    se = se.fillna(0)
    for c in se.columns[3:]:
        se[c] = pd.to_numeric(se[c], errors="coerce").fillna(0)
    se["total_reach"] = se["yt_total_views"] + se["audio_total"] + se["tt_views"] + se["ig_views"] + se["blog_views"]
    se["label"] = ep["label"].values
    se["short_label"] = ep["short_label"].values
    se["season_label"] = ep["season_label"].values
    data["summary_ep"] = se

    # Summary by season
    ss = pd.read_excel(xls, "Summary_By_Season")
    ss.columns = [
        "season", "yt_long_views", "yt_long_imp",
        "yt_shorts_views", "yt_shorts_imp", "yt_total_views",
        "audio_total", "tt_views", "ig_views", "blog_views",
        "li_imp", "li_likes", "li_comments", "li_shares",
    ]
    ss["season"] = ss["season"].astype(int)
    ss["season_label"] = ss["season"].map(get_season_label)
    ss["total_reach"] = ss["yt_total_views"] + ss["audio_total"] + ss["tt_views"] + ss["ig_views"] + ss["blog_views"]
    data["summary_season"] = ss

    # YouTube Long
    yt = pd.read_excel(xls, "YouTube_Long")
    yt.columns = ["season", "ep_num", "guest", "url", "pub_date", "title", "views", "impressions", "ctr"]
    yt["season"] = yt["season"].astype(int)
    yt["ep_num"] = yt["ep_num"].astype(int)
    yt["label"] = ep["label"].values
    data["yt_long"] = yt

    # Audio
    au = pd.read_excel(xls, "Audio")
    au.columns = ["season", "ep_num", "guest", "dl_fr", "dl_en", "dl_total"]
    au["season"] = au["season"].astype(int)
    au["ep_num"] = au["ep_num"].astype(int)
    au["label"] = ep["label"].values
    au["short_label"] = ep["short_label"].values
    data["audio"] = au

    return data


# ═══════════════════════════════════════════════════════════════════════════
# Main Dashboard
# ═══════════════════════════════════════════════════════════════════════════

def main():
    # ── File Upload ─────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="text-align:center; padding:0 0 0.2rem 0">
            <h1 style="margin-bottom:0; font-size:2.2rem; background: linear-gradient(135deg, #0f172a, #3b82f6);
                -webkit-background-clip:text; -webkit-text-fill-color:transparent;">
                🎙️ Healthier Humanity Podcast
            </h1>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Upload a performance tracking Excel file to refresh the dashboard",
        type=["xlsx"],
        help="Upload a new snapshot (.xlsx) following the same format. The dashboard updates instantly.",
    )

    if uploaded is None:
        st.markdown(
            """
            <div style="text-align:center; padding:3rem 1rem; color:#94a3b8">
                <p style="font-size:3rem; margin-bottom:0.5rem">📊</p>
                <p style="font-size:1.1rem; font-weight:500; color:#64748b">Upload your performance tracking Excel to generate the dashboard</p>
                <p style="font-size:0.85rem">Drag & drop or click "Browse files" above</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.stop()

    file_bytes = uploaded.getvalue()

    data = load_all(file_bytes)
    se = data["summary_ep"]
    ss = data["summary_season"]
    yt = data["yt_long"]
    au = data["audio"]
    snapshot = data["snapshot_date"]

    # ── Header subtitle ─────────────────────────────────────────────────
    st.markdown(
        f"""
        <div style="text-align:center; padding:0 0 0.8rem 0">
            <p style="font-size:1rem; color:#64748b; margin-top:0">
                Executive Performance Dashboard · Data as of {snapshot}
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ═══════════════════════════════════════════════════════════════════
    # SECTION 1: KPI Cards
    # ═══════════════════════════════════════════════════════════════════
    total_reach = se["total_reach"].sum()
    yt_views = se["yt_total_views"].sum()
    audio_dl = se["audio_total"].sum()
    blog_v = se["blog_views"].sum()
    tt_v = se["tt_views"].sum()
    ig_v = se["ig_views"].sum()
    n_episodes = len(se[se["ep_num"] > 0])

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Reach", fmt(total_reach), help="Sum of views/downloads across all platforms")
    c2.metric("YouTube Views", fmt(yt_views), help="Long-form + Shorts")
    c3.metric("Audio Downloads", fmt(audio_dl))
    c4.metric("TikTok Views", fmt(tt_v))
    c5.metric("Instagram Views", fmt(ig_v))
    c6.metric("Episodes Published", str(n_episodes))

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # SECTION 2: Growth Trajectory
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("## 📈 Growth Trajectory")

    ep_data = se[se["ep_num"] > 0].copy().reset_index(drop=True)

    # Cumulative reach
    ep_data["cum_reach"] = ep_data["total_reach"].cumsum()
    ep_data["cum_yt"] = ep_data["yt_total_views"].cumsum()
    ep_data["cum_audio"] = ep_data["audio_total"].cumsum()
    ep_data["cum_social"] = (ep_data["tt_views"] + ep_data["ig_views"]).cumsum()

    col_a, col_b = st.columns([3, 2])

    with col_a:
        st.markdown("### Cumulative Total Reach")
        fig_cum = go.Figure()
        fig_cum.add_trace(go.Scatter(
            x=ep_data["short_label"], y=ep_data["cum_yt"],
            name="YouTube", fill="tozeroy", mode="lines",
            line=dict(color=PALETTE["yt_long"], width=0),
            fillcolor="rgba(255,0,0,0.15)",
        ))
        fig_cum.add_trace(go.Scatter(
            x=ep_data["short_label"], y=ep_data["cum_yt"] + ep_data["cum_audio"],
            name="Audio", fill="tonexty", mode="lines",
            line=dict(color=PALETTE["audio_fr"], width=0),
            fillcolor="rgba(124,58,237,0.15)",
        ))
        fig_cum.add_trace(go.Scatter(
            x=ep_data["short_label"], y=ep_data["cum_reach"],
            name="Total (incl. Social + Blog)", fill="tonexty", mode="lines",
            line=dict(color=PALETTE["tiktok"], width=0),
            fillcolor="rgba(0,201,200,0.12)",
        ))
        # Add line on top
        fig_cum.add_trace(go.Scatter(
            x=ep_data["short_label"], y=ep_data["cum_reach"],
            name="", showlegend=False, mode="lines",
            line=dict(color="#0f172a", width=2.5),
        ))
        fig_cum.update_layout(
            **PLOTLY_LAYOUT, height=380,
            yaxis_title="Cumulative Views / Downloads",
            legend=dict(orientation="h", y=1.12, x=0.5, xanchor="center"),
            xaxis_tickangle=-45,
        )
        fig_cum.update_yaxes(gridcolor="#f1f5f9", zeroline=False)
        st.plotly_chart(fig_cum, use_container_width=True)

    with col_b:
        st.markdown("### Per-Episode Reach")
        fig_ep = go.Figure()
        colors = [get_season_color(s) for s in ep_data["season"]]
        fig_ep.add_trace(go.Bar(
            x=ep_data["short_label"], y=ep_data["total_reach"],
            marker_color=colors,
            text=[fmt(v) for v in ep_data["total_reach"]],
            textposition="outside", textfont=dict(size=9),
            hovertemplate="%{x}<br>Reach: %{y:,.0f}<extra></extra>",
        ))
        fig_ep.update_layout(
            **PLOTLY_LAYOUT, height=380, showlegend=False,
            yaxis_title="Total Reach",
            xaxis_tickangle=-45,
        )
        fig_ep.update_yaxes(gridcolor="#f1f5f9", zeroline=False)
        st.plotly_chart(fig_ep, use_container_width=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # SECTION 3: YouTube Deep Dive
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("## 🎬 YouTube Performance")
    col_yt1, col_yt2 = st.columns(2)

    with col_yt1:
        st.markdown("### Long-form vs Shorts Views")
        fig_yt = go.Figure()
        fig_yt.add_trace(go.Bar(
            x=ep_data["short_label"], y=ep_data["yt_long_views"],
            name="Long-form", marker_color=PALETTE["yt_long"],
            hovertemplate="Long-form: %{y:,.0f}<extra></extra>",
        ))
        fig_yt.add_trace(go.Bar(
            x=ep_data["short_label"], y=ep_data["yt_shorts_views"],
            name="Shorts", marker_color=PALETTE["yt_shorts"],
            hovertemplate="Shorts: %{y:,.0f}<extra></extra>",
        ))
        fig_yt.update_layout(
            **PLOTLY_LAYOUT, barmode="stack", height=400,
            yaxis_title="Views",
            legend=dict(orientation="h", y=1.1, x=0.5, xanchor="center"),
            xaxis_tickangle=-45,
        )
        fig_yt.update_yaxes(gridcolor="#f1f5f9", zeroline=False)
        st.plotly_chart(fig_yt, use_container_width=True)

    with col_yt2:
        st.markdown("### Click-Through Rate (CTR)")
        ctr_data = ep_data[ep_data["yt_long_ctr"] > 0].copy()
        ctr_colors = [get_season_color(s) for s in ctr_data["season"]]
        fig_ctr = make_subplots(specs=[[{"secondary_y": True}]])
        fig_ctr.add_trace(go.Bar(
            x=ctr_data["short_label"],
            y=ctr_data["yt_long_imp"],
            name="Impressions",
            marker_color=["rgba(59,130,246,0.18)"] * len(ctr_data),
            hovertemplate="Impressions: %{y:,.0f}<extra></extra>",
        ), secondary_y=False)
        fig_ctr.add_trace(go.Scatter(
            x=ctr_data["short_label"],
            y=ctr_data["yt_long_ctr"] * 100,
            name="CTR %",
            mode="lines+markers+text",
            text=[f"{v*100:.1f}%" for v in ctr_data["yt_long_ctr"]],
            textposition="top center", textfont=dict(size=9, color=PALETTE["yt_long"]),
            marker=dict(color=ctr_colors, size=10, line=dict(width=2, color="white")),
            line=dict(color="#ef4444", width=2),
            hovertemplate="CTR: %{y:.1f}%<extra></extra>",
        ), secondary_y=True)
        fig_ctr.update_layout(
            **PLOTLY_LAYOUT, height=400,
            legend=dict(orientation="h", y=1.1, x=0.5, xanchor="center"),
            xaxis_tickangle=-45,
        )
        fig_ctr.update_yaxes(title_text="Impressions", gridcolor="#f1f5f9", zeroline=False, secondary_y=False)
        fig_ctr.update_yaxes(title_text="CTR (%)", secondary_y=True, showgrid=False)
        st.plotly_chart(fig_ctr, use_container_width=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # SECTION 4: Season-over-Season Comparison
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("## 📊 Season-over-Season Comparison")

    # Filter out the Teaser (season 0) for season comparison
    ss_real = ss[ss["season"] > 0].copy()

    col_s1, col_s2 = st.columns(2)

    with col_s1:
        st.markdown("### Total Reach by Season")
        channels = ["yt_total_views", "audio_total", "tt_views", "ig_views", "blog_views"]
        ch_labels = ["YouTube", "Audio", "TikTok", "Instagram", "Blog"]
        ch_colors = [PALETTE["yt_long"], PALETTE["audio_fr"], PALETTE["tiktok"], PALETTE["instagram"], PALETTE["blog"]]

        fig_season = go.Figure()
        for ch, lbl, clr in zip(channels, ch_labels, ch_colors):
            fig_season.add_trace(go.Bar(
                x=ss_real["season_label"], y=ss_real[ch],
                name=lbl, marker_color=clr,
                hovertemplate=f"{lbl}: " + "%{y:,.0f}<extra></extra>",
            ))
        fig_season.update_layout(
            **PLOTLY_LAYOUT, barmode="stack", height=420,
            yaxis_title="Total Views / Downloads",
            legend=dict(orientation="h", y=1.12, x=0.5, xanchor="center"),
        )
        fig_season.update_yaxes(gridcolor="#f1f5f9", zeroline=False)
        st.plotly_chart(fig_season, use_container_width=True)

    with col_s2:
        st.markdown("### Season Growth Rates")
        growth_metrics = {
            "YouTube": "yt_total_views",
            "Audio": "audio_total",
            "TikTok": "tt_views",
            "Instagram": "ig_views",
            "Total Reach": "total_reach",
        }
        growth_rows = []
        for season_idx in range(1, len(ss_real)):
            curr = ss_real.iloc[season_idx]
            prev = ss_real.iloc[season_idx - 1]
            for metric_name, col in growth_metrics.items():
                prev_val = prev[col]
                curr_val = curr[col]
                if prev_val > 0:
                    pct = (curr_val - prev_val) / prev_val * 100
                else:
                    pct = 0
                growth_rows.append({
                    "transition": f"S{int(prev['season'])} → S{int(curr['season'])}",
                    "metric": metric_name,
                    "growth": pct,
                })
        df_growth = pd.DataFrame(growth_rows)
        if not df_growth.empty:
            fig_growth = go.Figure()
            for metric_name in growth_metrics.keys():
                subset = df_growth[df_growth["metric"] == metric_name]
                fig_growth.add_trace(go.Bar(
                    x=subset["transition"], y=subset["growth"],
                    name=metric_name,
                    marker_color=CHANNEL_COLORS.get(metric_name, "#64748b"),
                    text=[f"{v:+.0f}%" for v in subset["growth"]],
                    textposition="outside", textfont=dict(size=9),
                    hovertemplate=f"{metric_name}: " + "%{y:+.1f}%<extra></extra>",
                ))
            fig_growth.update_layout(
                **PLOTLY_LAYOUT, barmode="group", height=420,
                yaxis_title="Growth %",
                legend=dict(orientation="h", y=1.12, x=0.5, xanchor="center"),
            )
            fig_growth.update_yaxes(gridcolor="#f1f5f9", zeroline=True, zerolinecolor="#cbd5e1")
            fig_growth.add_hline(y=0, line_dash="dot", line_color="#94a3b8", line_width=1)
            st.plotly_chart(fig_growth, use_container_width=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # SECTION 5: Reach vs Deep Engagement
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("## 🎯 Reach vs Deep Engagement")
    st.caption(
        "**Deep listens** = YouTube long-form views + Audio downloads — people who actually engage with the full episode. "
        "**Viral reach** = Shorts + TikTok + Instagram — broad but shallow exposure."
    )

    ep_data["deep_listens"] = ep_data["yt_long_views"] + ep_data["audio_total"]
    ep_data["viral_reach"] = ep_data["yt_shorts_views"] + ep_data["tt_views"] + ep_data["ig_views"]
    ep_data["depth_ratio"] = (ep_data["deep_listens"] / ep_data["total_reach"].replace(0, 1) * 100).round(1)

    col_d1, col_d2 = st.columns(2)

    with col_d1:
        st.markdown("### Top 10 Episodes by Deep Listens")
        top_deep = ep_data.nlargest(10, "deep_listens").sort_values("deep_listens", ascending=True)
        fig_deep = go.Figure()
        fig_deep.add_trace(go.Bar(
            y=top_deep["label"], x=top_deep["yt_long_views"],
            name="YouTube Long-form", orientation="h",
            marker_color=PALETTE["yt_long"],
            hovertemplate="YT Long: %{x:,.0f}<extra></extra>",
        ))
        fig_deep.add_trace(go.Bar(
            y=top_deep["label"], x=top_deep["audio_total"],
            name="Audio Downloads", orientation="h",
            marker_color=PALETTE["audio_fr"],
            hovertemplate="Audio: %{x:,.0f}<extra></extra>",
        ))
        fig_deep.update_layout(
            **{k: v for k, v in PLOTLY_LAYOUT.items() if k != "margin"},
            barmode="stack", height=450,
            xaxis_title="Deep Listens",
            legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center"),
            margin=dict(l=200, r=20, t=40, b=50),
        )
        fig_deep.update_xaxes(gridcolor="#f1f5f9", zeroline=False)
        st.plotly_chart(fig_deep, use_container_width=True)

    with col_d2:
        st.markdown("### Deep Listens vs Viral Reach per Episode")
        fig_dv = go.Figure()
        fig_dv.add_trace(go.Bar(
            x=ep_data["short_label"], y=ep_data["deep_listens"],
            name="Deep Listens", marker_color="#1e40af",
            hovertemplate="Deep: %{y:,.0f}<extra></extra>",
        ))
        fig_dv.add_trace(go.Bar(
            x=ep_data["short_label"], y=ep_data["viral_reach"],
            name="Viral Reach", marker_color="#fbbf24",
            hovertemplate="Viral: %{y:,.0f}<extra></extra>",
        ))
        fig_dv.update_layout(
            **PLOTLY_LAYOUT, barmode="group", height=450,
            yaxis_title="Views / Downloads",
            legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center"),
            xaxis_tickangle=-45,
        )
        fig_dv.update_yaxes(gridcolor="#f1f5f9", zeroline=False)
        st.plotly_chart(fig_dv, use_container_width=True)

    # Engagement depth ratio chart
    st.markdown("### Engagement Depth Ratio")
    st.caption("What % of each episode's total reach comes from deep listens. Higher = more engaged audience.")
    sorted_depth = ep_data.sort_values("depth_ratio", ascending=True)
    bar_colors = ["#1e40af" if r >= 50 else "#f59e0b" if r >= 20 else "#ef4444" for r in sorted_depth["depth_ratio"]]
    fig_ratio = go.Figure(go.Bar(
        y=sorted_depth["label"], x=sorted_depth["depth_ratio"],
        orientation="h",
        marker_color=bar_colors,
        text=[f"{v:.0f}%" for v in sorted_depth["depth_ratio"]],
        textposition="outside", textfont=dict(size=10),
        hovertemplate="%{y}<br>Depth ratio: %{x:.1f}%<extra></extra>",
    ))
    fig_ratio.update_layout(
        **{k: v for k, v in PLOTLY_LAYOUT.items() if k != "margin"},
        height=max(350, len(sorted_depth) * 32 + 80),
        xaxis_title="Deep Listens as % of Total Reach",
        margin=dict(l=220, r=60, t=20, b=50),
    )
    fig_ratio.update_xaxes(gridcolor="#f1f5f9", zeroline=False, range=[0, min(105, sorted_depth["depth_ratio"].max() + 10)])
    fig_ratio.add_vline(x=50, line_dash="dot", line_color="#94a3b8", line_width=1,
                        annotation_text="50%", annotation_position="top")
    st.plotly_chart(fig_ratio, use_container_width=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # SECTION 6: Channel Mix & Top Episodes
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("## 🧩 Channel Mix & Top Content")
    col_m1, col_m2 = st.columns(2)

    with col_m1:
        st.markdown("### Channel Contribution to Total Reach")
        channel_totals = {
            "YouTube Long": se["yt_long_views"].sum(),
            "YouTube Shorts": se["yt_shorts_views"].sum(),
            "TikTok": se["tt_views"].sum(),
            "Instagram": se["ig_views"].sum(),
            "Audio": se["audio_total"].sum(),
            "Blog": se["blog_views"].sum(),
        }
        ch_df = pd.DataFrame({
            "channel": list(channel_totals.keys()),
            "value": list(channel_totals.values()),
        })
        ch_df = ch_df[ch_df["value"] > 0].sort_values("value", ascending=False)

        fig_mix = go.Figure(go.Pie(
            labels=ch_df["channel"],
            values=ch_df["value"],
            marker=dict(colors=[
                PALETTE["yt_long"], PALETTE["yt_shorts"], PALETTE["tiktok"],
                PALETTE["instagram"], PALETTE["audio_fr"], PALETTE["blog"],
            ]),
            hole=0.5,
            textinfo="label+percent",
            textposition="outside",
            textfont=dict(size=11),
            hovertemplate="%{label}<br>%{value:,.0f} (%{percent})<extra></extra>",
        ))
        fig_mix.update_layout(**PLOTLY_LAYOUT, height=420, showlegend=False)
        st.plotly_chart(fig_mix, use_container_width=True)

    with col_m2:
        st.markdown("### Top 10 Episodes by Total Reach")
        top = ep_data.nlargest(10, "total_reach").sort_values("total_reach", ascending=True)
        fig_top = go.Figure(go.Bar(
            y=top["label"], x=top["total_reach"],
            orientation="h",
            marker_color=[get_season_color(s) for s in top["season"]],
            text=[fmt(v) for v in top["total_reach"]],
            textposition="outside", textfont=dict(size=10),
            hovertemplate="%{y}<br>Reach: %{x:,.0f}<extra></extra>",
        ))
        fig_top.update_layout(
            **{k: v for k, v in PLOTLY_LAYOUT.items() if k != "margin"},
            height=420,
            xaxis_title="Total Reach",
            margin=dict(l=200, r=60, t=40, b=50),
        )
        fig_top.update_xaxes(gridcolor="#f1f5f9", zeroline=False)
        st.plotly_chart(fig_top, use_container_width=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # SECTION 6: Audio & Social Media
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("## 🎧 Audio & Social Media")
    col_au, col_soc = st.columns(2)

    with col_au:
        au_ep = au[au["ep_num"] > 0].copy()
        st.markdown("### Audio Downloads: French vs English")
        fig_audio = go.Figure()
        fig_audio.add_trace(go.Bar(
            x=au_ep["short_label"], y=au_ep["dl_fr"],
            name="French", marker_color=PALETTE["audio_fr"],
        ))
        fig_audio.add_trace(go.Bar(
            x=au_ep["short_label"], y=au_ep["dl_en"],
            name="English", marker_color=PALETTE["audio_en"],
        ))
        fig_audio.update_layout(
            **PLOTLY_LAYOUT, barmode="stack", height=400,
            yaxis_title="Downloads",
            legend=dict(orientation="h", y=1.1, x=0.5, xanchor="center"),
            xaxis_tickangle=-45,
        )
        fig_audio.update_yaxes(gridcolor="#f1f5f9", zeroline=False)
        st.plotly_chart(fig_audio, use_container_width=True)

    with col_soc:
        st.markdown("### TikTok & Instagram Views per Episode")
        fig_soc = go.Figure()
        fig_soc.add_trace(go.Bar(
            x=ep_data["short_label"], y=ep_data["tt_views"],
            name="TikTok", marker_color=PALETTE["tiktok"],
        ))
        fig_soc.add_trace(go.Bar(
            x=ep_data["short_label"], y=ep_data["ig_views"],
            name="Instagram", marker_color=PALETTE["instagram"],
        ))
        fig_soc.update_layout(
            **PLOTLY_LAYOUT, barmode="group", height=400,
            yaxis_title="Views",
            legend=dict(orientation="h", y=1.1, x=0.5, xanchor="center"),
            xaxis_tickangle=-45,
        )
        fig_soc.update_yaxes(gridcolor="#f1f5f9", zeroline=False)
        st.plotly_chart(fig_soc, use_container_width=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # SECTION 7: Cross-Platform Heatmap
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("## 🔥 Cross-Platform Performance Heatmap")
    hm_cols = ["yt_total_views", "audio_total", "tt_views", "ig_views", "blog_views"]
    hm_labels = ["YouTube", "Audio", "TikTok", "Instagram", "Blog"]
    raw = ep_data[hm_cols].values.astype(float)
    maxs = raw.max(axis=0)
    maxs[maxs == 0] = 1
    normed = raw / maxs
    text_labels = [[fmt(v) for v in row] for row in raw]

    fig_hm = go.Figure(go.Heatmap(
        z=normed,
        x=hm_labels,
        y=ep_data["label"],
        colorscale=[
            [0, "#f8fafc"], [0.15, "#dbeafe"], [0.3, "#93c5fd"],
            [0.5, "#3b82f6"], [0.7, "#1d4ed8"], [1.0, "#1e3a5f"],
        ],
        showscale=True,
        text=text_labels,
        texttemplate="%{text}",
        textfont=dict(size=10),
        hovertemplate="%{y}<br>%{x}: %{text}<extra></extra>",
        colorbar=dict(title="Relative<br>Perf.", thickness=15),
    ))
    fig_hm.update_layout(
        **{k: v for k, v in PLOTLY_LAYOUT.items() if k != "margin"},
        height=620,
        yaxis=dict(autorange="reversed"),
        margin=dict(l=220, r=20, t=40, b=50),
    )
    st.plotly_chart(fig_hm, use_container_width=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # SECTION 8: YouTube Content Ranking
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("## 🏆 YouTube Long-form Content Ranking")
    yt_sorted = yt.sort_values("views", ascending=False).reset_index(drop=True)
    yt_sorted["rank"] = range(1, len(yt_sorted) + 1)
    yt_sorted["ctr_pct"] = (yt_sorted["ctr"] * 100).round(1)

    display_yt = yt_sorted[["rank", "label", "title", "views", "impressions", "ctr_pct"]].copy()
    display_yt.columns = ["#", "Episode", "Title", "Views", "Impressions", "CTR %"]

    st.dataframe(
        display_yt.style.format({"Views": "{:,.0f}", "Impressions": "{:,.0f}", "CTR %": "{:.1f}%"}),
        use_container_width=True,
        height=500,
        hide_index=True,
    )

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════
    # SECTION 9: Full Data Table
    # ═══════════════════════════════════════════════════════════════════
    st.markdown("## 📋 Complete Episode Performance Data")

    seasons_filter = ["All Seasons"] + [get_season_label(s) for s in sorted(se["season"].unique()) if s > 0]
    sel = st.selectbox("Filter by season:", seasons_filter)
    if sel == "All Seasons":
        show = se[se["ep_num"] > 0].copy()
    else:
        # Reverse lookup: "Season 3" → 3
        label_to_num = {get_season_label(s): s for s in se["season"].unique()}
        s_num = label_to_num[sel]
        show = se[se["season"] == s_num].copy()

    col_rename = {
        "label": "Episode",
        "season_label": "Season",
        "yt_long_views": "YT Long Views",
        "yt_shorts_views": "YT Shorts Views",
        "yt_total_views": "YT Total",
        "yt_long_imp": "YT Impressions",
        "yt_long_ctr": "YT CTR",
        "audio_total": "Audio DL",
        "tt_views": "TikTok Views",
        "ig_views": "IG Views",
        "blog_views": "Blog Views",
        "total_reach": "Total Reach",
    }
    tbl = show[list(col_rename.keys())].rename(columns=col_rename)

    st.dataframe(
        tbl.style.format({
            "YT Long Views": "{:,.0f}",
            "YT Shorts Views": "{:,.0f}",
            "YT Total": "{:,.0f}",
            "YT Impressions": "{:,.0f}",
            "YT CTR": "{:.1%}",
            "Audio DL": "{:,.0f}",
            "TikTok Views": "{:,.0f}",
            "IG Views": "{:,.0f}",
            "Blog Views": "{:,.0f}",
            "Total Reach": "{:,.0f}",
        }),
        use_container_width=True,
        height=520,
        hide_index=True,
    )

    # ── Footer ──────────────────────────────────────────────────────────
    st.markdown(
        f"""
        <div style="text-align:center; padding:2rem 0 1rem 0; color:#94a3b8; font-size:0.8rem">
            Healthier Humanity Podcast · Executive Dashboard · Data snapshot: {snapshot}
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
