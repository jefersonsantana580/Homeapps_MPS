"""Pagina 4: historico de revisoes em HTML, sem servicos externos."""
import io
import json
import html
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from plotly.offline import get_plotlyjs

st.set_page_config(page_title="Histórico de volumes nas revisões", layout="wide")

# Logo existente no repositorio; nao altera a navegacao ou o painel HTML.
_raiz_pagina = Path(__file__).resolve().parent
_candidatos_logo = [
    _raiz_pagina.parent / "images" / "agco.jpg",
    _raiz_pagina / "images" / "agco.jpg",
    Path("images/agco.jpg"),
]
_logo_mps = next((p for p in _candidatos_logo if p.is_file()), None)
if _logo_mps is not None:
    st.logo(str(_logo_mps), size="large")
else:
    st.sidebar.warning("Logo não encontrado: images/agco.jpg")

ORDEM_CICLOS = ["0+0 Bgt", "0+12", "01+11", "02+10", "03+09", "04+08", "05+07", "06+06", "07+05", "08+04", "09+03", "10+02", "11+01", "12+0"]
ARQUIVO_EXCEL = "dados/base_volume_sites.xlsx"
ABA = "base"
CORES = {"DF":"#1F77B4", "MOM":"#6DC8A0", "RIG":"#6F4C9B", "TA":"#D62728", "PU":"#2CA02C", "CO":"#FF7F0E", "CO PKD":"#8C564B"}

def localizar_base():
    candidatos = [Path(ARQUIVO_EXCEL), Path(__file__).resolve().parent / ARQUIVO_EXCEL, Path(__file__).resolve().parent.parent / ARQUIVO_EXCEL]
    for caminho in candidatos:
        if caminho.is_file():
            return caminho
    raise FileNotFoundError(f"Base não encontrada: {ARQUIVO_EXCEL}")

@st.cache_data
def carregar_dados(caminho, modificacao):
    df = pd.read_excel(caminho, sheet_name=ABA, engine="openpyxl")
    df.columns = df.columns.astype(str).str.strip()
    obrigatorias = ["Tipo Base", "ANO", "BRAND", "PRODUCT MARKET", "SITE", "Product DR", "Nº CICLO", "Total"]
    faltantes = [c for c in obrigatorias if c not in df.columns]
    if faltantes:
        raise ValueError("Colunas ausentes: " + ", ".join(faltantes))
    for c in ["Tipo Base", "BRAND", "PRODUCT MARKET", "SITE", "Product DR", "Nº CICLO"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    for c in ["Tipo Base", "SITE", "Product DR"]:
        df[c] = df[c].str.upper()
    df["Nº CICLO"] = df["Nº CICLO"].replace({"0+0 BGT":"0+0 Bgt", "0+0 bgt":"0+0 Bgt"})
    df["ANO"] = df["ANO"].map(lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) and float(x).is_integer() else str(x).strip() if pd.notna(x) else "")
    df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0)
    df = df[df["Tipo Base"].eq("F_RESPONSE")].copy()
    return df[~(df["Product DR"].eq("PC") | (df["SITE"].eq("GENERAL RODRIGUEZ") & df["Product DR"].eq("CO")))].copy()

def consolidar(df, ciclos):
    # Sem fill_value: ausencia de registro nao deve virar volume zero.
    return df.pivot_table(index=["SITE", "Product DR"], columns="Nº CICLO", values="Total", aggfunc="sum").reindex(columns=ciclos).reset_index().sort_values(["SITE", "Product DR"])

def fmt(v, sinal=False):
    if pd.isna(v):
        return "—"
    return (f"{v:+,.0f}" if sinal else f"{v:,.0f}").replace(",", ".")

def excel(tabela, resumo):
    b = io.BytesIO()
    with pd.ExcelWriter(b, engine="openpyxl") as w:
        tabela.to_excel(w, sheet_name="Historico", index=False)
        resumo.to_excel(w, sheet_name="Comparativo", index=False)
        for ws in w.book.worksheets:
            ws.freeze_panes = "C2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = max(15, min(32, max(len(str(c.value or "")) for c in col) + 2))
    return b.getvalue()

def painel(tabela, resumo, ciclos, referencia, atual, contexto):
    totais = tabela[ciclos].sum(min_count=1)
    # KPI comparativo usa somente pares com registros nos dois ciclos.
    pares = resumo.dropna(subset=["Referência", "Atual"])
    base = pares["Referência"].sum() if len(pares) else float("nan")
    final = pares["Atual"].sum() if len(pares) else float("nan")
    delta = final - base
    percentual = f"{delta / base * 100:+.1f}%".replace(".", ",") if pd.notna(base) and base != 0 else "N/A"
    esc = lambda x: html.escape(str(x), quote=True)
    cards = [("Volume no ciclo selecionado", fmt(totais[atual]), atual), ("Referência comparável", fmt(base), referencia), ("Diferença comparável", fmt(delta, True), f"{atual} − {referencia}"), ("Variação comparável", percentual, f"{len(pares)} pares com registros nos dois ciclos")]
    cards_html = ''.join(f'<div class="card"><small>{esc(a)}</small><strong>{esc(b)}</strong><span>{esc(c)}</span></div>' for a,b,c in cards)
    fig = go.Figure()
    for produto in sorted(tabela["Product DR"].unique()):
        sub = tabela[tabela["Product DR"].eq(produto)]
        vals = sub[ciclos].sum(min_count=1)
        fig.add_trace(go.Scatter(x=ciclos, y=vals.tolist(), name=esc(produto), mode="lines+markers+text", connectgaps=False, text=[fmt(v) for v in vals], textposition="top center", line=dict(color=CORES.get(produto, "#64748b"), width=3), hovertemplate="%{x}<br>Volume: %{y:,.0f}<extra>%{fullData.name}</extra>"))
    fig.update_layout(template="plotly_white", height=340, margin=dict(l=30,r=30,t=35,b=35), legend=dict(orientation="h",y=1.16), yaxis_title="Volume", xaxis=dict(type="category"), font=dict(family="Arial",color="#334155"))
    grafico = fig.to_html(full_html=False, include_plotlyjs=False, config={"displaylogo":False,"responsive":True})
    cabecalho = '<th>Filial</th><th>Product DR</th>' + ''.join(f'<th>{esc(c)}</th>' for c in ciclos) + '<th>Diferença</th>'
    linhas=[]
    for _, row in tabela.iterrows():
        cells=[f'<td class="fixed">{esc(row["SITE"])}</td>',f'<td>{esc(row["Product DR"])}</td>']
        anterior=None
        for c in ciclos:
            v=row[c]
            classe=""
            if anterior is not None and pd.notna(v) and pd.notna(anterior):
                classe="up" if v>anterior else "down" if v<anterior else ""
            cells.append(f'<td class="{classe}">{fmt(v)}</td>')
            anterior=v
        d=row[atual]-row[referencia]
        cells.append(f'<td class="diff">{fmt(d,True)}</td>')
        linhas.append('<tr>'+''.join(cells)+'</tr>')
    ranking=resumo.dropna(subset=["Diferença"]).sort_values("Diferença",key=lambda s:s.abs(),ascending=False).head(6)
    lista=''.join(f'<div class="change"><span>{esc(r["SITE"])} · {esc(r["Product DR"])}</span><b>{fmt(r["Diferença"],True)}</b></div>' for _,r in ranking.iterrows()) or '<p>Sem pares comparáveis.</p>'
    css="""*{box-sizing:border-box}body{margin:0;background:#f3f5f9;color:#172b4d;font-family:Arial,sans-serif;padding:22px}.hero{background:linear-gradient(110deg,#142338,#263f5c);border-radius:16px;padding:25px;color:white}.hero h1{margin:5px 0 10px;font-size:26px}.hero p{color:#d0dbea;font-size:13px}.tag{font-size:11px;letter-spacing:2px;color:#a8c4e4}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin:18px 0}.card,.panel{background:white;border:1px solid #e1e7ef;border-radius:14px;padding:19px}.card small{color:#64748b;display:block}.card strong{display:block;font-size:30px;margin:12px 0}.card span{font-size:12px;color:#64748b}.grid{display:grid;grid-template-columns:3fr 1fr;gap:16px}.panel h2{font-size:17px;margin:0 0 12px}.change{display:flex;justify-content:space-between;gap:10px;padding:14px 0;border-bottom:1px solid #edf1f6;font-size:12px}.change b{white-space:nowrap}.tablepanel{margin-top:16px}.scroll{overflow:auto;max-height:480px}table{border-collapse:separate;border-spacing:0;width:100%;font-size:12px;white-space:nowrap}th{position:sticky;top:0;background:#e9eef5;text-align:right;padding:13px;z-index:2}td{padding:12px;text-align:right;border-bottom:1px solid #edf1f6}td:first-child,th:first-child{text-align:left}tr:nth-child(even){background:#f8fafc}.up{background:#e8f5ee;color:#166534}.down{background:#fff0ee;color:#a12727}.diff{font-weight:bold;background:#eef2ff}.note{font-size:12px;color:#64748b;line-height:1.6}input{border:1px solid #cbd5e1;border-radius:8px;padding:10px;width:280px;margin:0 0 14px}@media(max-width:900px){.cards{grid-template-columns:repeat(2,1fr)}.grid{grid-template-columns:1fr}}"""
    return '<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>'+css+'</style><script>'+get_plotlyjs()+'</script></head><body><div class="hero"><div class="tag">MPS · HISTÓRICO DE REVISÕES</div><h1>Evolução dos volumes</h1><p>'+esc(contexto)+'</p></div><div class="cards">'+cards_html+'</div><div class="grid"><section class="panel"><h2>Volumes por ciclo e produto</h2>'+grafico+'</section><section class="panel"><h2>Maiores mudanças em volume</h2><p class="note">'+esc(atual)+' − '+esc(referencia)+'</p>'+lista+'</section></div><section class="panel tablepanel"><h2>Histórico consolidado</h2><p class="note">Verde: aumento. Vermelho: redução em relação ao ciclo anterior exibido. Cores não indicam avaliação de desempenho. —: sem registro. Zero: volume registrado igual a zero.</p><input id="busca" placeholder="Buscar filial ou produto" aria-label="Buscar na tabela"><div class="scroll"><table><thead><tr>'+cabecalho+'</tr></thead><tbody>'+''.join(linhas)+'</tbody></table></div><p class="note">'+str(len(tabela))+' linhas. KPIs de diferença consideram apenas pares presentes nos dois ciclos. Ciclos são revisões do mesmo volume: não são somados entre si.</p></section><script>document.getElementById("busca").addEventListener("input",function(){const q=this.value.toLocaleLowerCase();document.querySelectorAll("tbody tr").forEach(r=>{r.style.display=r.textContent.toLocaleLowerCase().includes(q)?"":"none";});});</script></body></html>'

st.title("Histórico de volumes nas revisões")
st.caption("Selecione o recorte. O painel HTML acompanha os filtros, sem alterar a base original.")
try:
    caminho = localizar_base()
    df = carregar_dados(str(caminho), caminho.stat().st_mtime_ns)
except Exception as e:
    st.error(f"Erro ao carregar a base: {e}")
    st.stop()
if df.empty:
    st.warning("Sem dados F_RESPONSE após as exclusões.")
    st.stop()
cols=st.columns(5)
anos=sorted([a for a in df["ANO"].unique() if a])
if not anos:
    st.error("Nenhum ano válido na base.")
    st.stop()
ano=cols[0].selectbox("Ano",anos,index=len(anos)-1,key="hist_ano")
recorte=df[df["ANO"].eq(ano)].copy()
for container,col,label in zip(cols[1:],["SITE","Product DR","PRODUCT MARKET","BRAND"],["Filial","Product DR","Mercado","Marca"]):
    escolhas=container.multiselect(label,sorted(recorte[col].unique()),key="hist_"+col,placeholder="Todos")
    if escolhas:
        recorte=recorte[recorte[col].isin(escolhas)]
if recorte.empty:
    st.warning("Nenhum dado para os filtros selecionados.")
    st.stop()
desconhecidos=sorted(set(recorte["Nº CICLO"])-set(ORDEM_CICLOS))
if desconhecidos:
    st.warning("Ciclos fora da ordem configurada, não exibidos: "+", ".join(desconhecidos))
ciclos=[c for c in ORDEM_CICLOS if c in set(recorte["Nº CICLO"])]
if not ciclos:
    st.warning("Nenhum ciclo reconhecido.")
    st.stop()
a,b=st.columns(2)
referencia=a.selectbox("Ciclo de referência",ciclos,key="hist_ref")
atual=b.selectbox("Ciclo de comparação",ciclos,index=len(ciclos)-1,key="hist_atual")
st.caption("O ciclo de comparação padrão é o último com registros na ordem configurada, inclusive quando o volume é zero. Ajuste acima se necessário.")
tabela=consolidar(recorte,ciclos)
resumo=tabela[["SITE","Product DR"]].copy()
resumo["Referência"]=tabela[referencia]
resumo["Atual"]=tabela[atual]
resumo["Diferença"]=resumo["Atual"]-resumo["Referência"]
contexto=f"Ano {ano} | "+" | ".join(f"{c}: {', '.join(sorted(recorte[c].unique()))}" for c in ["SITE","Product DR","PRODUCT MARKET","BRAND"])
documento=painel(tabela,resumo,ciclos,referencia,atual,contexto)
if hasattr(st,"iframe"):
    st.iframe(documento,height=1100)
else:
    import streamlit.components.v1 as components
    components.html(documento,height=1100,scrolling=True)
st.caption("Downloads respeitam os filtros acima. A busca textual dentro do HTML é apenas visual.")
c1,c2=st.columns(2)
c1.download_button("Baixar Excel do recorte",excel(tabela,resumo),file_name=f"historico_volumes_{ano}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
c2.download_button("Baixar painel HTML",documento,file_name=f"historico_volumes_{ano}.html",mime="text/html")
