import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle
import numpy as np, os, sys
OUT = sys.argv[1]
os.makedirs(OUT, exist_ok=True)
NAVY="#1F2F4D"; GOLD="#B8975A"; GREY="#8C8279"; CREAM="#F4F0E8"; INK="#2B2B2B"
plt.rcParams.update({"font.family":"Carlito","font.size":10,"axes.edgecolor":"#9A9088"})

# 1. Hour distribution (official PPC values)
fig, ax = plt.subplots(figsize=(6.5,2.3), dpi=200)
segs=[(2080,NAVY,"Formação Geral Básica\n2.080 h (sem integradoras)"),(120,GOLD,"Integradoras\nEPT 120 h"),(800,GREY,"Itinerário técnico\n800 h")]
x=0
for v,c,l in segs:
    ax.barh(0,v,left=x,color=c,height=0.5,edgecolor="white")
    if v>500:
        ax.text(x+v/2,0,l,ha="center",va="center",color="white",fontsize=9,fontweight="bold",linespacing=1.2)
    else:
        ax.text(x+v/2,-0.42,"Integradoras EPT\n120 h (dentro da FGB)",ha="center",va="top",color=GOLD,fontsize=8.5,fontweight="bold",linespacing=1.2)
    x+=v
ax.plot([0,2200],[0.42,0.42],color=NAVY,lw=1); ax.plot([0,0],[0.37,0.42],color=NAVY,lw=1); ax.plot([2200,2200],[0.37,0.42],color=NAVY,lw=1)
ax.text(1100,0.48,"Formação Geral Básica = 2.200 h (inclui as 120 h integradoras)",ha="center",va="bottom",fontsize=9,color=NAVY)
ax.text(3000,-0.42,"Total: 3.000 h",ha="right",va="center",fontsize=10,color=INK,fontweight="bold")
ax.set_xlim(0,3000); ax.set_ylim(-0.6,0.8); ax.axis("off")
fig.text(0.01,0.01,"Fonte: PPC SED 00178627/2026. 2.080 h = 2.200 h − 120 h (cálculo derivado para evitar dupla contagem).",fontsize=7.5,color=GREY)
fig.tight_layout(); fig.savefig(f"{OUT}/fig1-carga-horaria.png"); plt.close(fig)

# 2. Three-year roadmap (from Anexo F1 and section 11)
fig, ax = plt.subplots(figsize=(6.5,3.6), dpi=200)
years=["2027","2028","2029"]
items={
 "2027":["T1 · Diagnóstico da comunicação\nde projeto escolar","T2 · Briefing para professor\nde outra área","T3 · Campanha integrada para\nmostra escolar"],
 "2028":["T1 · Pesquisa de problema\nde serviço","T2 · Campanha digital simulada\ncom professor-cliente","T3 · Revisão da campanha com\ndados simulados ou permitidos"],
 "2029":["T1 · Seleção de projeto de conclusão\ncom escopo seguro","T2 · Execução supervisionada\nou simulação validada","T3 · Banca de projeto técnico\ne mostra escolar"]}
gates={"2027":"Gate: piloto e 1º boletim\nà Direção","2028":"Gate: revisão semestral","2029":"Gate: avaliação final e\nrecomendação de continuidade"}
for i,y in enumerate(years):
    x0=i*2.1
    ax.add_patch(FancyBboxPatch((x0,0.3),1.95,3.2,boxstyle="round,pad=0.02,rounding_size=0.08",fc=CREAM,ec="#D9CFC0"))
    ax.add_patch(Rectangle((x0,3.25),1.95,0.5,fc=NAVY,ec="none"))
    ax.text(x0+0.975,3.5,f"Ano {i+1} · {y}",ha="center",va="center",color="white",fontsize=10.5,fontweight="bold")
    for j,t in enumerate(items[y]):
        yy=2.55-j*0.78
        ax.add_patch(Rectangle((x0+0.08,yy-0.27),0.05,0.54,fc=GOLD,ec="none"))
        ax.text(x0+0.2,yy,t,va="center",fontsize=7.6,color=INK,linespacing=1.15)
    ax.text(x0+0.975,0.55,gates[y],ha="center",va="center",fontsize=7.6,color=NAVY,style="italic")
ax.set_xlim(-0.05,6.25); ax.set_ylim(0,3.85); ax.axis("off")
fig.text(0.01,0.01,"Plano de implantação proposto (Anexo F1; seção 11). Cronograma planejado, não resultado observado.",fontsize=7.5,color=GREY)
fig.tight_layout(); fig.savefig(f"{OUT}/fig2-roadmap-3-anos.png"); plt.close(fig)

# 3. Competency targets heatmap (section 13.3) - PROPOSED TARGETS
comp=["C01 Briefing/escuta","C02 Fundamentos","C03 Pesquisa/dados","C04 Estratégia","C05 Design/criação","C06 Operação digital",
"C07 Indicadores","C08 IA crítica","C09 Ética/direitos","C10 Gestão de projeto","C11 Oralidade/equipe","C12 Aprendizagem/portfólio"]
vals=np.array([[2.0,3.0,3.5],[2.0,3.0,3.5],[1.5,2.8,3.5],[1.5,2.8,3.5],[2.0,3.0,3.5],[1.0,2.8,3.5],[1.0,2.8,3.5],[2.0,3.0,3.5],[2.0,3.0,3.8],[1.5,2.8,3.5],[2.0,3.0,3.5],[2.0,3.0,3.5]])
from matplotlib.colors import LinearSegmentedColormap
cmap=LinearSegmentedColormap.from_list("nv",[CREAM,"#D8C49A",GOLD,NAVY])
fig, ax = plt.subplots(figsize=(6.5,4.4), dpi=200)
im=ax.imshow(vals,cmap=cmap,vmin=0,vmax=4,aspect="auto")
for r in range(vals.shape[0]):
    for c in range(3):
        v=vals[r,c]; ax.text(c,r,f"{v:.1f}".replace(".",","),ha="center",va="center",fontsize=9,color="white" if v>2.6 else INK)
ax.set_xticks(range(3)); ax.set_xticklabels(["Fim do 1º ano (2027)","Fim do 2º ano (2028)","Fim do 3º ano (2029)"],fontsize=9)
ax.set_yticks(range(len(comp))); ax.set_yticklabels(comp,fontsize=8.5)
ax.tick_params(length=0); [s.set_visible(False) for s in ax.spines.values()]
cb=fig.colorbar(im,ax=ax,fraction=0.04,pad=0.02); cb.set_label("Nível alvo na rubrica (0–4)",fontsize=8); cb.ax.tick_params(labelsize=8)
ax.set_title("PROPOSED TARGETS / NOT OBSERVED RESULTS",loc="left",fontsize=10,color=NAVY,fontweight="bold")
fig.text(0.01,0.01,"Trajetórias-alvo propostas (seção 13.3) — não são resultados de avaliação nem dados observados.",fontsize=7.5,color=GREY)
fig.tight_layout(); fig.savefig(f"{OUT}/fig3-metas-competencias.png"); plt.close(fig)

# 4. Agency-school review gates G0-G8 (section 14.2)
steps=[("G0","Demanda"),("G1","Escuta"),("G2","Pesquisa"),("G3","Estratégia"),("G4","Produção"),("G5","Aprovação"),("G6","Execução"),("G7","Mensuração"),("G8","Retrospectiva")]
fig, ax = plt.subplots(figsize=(6.5,2.6), dpi=200)
for i,(g,l) in enumerate(steps):
    x=i*0.72
    col = GOLD if g=="G5" else NAVY
    ax.add_patch(FancyBboxPatch((x,0.9),0.62,0.9,boxstyle="round,pad=0.01,rounding_size=0.05",fc=col,ec="none"))
    ax.text(x+0.31,1.5,g,ha="center",va="center",color="white",fontsize=10,fontweight="bold")
    ax.text(x+0.31,1.15,l,ha="center",va="center",color="white",fontsize=7.2)
    if i<len(steps)-1: ax.annotate("",xy=(x+0.72,1.35),xytext=(x+0.63,1.35),arrowprops=dict(arrowstyle="-|>",color=GREY,lw=0.8))
ax.text(0.31+5*0.72,0.45,"Portão de autorização: docente e direção autorizam\nmaterial, imagem, marca, veiculação e orçamento.\nSem autorização, permanece simulação.",ha="center",va="center",fontsize=7.6,color=INK,linespacing=1.25)
ax.set_xlim(-0.05,6.6); ax.set_ylim(0,2.4); ax.axis("off")
fig.text(0.01,0.01,"Fluxo de atendimento da Agência-Escola (seção 14.2). Regra: o estudante não assume compromisso comercial em nome da escola.",fontsize=7.2,color=GREY)
fig.tight_layout(); fig.savefig(f"{OUT}/fig4-portoes-agencia.png"); plt.close(fig)

# 5. Governance / approval gates (sections 10-11, 12.1, Annex I1)
g=["Deliberação\nda Direção","Aceite\ninstitucional","Consulta aos\ndocentes","Piloto\n(1º tri 2027)","Revisão do piloto\ne indicadores","Ajuste e\nexpansão"]
fig, ax = plt.subplots(figsize=(6.5,2.5), dpi=200)
for i,t in enumerate(g):
    x=i*1.08
    col = GOLD if i==0 else (GREY if i in (1,2) else NAVY)
    ax.add_patch(FancyBboxPatch((x,0.8),0.92,0.95,boxstyle="round,pad=0.01,rounding_size=0.05",fc=col,ec="none"))
    ax.text(x+0.46,1.275,t,ha="center",va="center",color="white",fontsize=7.6,linespacing=1.15)
    if i<len(g)-1: ax.annotate("",xy=(x+1.08,1.27),xytext=(x+0.93,1.27),arrowprops=dict(arrowstyle="-|>",color=GREY,lw=0.8))
ax.text(0.46,0.45,"Gate 1: decisão\nexpressa (5 estados)",ha="center",fontsize=7.2,color=NAVY)
ax.text(3.7,0.45,"Gate 2: revisão\ntrimestral",ha="center",fontsize=7.2,color=NAVY)
ax.set_xlim(-0.05,6.5); ax.set_ylim(0,2.2); ax.axis("off")
fig.text(0.01,0.01,"Sequência de governança (seções 10, 11 e Anexo I1). Consulta a docentes ocorre somente após o aceite da Direção.",fontsize=7.2,color=GREY)
fig.tight_layout(); fig.savefig(f"{OUT}/fig5-portoes-governanca.png"); plt.close(fig)
print("ok")
