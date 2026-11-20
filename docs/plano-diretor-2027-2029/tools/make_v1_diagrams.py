"""Original vector diagrams for the Planejamento Pedagógico Trienal V1 (no results, no people).
Usage: python3 tools/make_v1_diagrams.py <out_dir>   -> writes PNG (300 dpi), SVG and PDF per figure."""
import sys, os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Circle, FancyArrowPatch
import numpy as np
OUT = sys.argv[1]; os.makedirs(OUT, exist_ok=True)
NAVY="#1F2F4D"; GOLD="#B8975A"; GREY="#8C8279"; CREAM="#F4F0E8"; INK="#2B2B2B"; PALE="#E9E4DA"; WHITE="#FFFFFF"
plt.rcParams.update({"font.family":"Carlito","font.size":9.5})
def save(fig, name):
    for ext, kw in (("png", {"dpi":300}), ("svg", {}), ("pdf", {})):
        fig.savefig(f"{OUT}/{name}.{ext}", **kw)
    plt.close(fig)
def box(ax, x, y, w, h, text, fc=CREAM, ec=GREY, tc=INK, fs=8.5, bold=False):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.02,rounding_size=0.06", fc=fc, ec=ec, lw=0.8))
    ax.text(x+w/2, y+h/2, text, ha="center", va="center", fontsize=fs, color=tc, fontweight="bold" if bold else "normal", linespacing=1.2)
def arrow(ax, x1, y1, x2, y2, color=GREY):
    ax.add_patch(FancyArrowPatch((x1,y1),(x2,y2), arrowstyle="-|>", mutation_scale=9, color=color, lw=0.9))

# F1 — trajectories of the two 2027 cohorts (PDP §5 / Anexo F; planned, not executed)
fig, ax = plt.subplots(figsize=(6.6, 3.6))
ax.set_xlim(0, 10); ax.set_ylim(0, 6); ax.axis("off")
for i, lab in enumerate(["2027", "2028", "2029"]):
    ax.add_patch(FancyBboxPatch((0.2+i*3.2, 5.2), 2.9, 0.55, boxstyle="round,pad=0.02,rounding_size=0.05", fc=NAVY, ec="none"))
    ax.text(0.2+i*3.2+1.45, 5.475, f"Ano {lab}", ha="center", va="center", color=WHITE, fontsize=10, fontweight="bold")
rowA = [("1º ano · Coorte A\nobservar, nomear, investigar", "Ingressantes em 2027"),
        ("2º ano · Coorte A\nexperimentar, testar, medir", ""),
        ("3º ano · Coorte A\nintegrar, operar, defender", "")]
for i,(t,_) in enumerate(rowA):
    box(ax, 0.2+i*3.2, 3.2, 2.9, 1.3, t, fc=CREAM, ec=GOLD)
    if i < 2: arrow(ax, 0.2+i*3.2+2.9, 3.85, 0.2+(i+1)*3.2, 3.85, color=NAVY)
ax.text(0.2, 2.85, "Coorte A — trilha completa de 2027 a 2029 (plano proposto)", fontsize=8.5, color=NAVY, fontweight="bold")
box(ax, 0.2, 0.9, 2.9, 1.2, "3º ano · Coorte B\nconcluintes em 2027\nintegrar e concluir", fc=PALE, ec=NAVY)
ax.text(3.4, 1.5, "Trilha de transição própria: sem exigir retroativamente\nos dois primeiros anos do programa novo.", fontsize=8, color=INK, va="center")
ax.text(0.2, 0.45, "Legenda: planejado · dependente de autorização · não executado. Nenhum resultado é apresentado como alcançado.", fontsize=7.5, color=GREY)
save(fig, "v1-fig1-trilha-coortes")

# F2 — classroom cycle: six movements (PDP §3.1 / V1 §3.1)
fig, ax = plt.subplots(figsize=(6.6, 6.2))
ax.set_xlim(-2.2, 2.2); ax.set_ylim(-2.25, 2.25); ax.set_aspect("equal"); ax.axis("off")
mov = ["1. Mobilizar", "2. Compreender", "3. Experimentar", "4. Produzir", "5. Defender", "6. Melhorar"]
ev = ["pergunta e\nhipótese", "nota\nconceitual", "rascunho e\nprotótipo", "versão 1 e\nficha de autoria", "defesa e\nrubrica", "versão 2 e\nreflexão"]
for k,(m,e) in enumerate(zip(mov, ev)):
    ang = np.pi/2 - k*np.pi/3
    x, y = 1.3*np.cos(ang), 1.3*np.sin(ang)
    ax.add_patch(Circle((x,y), 0.5, fc=NAVY if k%2==0 else GOLD, ec="none"))
    ax.text(x, y+0.13, m, ha="center", va="center", color=WHITE, fontsize=10.5, fontweight="bold")
    ax.text(x, y-0.17, e, ha="center", va="center", color=WHITE, fontsize=8.2, linespacing=1.05)
    nang = ang - np.pi/6
    arrow(ax, 0.85*np.cos(ang-np.pi/9), 0.85*np.sin(ang-np.pi/9), 0.85*np.cos(ang-np.pi/6-np.pi/9), 0.85*np.sin(ang-np.pi/6-np.pi/9))
ax.text(0, 0.12, "Ciclo pedagógico", ha="center", fontsize=12.5, color=NAVY, fontweight="bold")
ax.text(0, -0.2, "em seis movimentos", ha="center", fontsize=10.5, color=INK)
ax.text(0, -2.08, "Evidência em cada movimento; o feedback retorna ao início do ciclo.\nSem percentuais de carga horária: é desenho didático, não grade do PPC.", ha="center", fontsize=8.8, color=GREY)
save(fig, "v1-fig2-ciclo-seis-movimentos")

# F3 — assessment: three planes, recovery loop (V1 §12.1 / §12.4; no weights, no numbers)
fig, ax = plt.subplots(figsize=(6.6, 3.5))
ax.set_xlim(0, 10); ax.set_ylim(0, 6); ax.axis("off")
planes = [("Conhecimento", "prova individual, questões\ncontextualizadas, análise de caso"),
          ("Aplicação", "artefato, planilha, roteiro,\nprojeto e demonstração"),
          ("Metacognição", "defesa, reflexão e comparação\nentre versões")]
for i,(t,d) in enumerate(planes):
    box(ax, 0.2+i*3.25, 3.5, 2.9, 2.1, f"{t}\n\n{d}", fc=CREAM, ec=GOLD, fs=7.6)
ax.text(0.2, 3.1, "Três planos de observação — cada um com evidência individual", fontsize=8.5, color=NAVY, fontweight="bold")
box(ax, 0.2, 0.8, 2.6, 1.2, "Instrumento\ninsuficiente", fc=PALE, ec=NAVY, fs=8.5)
box(ax, 3.7, 0.8, 2.6, 1.2, "Identificar a\nhabilidade (não só a nota)", fc=CREAM, ec=GREY, fs=8)
box(ax, 7.2, 0.8, 2.6, 1.2, "Exercício equivalente e\nnova demonstração", fc=CREAM, ec=GREY, fs=8)
arrow(ax, 2.8, 1.4, 3.7, 1.4, color=NAVY); arrow(ax, 6.3, 1.4, 7.2, 1.4, color=NAVY)
ax.text(5, 0.35, "Registro e devolutiva no sistema escolar. Recuperação paralela segue o PPC (pp. 38–40).", ha="center", fontsize=7.5, color=GREY)
save(fig, "v1-fig3-avaliacao-recuperacao")

# F4 — project gates G0–G6 (V1 §13.2)
fig, ax = plt.subplots(figsize=(6.6, 2.7))
ax.set_xlim(0, 10); ax.set_ylim(0, 4); ax.axis("off")
gates = [("G0", "solicitação"), ("G1", "briefing"), ("G2", "diagnóstico"), ("G3", "proposta"), ("G4", "autorização"), ("G5", "execução"), ("G6", "encerramento")]
w = 1.36
for i,(g,t) in enumerate(gates):
    x = 0.1 + i*(w+0.07)
    fc = GOLD if g == "G4" else NAVY
    ax.add_patch(FancyBboxPatch((x, 1.5), w, 1.4, boxstyle="round,pad=0.02,rounding_size=0.05", fc=fc, ec="none"))
    ax.text(x+w/2, 2.45, g, ha="center", va="center", color=WHITE, fontsize=10, fontweight="bold")
    ax.text(x+w/2, 1.95, t, ha="center", va="center", color=WHITE, fontsize=7.5)
    if i < len(gates)-1: arrow(ax, x+w, 2.2, x+w+0.07, 2.2)
ax.text(0.1, 0.9, "G4: a Direção ou o docente responsável aprova o que será exibido ou executado.", fontsize=8, color=INK)
ax.text(0.1, 0.45, "Cliente real não implica autorização para publicar, contratar, receber valores, prometer desempenho,\nacessar credenciais ou divulgar imagem de estudantes.", fontsize=7.5, color=GREY)
save(fig, "v1-fig4-portoes-projeto")

# F5 — governance gates and their current state (V1 §19; states as written, nothing approved)
fig, ax = plt.subplots(figsize=(6.6, 3.6))
ax.set_xlim(0, 10); ax.set_ylim(0, 5.4); ax.axis("off")
steps = [("G1", "Revisão documental", "DRAFT"), ("G2", "Deliberação da Direção", "PENDENTE"),
         ("G3", "Pactuação docente e prontidão", "PENDENTE"), ("G4", "Autorização por projeto", "PENDENTE POR PROJETO"),
         ("G5", "Acompanhamento", "NÃO INICIADO")]
w = 1.8
for i,(g,t,s) in enumerate(steps):
    x = 0.1 + i*(w+0.1)
    box(ax, x, 1.5, w, 2.8, f"{g}\n\n" + t.replace(" ", "\n", 1), fc=CREAM, ec=GOLD, fs=7.6)
    ax.add_patch(FancyBboxPatch((x+0.1, 0.5), w-0.2, 0.75, boxstyle="round,pad=0.01,rounding_size=0.04", fc=WHITE, ec=NAVY, lw=0.8))
    ax.text(x+w/2, 0.875, s.replace(" POR ", "\nPOR "), ha="center", va="center", fontsize=6.6, color=NAVY, fontweight="bold", linespacing=1.05)
    if i < len(steps)-1: arrow(ax, x+w, 2.9, x+w+0.1, 2.9)
ax.text(0.1, 4.95, "Portas de decisão separadas: nenhuma etapa além de G1 está concluída neste documento.", fontsize=8.5, color=NAVY, fontweight="bold")
save(fig, "v1-fig5-governanca-portoes")
print("diagrams:", sorted(f for f in os.listdir(OUT) if f.endswith(".png")))
