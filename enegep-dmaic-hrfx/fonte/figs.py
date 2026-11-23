import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch
from PIL import Image
plt.rcParams["font.family"] = "DejaVu Serif"
plt.rcParams["font.size"] = 10

# ---------- Pareto ----------
causas = ["Ausência de\nchecklist\nformalizado", "Falta de\nendereçamento\nvisual", "Sobrecarga\ndo\nproprietário", "Ausência de\nplanejamento\nantecipado", "Outros\nfatores\npontuais"]
freq = [14, 12, 7, 5, 3]
tot = sum(freq); acum=[]; s=0
for f in freq: s+=f; acum.append(100*s/tot)
fig, ax = plt.subplots(figsize=(7.2, 3.6), dpi=200)
bars = ax.bar(range(5), freq, color=["#3B5B8C"]*3+["#9DB0CF"]*2, edgecolor="black", linewidth=0.6)
for i,f in enumerate(freq): ax.text(i, f/2, str(f), ha="center", va="center", fontsize=10, color="white", weight="bold")
ax.set_ylabel("Frequência (ocorrências)"); ax.set_ylim(0, 41); ax.set_xticks(range(5)); ax.set_xticklabels(causas, fontsize=8.5)
ax2 = ax.twinx(); ax2.plot(range(5), acum, color="#B03A2E", marker="o", linewidth=1.5)
for i,a in enumerate(acum): ax2.text(i+0.12, a+4, f"{a:.1f}%".replace(".",","), ha="center", fontsize=8, color="#B03A2E")
ax2.set_ylim(0, 110); ax2.set_ylabel("% acumulado"); ax2.axhline(80, color="gray", linestyle="--", linewidth=0.8)
ax.spines["top"].set_visible(False); ax2.spines["top"].set_visible(False)
fig.tight_layout(); fig.savefig("pareto.png", facecolor="white")

# ---------- Ishikawa ----------
fig, ax = plt.subplots(figsize=(9, 5.6), dpi=220)
ax.set_xlim(0, 110); ax.set_ylim(-1, 51); ax.axis("off")
ax.annotate("", xy=(86, 25), xytext=(1, 25), arrowprops=dict(arrowstyle="-|>", lw=2, color="black"))
ax.add_patch(plt.Rectangle((86.5, 16), 23.5, 18, fill=True, facecolor="#F2F2F2", edgecolor="black"))
ax.text(98.25, 25, "Falhas na\npreparação\ndos kits\n(itens esquecidos\ne separação lenta)", ha="center", va="center", fontsize=10.5, weight="bold")
top = [("Método", ["Separação de", "memória, sem", "checklist por tipo", "de evento", "Ausência de POP", "Preparação na", "véspera"]),
       ("Mão de obra", ["Conhecimento", "tácito concentrado", "no proprietário", "Auxiliares sem", "treinamento formal", "Sobrecarga às", "sextas-feiras"]),
       ("Meio ambiente", ["Depósito sem", "endereçamento", "visual", "Obsoletos", "misturados aos", "itens de uso", "Layout sem", "critério de giro"])]
bot = [("Material", ["Itens pequenos", "sem local fixo", "Estoque de baixo", "giro ocupando", "espaço", "Sem estoque", "mínimo"]),
       ("Medição", ["Sem indicadores", "de erro e", "de tempo", "Conferência", "visual sem", "registro"]),
       ("Máquina/Equipamento", ["Cabos e", "equipamentos sem", "identificação", "por kit", "Sem teste", "funcional antes", "da carga"])]
xs = [30, 57, 84]
for (cat, items), x in zip(top, xs):
    ax.annotate("", xy=(x, 25), xytext=(x-12, 47), arrowprops=dict(arrowstyle="-|>", lw=1.3))
    ax.text(x-12, 48.3, cat, ha="center", fontsize=12, weight="bold")
    ax.text(x-12.5, 43.5, "\n".join(items), ha="right", va="top", fontsize=9.3, linespacing=1.25)
for (cat, items), x in zip(bot, xs):
    ax.annotate("", xy=(x, 25), xytext=(x-12, 3), arrowprops=dict(arrowstyle="-|>", lw=1.3))
    ax.text(x-12, -0.8, cat, ha="center", fontsize=12, weight="bold")
    ax.text(x-10.5, 21.5, "\n".join(items), ha="right", va="top", fontsize=9.3, linespacing=1.25)
fig.tight_layout(); fig.savefig("ishikawa.png", facecolor="white")

# ---------- header logo without year ----------
im = Image.open("image2.jpg").convert("RGB")
from PIL import ImageDraw
d = ImageDraw.Draw(im); d.rectangle([100, 1620, 1240, 1830], fill="white")
im.save("logo.jpg", quality=92)
