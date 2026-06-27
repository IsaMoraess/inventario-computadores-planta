from pathlib import Path

import matplotlib.pyplot as plt
from PIL import Image

BASE_DIR = Path(__file__).resolve().parent.parent
PLANTA_PATH = BASE_DIR / "assets" / "planta.png"

img = Image.open(PLANTA_PATH)

fig, ax = plt.subplots(figsize=(14, 10))
ax.imshow(img)
ax.set_xlim(0, img.width)
ax.set_ylim(img.height, 0)
plt.subplots_adjust(left=0, right=1, top=0.95, bottom=0)
ax.set_title("Clique nos computadores. Feche a janela quando terminar.")
ax.axis("off")

pontos = []


def onclick(event):
    if event.xdata is None or event.ydata is None:
        return

    x = int(event.xdata)
    y = int(event.ydata)

    pontos.append((x, y))

    numero = len(pontos)
    ax.scatter(x, y, s=80)
    ax.text(x + 5, y - 5, f"PC-{numero:02d}", fontsize=9)

    print(f"PC-{numero:02d},{x},{y}")

    fig.canvas.draw()


fig.canvas.mpl_connect("button_press_event", onclick)
plt.show()

print("\nCoordenadas capturadas:")
for i, (x, y) in enumerate(pontos, start=1):
    print(f"PC-{i:02d},Sala, {x}, {y}, Ativo, 192.168.0.{i+10}, TI{i:03d}")
