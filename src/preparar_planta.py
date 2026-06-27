from pathlib import Path
from PIL import Image

BASE_DIR = Path(__file__).resolve().parent.parent

entrada = BASE_DIR / "assets" / "planta.png"
saida = BASE_DIR / "assets" / "planta_base.jpg"

img = Image.open(entrada).convert("RGBA")

fundo = Image.new("RGBA", img.size, "WHITE")
fundo.paste(img, mask=img.split()[3])

fundo.convert("RGB").save(saida, quality=95)

print(f"Planta preparada em: {saida}")
