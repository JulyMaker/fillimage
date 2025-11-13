import pandas as pd
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import numpy as np
from collections import Counter
from PIL import Image
import argparse
import random

def generateWordsCloud(edges=True, imageName="cinco.png", minSize=10, background="white", palabra=""):
    # === Excel ===
    df = pd.read_excel('nombres.xlsx', sheet_name='Hoja1')
    firstCol = df.columns[0]

    finalFreq = Counter(df[firstCol].dropna().astype(str))  # Name Counts

    # === Center Word ===
    if palabra:
        print(f"\n✅ Palabra central solicitada")

    totalNames = sum(finalFreq.values())
    uniqueNames = len(finalFreq)

    # === Load Mask (RGB or RGBA) ===
    img = Image.open(imageName).convert("L")  # Escala de grises (1 canal)
    mask = np.array(img)

    mean_value = np.mean(mask)
    if mean_value > 127:
        # Fondo blanco, número negro → OK
        print("\n🎨 Detección de máscara: fondo blanco, imagen oscura (correcto).")
    else:
        # Fondo negro, imagen blanca → invertir
        # mask = 255 - mask
        print("\n🎨 Detección de máscara: fondo oscuro, imagen clara → máscara invertida automáticamente.")

    # === Colors functions ===
    def color_func(word, font_size, position, orientation, random_state=None, **kwargs):
        randomValue = random.randint(0, 255)
        color = plt.cm.viridis(randomValue / 255.0)
        return f"rgb({int(color[0]*255)}, {int(color[1]*255)}, {int(color[2]*255)})"

    # === Generate word cloud ===
    wordcloud = WordCloud(
        #width=1000,
        #height=800,
        background_color=background,
        mask=mask,
        contour_width=2 if edges else 0,
        contour_color='black',
        colormap='viridis',
        color_func=color_func,
        relative_scaling=0.1,
        max_words=totalNames,
        prefer_horizontal=0.6,
        min_font_size=minSize,
        max_font_size=80,
        collocations=False,
        repeat=True,  # Para rellenar bien la figura
        random_state=42,
        scale=2,
        margin=1
    ).generate_from_frequencies(finalFreq)

    # === Show ===
    plt.figure(figsize=(16, 12))
    plt.imshow(wordcloud, interpolation='bilinear')
    plt.axis('off')

    tam=60
    transparency = 0.5
    # === Mask Center ===
    if palabra:
        ys, xs = np.where(mask < 128)
        cx, cy = np.mean(xs), np.mean(ys)

        plt.figtext(
            0.5, 0.5,
            palabra,
            fontsize=tam,
            color='white',
            ha='center',
            va='center',
            alpha= transparency,
            fontweight='bold'
            #,bbox=dict(boxstyle="round,pad=0.3", facecolor='black', alpha=0.6)
        )

    plt.tight_layout()

    # === Save final image ===
    baseName = imageName.split(".")[0]
    outName = f'finalImage_{baseName}_{palabra if palabra else "nube"}.png'

    plt.savefig(outName, dpi=350, bbox_inches='tight', facecolor=background, pad_inches=0)
    plt.show()

    # === Show INFO ===
    print("\n📊 ESTADÍSTICAS DE GENERACIÓN:")
    print(f"  - Total de nombres en Excel: {totalNames}")
    print(f"  - Nombres únicos: {uniqueNames}")
    print(f"  - Palabra central: {palabra if palabra else 'Ninguna'} (tamaño: {tam if palabra else "0"}  transparency: { transparency if palabra else "0"})")
    print(f"  - Archivo guardado: {outName}")

    print(f"\n⚙️ Configuración utilizada:")
    print(f"  - Bordes: {'Sí' if edges else 'No'}")
    print(f"  - Imagen: {imageName}")
    print(f"  - Tamaño mínimo: {minSize}")
    print(f"  - Color fondo: {background}")

    print("\nTop 20 nombres más frecuentes:")
    for name, count in finalFreq.most_common(20):
        print(f"  {name}: {count} veces")

# ===================== MAIN =====================
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Generar nube de palabras personalizada')

    parser.add_argument('--bordes', type=str, default='si', choices=['si', 'no'],
                        help='Dibujar bordes en la imagen (si/no) - por defecto: si')

    parser.add_argument('--imagen', type=str, default='cinco.png',
                        help='Archivo de imagen máscara - por defecto: cinco.png')

    parser.add_argument('--tamaño-minimo', type=int, default=6,
                        help='Tamaño mínimo de las palabras - por defecto: 6')

    parser.add_argument('--fondo', type=str, default='white',
                        help='Color de fondo de la nube - por defecto: white')

    parser.add_argument('--palabra', type=str, default='',
                        help='Palabra central grande y centrada - por defecto: vacío')

    args = parser.parse_args()

    bordes = args.bordes.lower() == 'si'

    generateWordsCloud(
        edges=bordes,
        imageName=args.imagen,
        minSize=args.tamaño_minimo,
        background=args.fondo,
        palabra=args.palabra
    )
