import os
import csv
import pdfplumber
import pytesseract
import unicodedata
import re
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance, ImageFilter
from pdfminer.high_level import extract_text
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def normalizar(texto):
    return unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII").lower()

def limpar_ocr(texto):
    return re.sub(r'[^a-zA-Z0-9]', '', texto).lower()

def preprocessar(imagem):
    img = imagem.convert("L")
    img = ImageEnhance.Contrast(img).enhance(2)
    img = img.filter(ImageFilter.MedianFilter())
    return img

def destacar_termo(texto, termo):
    return re.sub(f"({re.escape(termo)})", r">>>\1<<<", texto, flags=re.IGNORECASE)

def buscar_em_pdfs(pasta, termo):
    resultados = []
    termo_normalizado = limpar_ocr(normalizar(termo))

    for arquivo in os.listdir(pasta):
        if arquivo.lower().endswith(".pdf"):
            caminho = os.path.join(pasta, arquivo)
            try:
                texto_pdfminer = extract_text(caminho)
                if texto_pdfminer:
                    debug.insert(tk.END, f"[DEBUG] Arquivo: {arquivo} | Origem: Texto embutido\n")
                    debug.insert(tk.END, texto_pdfminer[:300] + "\n\n")

                    if termo_normalizado in limpar_ocr(normalizar(texto_pdfminer)):
                        trecho = destacar_termo(texto_pdfminer[:200], termo)
                        resultados.append((arquivo, "?", trecho, "Texto embutido"))
                        resultados_box.insert(tk.END, f"📄 Arquivo: {arquivo} | Página: ? | Origem: Texto embutido\nTrecho: {trecho}\n\n")
                    continue

                with pdfplumber.open(caminho) as pdf:
                    for i, pagina in enumerate(pdf.pages):
                        texto = pagina.extract_text()
                        origem = "Texto embutido"
                        if not texto:
                            imagens = convert_from_path(caminho, dpi=300, first_page=i+1, last_page=i+1)
                            texto = pytesseract.image_to_string(preprocessar(imagens[0]), lang="por")
                            origem = "OCR"
                        texto_limpo = texto.replace("\n", " ").strip()

                        debug.insert(tk.END, f"[DEBUG] Arquivo: {arquivo} | Página: {i+1} | Origem: {origem}\n")
                        debug.insert(tk.END, texto_limpo[:300] + "\n\n")

                        texto_normalizado = limpar_ocr(normalizar(texto_limpo))
                        if termo_normalizado in texto_normalizado:
                            trecho = destacar_termo(texto_limpo[:200], termo)
                            resultados.append((arquivo, i+1, trecho, origem))
                            resultados_box.insert(tk.END, f"📄 Arquivo: {arquivo} | Página: {i+1} | Origem: {origem}\nTrecho: {trecho}\n\n")
            except Exception as e:
                debug.insert(tk.END, f"[DEBUG] Erro ao abrir {arquivo}: {e}\n")
    return resultados

def iniciar_busca():
    resultados_box.delete(1.0, tk.END)
    debug.delete(1.0, tk.END)
    pasta = pasta_entry.get()
    termo = termo_entry.get()
    if not pasta or not termo:
        messagebox.showwarning("Aviso", "Escolha a pasta e digite o termo!")
        return
    global resultados
    resultados = buscar_em_pdfs(pasta, termo)
    if resultados:
        resultados_box.insert(tk.END, "✅ Busca concluída! Resultados exibidos acima.\n")
    else:
        resultados_box.insert(tk.END, "Nenhum resultado encontrado.\n")

def exportar_csv():
    if not resultados:
        messagebox.showwarning("Aviso", "Nenhum resultado para exportar!")
        return
    filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files","*.csv")])
    if filename:
        with open(filename, mode="w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Arquivo", "Página", "Trecho", "Origem"])
            for arquivo, pagina, trecho, origem in resultados:
                writer.writerow([arquivo, pagina, trecho, origem])
        messagebox.showinfo("Exportação", f"Resultados exportados para '{filename}'.")

# Interface Tkinter
root = tk.Tk()
root.title("Buscador de PDFs com OCR + pdfminer")

tk.Label(root, text="Selecione a pasta dos PDFs:").pack()
pasta_entry = tk.Entry(root, width=80)
pasta_entry.pack()
tk.Button(root, text="Escolher Pasta", command=lambda: pasta_entry.insert(0, filedialog.askdirectory())).pack()

tk.Label(root, text="Digite a palavra ou número:").pack()
termo_entry = tk.Entry(root, width=40)
termo_entry.pack()

tk.Button(root, text="Buscar", command=iniciar_busca).pack()
tk.Button(root, text="Exportar CSV", command=exportar_csv).pack()

tk.Label(root, text="Debug (texto lido):").pack()
debug = scrolledtext.ScrolledText(root, width=100, height=15)
debug.pack()

tk.Label(root, text="Resultados da busca:").pack()
resultados_box = scrolledtext.ScrolledText(root, width=100, height=15)
resultados_box.pack()

root.mainloop()
