import os
import csv
import pdfplumber
import pytesseract
import unicodedata
import re
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance, ImageFilter
from pdfminer.high_level import extract_text
import customtkinter as ctk
from tkinter import filedialog, messagebox

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def normalizar(texto):
    return unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII").lower()

def limpar_ocr(texto):
    return re.sub(r'[^a-zA-Z0-9]', '', texto).lower()

# 🔧 Pré-processamento mais forte para OCR
def preprocessar(imagem):
    img = imagem.convert("L")  # escala de cinza
    img = ImageEnhance.Contrast(img).enhance(3)  # aumenta contraste
    img = img.filter(ImageFilter.SHARPEN)        # nitidez
    img = img.point(lambda x: 0 if x < 128 else 255, '1')  # binarização
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
                    debug_box.insert("end", f"[DEBUG] Arquivo: {arquivo} | Origem: Texto embutido\n")
                    debug_box.insert("end", texto_pdfminer[:300] + "\n\n")

                    texto_normalizado = limpar_ocr(normalizar(texto_pdfminer))
                    # 🔧 Regex exata
                    if re.search(rf"\b{re.escape(termo_normalizado)}\b", texto_normalizado):
                        trecho = destacar_termo(texto_pdfminer[:200], termo)
                        resultados.append((arquivo, "?", trecho, "Texto embutido"))
                        resultados_box.insert("end", f"📄 Arquivo: {arquivo} | Página: ? | Origem: Texto embutido\nTrecho: {trecho}\n\n")
                    continue

                with pdfplumber.open(caminho) as pdf:
                    for i, pagina in enumerate(pdf.pages):
                        texto = pagina.extract_text()
                        origem = "Texto embutido"
                        if not texto:
                            imagens = convert_from_path(caminho, dpi=300, first_page=i+1, last_page=i+1)
                            # 🔧 OCR configurado para números/letras
                            texto = pytesseract.image_to_string(
                                preprocessar(imagens[0]),
                                lang="por",
                                config="--psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                            )
                            origem = "OCR"
                        texto_limpo = texto.replace("\n", " ").strip()

                        debug_box.insert("end", f"[DEBUG] Arquivo: {arquivo} | Página: {i+1} | Origem: {origem}\n")
                        debug_box.insert("end", texto_limpo[:300] + "\n\n")

                        texto_normalizado = limpar_ocr(normalizar(texto_limpo))
                        # 🔧 Regex exata
                        if re.search(rf"\b{re.escape(termo_normalizado)}\b", texto_normalizado):
                            trecho = destacar_termo(texto_limpo[:200], termo)
                            resultados.append((arquivo, i+1, trecho, origem))
                            resultados_box.insert("end", f"📄 Arquivo: {arquivo} | Página: {i+1} | Origem: {origem}\nTrecho: {trecho}\n\n")
            except Exception as e:
                debug_box.insert("end", f"[DEBUG] Erro ao abrir {arquivo}: {e}\n")
    return resultados

def iniciar_busca():
    resultados_box.delete("1.0", "end")
    debug_box.delete("1.0", "end")
    pasta = pasta_entry.get()
    termo = termo_entry.get()
    if not pasta or not termo:
        messagebox.showwarning("Aviso", "Escolha a pasta e digite o termo!")
        return
    global resultados
    resultados = buscar_em_pdfs(pasta, termo)
    resultados_box.insert("end", f"🔍 Resultado para a busca: {termo}\n\n")
    if resultados:
        resultados_box.insert("end", "✅ Busca concluída! Resultados exibidos acima.\n")
    else:
        resultados_box.insert("end", "Nenhum resultado encontrado.\n")

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

# Interface CustomTkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Buscador de PDFs com OCR + pdfminer")
root.geometry("900x800")

frame = ctk.CTkFrame(root, fg_color="transparent")
frame.pack(padx=20, pady=20, fill="both", expand=True)

ctk.CTkLabel(frame, text="Selecione a pasta dos PDFs:").pack()
pasta_entry = ctk.CTkEntry(frame, width=600)
pasta_entry.pack(pady=5)
ctk.CTkButton(frame, text="Escolher Pasta", command=lambda: pasta_entry.insert(0, filedialog.askdirectory())).pack()

ctk.CTkLabel(frame, text="Digite a palavra ou número:").pack()
termo_entry = ctk.CTkEntry(frame, width=300)
termo_entry.pack(pady=5)

ctk.CTkButton(frame, text="Buscar", command=iniciar_busca).pack(pady=5)
ctk.CTkButton(frame, text="Exportar CSV", command=exportar_csv).pack(pady=5)

ctk.CTkLabel(frame, text="Debug (texto lido):").pack()
debug_box = ctk.CTkTextbox(frame, width=800, height=200)
debug_box.pack(pady=5)

ctk.CTkLabel(frame, text="Resultados da busca:").pack()
resultados_box = ctk.CTkTextbox(frame, width=800, height=250)
resultados_box.pack(pady=5)

root.mainloop()
