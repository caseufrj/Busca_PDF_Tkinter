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

# Caminho do Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Funções auxiliares
def normalizar(texto):
    return unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII").lower()

def limpar_ocr(texto):
    return re.sub(r'[^a-zA-Z0-9]', '', texto).lower()

def preprocessar(imagem):
    img = imagem.convert("L")
    img = ImageEnhance.Contrast(img).enhance(3)
    img = img.filter(ImageFilter.SHARPEN)
    img = img.point(lambda x: 0 if x < 128 else 255, '1')
    return img

def destacar_termo(texto, termo):
    return re.sub(f"({re.escape(termo)})", r">>>\1<<<", texto, flags=re.IGNORECASE)

# Função principal de busca
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
                    if termo_normalizado in texto_normalizado:
                        trecho = destacar_termo(texto_pdfminer[:500], termo)
                        resultados.append((arquivo, "?", trecho, "Texto embutido"))
                        resultados_box.insert("end", f"📄 Arquivo: {arquivo} | Página: ? | Origem: Texto embutido\nTrecho: {trecho}\n\n")
                    continue

                imagens = convert_from_path(caminho, dpi=400)
                texto_total = ""
                for img in imagens:
                    texto_total += pytesseract.image_to_string(
                        preprocessar(img),
                        lang="por",
                        config="--psm 7 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                    ) + "\n"

                debug_box.insert("end", f"[OCR batch] Arquivo: {arquivo}\n{texto_total[:500]}\n\n")

                texto_normalizado = limpar_ocr(normalizar(texto_total))
                if termo_normalizado in texto_normalizado:
                    trecho = destacar_termo(texto_total[:500], termo)
                    resultados.append((arquivo, "OCR batch", trecho, "OCR"))
                    resultados_box.insert("end", f"📄 Arquivo: {arquivo} | Origem: OCR batch\nTrecho: {trecho}\n\n")

            except Exception as e:
                debug_box.insert("end", f"[DEBUG] Erro ao abrir {arquivo}: {e}\n")
    return resultados

# Botões e interface
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
root.iconbitmap("icone_app.ico")
root.title("Buscador de PDFs com OCR + pdfminer")
root.geometry("900x800")

frame = ctk.CTkFrame(root, fg_color="transparent")
frame.pack(padx=20, pady=20, fill="both", expand=True)

# Carregar ícones
icone_busca = ctk.CTkImage(light_image=Image.open("icone_busca.png"), size=(32, 32))
icone_download = ctk.CTkImage(light_image=Image.open("icone_download.png"), size=(32, 32))

# Título com ícone de lupa
ctk.CTkLabel(frame, image=icone_busca, text="Buscador de PDFs", compound="left").pack(pady=5)

ctk.CTkLabel(frame, text="Selecione a pasta dos PDFs:").pack()
pasta_entry = ctk.CTkEntry(frame, width=600)
pasta_entry.pack(pady=5)
ctk.CTkButton(frame, text="Escolher Pasta", command=lambda: pasta_entry.insert(0, filedialog.askdirectory())).pack()

ctk.CTkLabel(frame, text="Digite a palavra ou número:").pack()
termo_entry = ctk.CTkEntry(frame, width=300)
termo_entry.pack(pady=5)

# Botões com ícones
ctk.CTkButton(frame, image=icone_busca, text="Buscar", compound="left", command=iniciar_busca).pack(pady=5)
ctk.CTkButton(frame, image=icone_download, text="Exportar CSV", compound="left", command=exportar_csv).pack(pady=5)

ctk.CTkLabel(frame, text="Debug (texto lido):").pack()
debug_box = ctk.CTkTextbox(frame, width=800, height=200)
debug_box.pack(pady=5)

ctk.CTkLabel(frame, text="Resultados da busca:").pack()
resultados_box = ctk.CTkTextbox(frame, width=800, height=250)
resultados_box.pack(pady=5)

root.mainloop()

