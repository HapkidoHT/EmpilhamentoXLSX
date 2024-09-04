import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

def combine_excel_files(directory, progress_bar):
    # Lista para armazenar os DataFrames de cada planilha
    data_frames = []

    # Obter a lista de arquivos Excel no diretório
    files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]
    total_files = len(files)

    # Loop através de todos os arquivos no diretório
    for i, filename in enumerate(files):
        file_path = os.path.join(directory, filename)
        df = pd.read_excel(file_path, sheet_name='Sheet1') #Adequar com a sheet de cada planilha.
        data_frames.append(df)

        # Atualizar a barra de progresso
        progress_bar['value'] = ((i + 1) / total_files) * 100
        root.update_idletasks()

    # Concatenar todos os DataFrames em um único DataFrame
    combined_df = pd.concat(data_frames, ignore_index=True)
    
    # Caminho para salvar a planilha combinada no diretório selecionado
    save_path = os.path.join(directory, 'planilha_combinada.xlsx')
    
    # Salvar o DataFrame combinado em uma nova planilha
    combined_df.to_excel(save_path, index=False)

    return save_path

def select_directory():
    directory = filedialog.askdirectory()
    if directory:
        try:
            progress_bar['value'] = 0
            save_path = combine_excel_files(directory, progress_bar)
            messagebox.showinfo("Sucesso", f"As planilhas foram combinadas com sucesso!\nSalvo em: {save_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao combinar as planilhas:\n{e}")

# Configurar a interface gráfica
root = tk.Tk()
root.title("Combinar Planilhas Excel")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack(padx=10, pady=10)

label = tk.Label(frame, text="Selecione o diretório onde estão as planilhas:")
label.pack(pady=10)

button = tk.Button(frame, text="Selecionar Diretório", command=select_directory)
button.pack(pady=10)

progress_bar = ttk.Progressbar(frame, orient='horizontal', length=300, mode='determinate')
progress_bar.pack(pady=20)

root.mainloop()