# processing.py

import os
import zipfile
from datetime import datetime
from PyQt5.QtCore import QThread, pyqtSignal
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# Importa os módulos de geração de arquivos
import dxf_engine
import pdf_generator 

class ProcessThread(QThread):
    update_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, dataframe_to_process, generate_pdf=True, generate_dxf=True, project_directory=".", project_number=""):
        super().__init__()
        self.df = dataframe_to_process
        self.generate_pdf = generate_pdf
        self.generate_dxf = generate_dxf
        self.project_directory = project_directory
        self.project_number = project_number

    def run(self):
        try:
            total_items = len(self.df)
            if total_items == 0:
                self.finished_signal.emit(True, "Nada a processar. A lista de peças está vazia.")
                return

            self.update_signal.emit("Iniciando processamento...")

            if self.generate_pdf:
                self.update_signal.emit("--- Gerando PDFs ---")
                pdf_output_dir = os.path.join(self.project_directory, "PDFs")
                os.makedirs(pdf_output_dir, exist_ok=True)
                
                df_pdf = self.df.copy()
                df_pdf['espessura'] = df_pdf['espessura'].fillna('Sem_Espessura')
                grouped = df_pdf.groupby('espessura')

                for espessura, group in grouped:
                    espessura_str = str(espessura).replace('.', '_')
                    pdf_filename = os.path.join(pdf_output_dir, f"{self.project_number}_Desenhos_PDF_Espessura_{espessura_str}mm.pdf")
                    c = canvas.Canvas(pdf_filename, pagesize=A4)
                    for _, row in group.iterrows():
                        # Usa a função do módulo pdf_generator
                        pdf_generator.desenhar_forma(c, row)
                        c.showPage()
                    c.save()
                    self.update_signal.emit(f"PDF salvo em: {pdf_filename}")

            if self.generate_dxf:
                self.update_signal.emit("--- Gerando DXFs ---")
                dxf_output_dir = os.path.join(self.project_directory, "DXFs")
                os.makedirs(dxf_output_dir, exist_ok=True)
                zip_filename = os.path.join(dxf_output_dir, f"LOTE_DXF_{datetime.now():%Y%m%d_%H%M%S}.zip")
                
                with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for index, row in self.df.iterrows():
                        raw_data = row.to_dict()
                        
                        prepared_data, error = dxf_engine.prepare_and_validate_dxf_data(raw_data)
                        if error:
                            self.update_signal.emit(f"AVISO: Pulando DXF '{raw_data.get('nome_arquivo')}': {error}")
                            continue
                        
                        dxf_content, filename = dxf_engine.create_dxf_drawing(prepared_data)
                        if dxf_content:
                            zf.writestr(filename, dxf_content)
                        else:
                            self.update_signal.emit(f"ERRO: Falha ao gerar DXF para '{raw_data.get('nome_arquivo')}'.")
                        
                        progress = int(((index + 1) / total_items) * 100)
                        self.progress_signal.emit(progress)

                self.update_signal.emit(f"Arquivo ZIP com DXFs salvo em: {zip_filename}")
            
            self.finished_signal.emit(True, "Processamento concluído com sucesso!")

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.finished_signal.emit(False, f"Erro crítico no processamento: {str(e)}")