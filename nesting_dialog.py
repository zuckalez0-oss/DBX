import sys
import os
import colorsys
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QFormLayout, QLineEdit, 
                             QPushButton, QDialogButtonBox, QMessageBox, QComboBox,
                             QGroupBox, QLabel, QWidget, QHBoxLayout, QScrollArea,
                             QFileDialog)


import logging
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QPointF
from PyQt5.QtGui import QPainter, QColor, QPen, QBrush, QPainterPath
# Importações para gerar PDF
import ezdxf
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import pdf_generator
# Importe sua função de cálculo
from calculo_cortes import orquestrar_planos_de_corte, status_signaler

# --- INÍCIO: CLASSE DA THREAD DE CÁLCULO ---
class CalculationThread(QThread):
    """Thread para executar o cálculo de nesting em segundo plano."""
    # --- INÍCIO: NOVA FUNÇÃO PARA OFFSET DINÂMICO ---
    # --- CORREÇÃO: A função agora prioriza o input do usuário se for diferente do padrão '8'. ---
    def _get_dynamic_offset_and_margin(self, espessura, default_offset, default_margin):
        """Retorna o offset e a margem com base na espessura."""
        # Se o usuário inseriu um valor diferente do padrão (8), usa o valor do usuário.
        if abs(default_offset - 8.0) > 1e-5:
            return default_offset, default_margin

        # Regra 1: 0 a 6.35mm
        if 0 < espessura <= 6.35:
            return 5, 10
        # Regra 2: 6.35 a 15.88mm
        elif 6.35 < espessura <= 15.88:
            return 10, default_margin
        # Regra 3: 15.88 a 20mm
        elif 15.88 < espessura <= 20:
            return 17, default_margin
        # Regra 4: Exatamente 22.22mm
        elif abs(espessura - 22.22) < 1e-5:
            return 20, default_margin
        # Regra 5: 25.4 a 38mm
        elif 25.4 <= espessura <= 38:
            return 25, default_margin
        # Fallback para os valores da UI se nenhuma regra for atendida
        return default_offset, default_margin
    # --- FIM: NOVA FUNÇÃO PARA OFFSET DINÂMICO ---
    # Sinal que emite o resultado para uma espessura: (espessura, resultado_dict)
    result_ready = pyqtSignal(float, dict)
    # Sinal emitido quando todos os cálculos terminam
    finished = pyqtSignal()
    # Sinal para reportar erros: (titulo_erro, mensagem_erro)
    error = pyqtSignal(str, str)
    # Sinal para atualizações de status em tempo real
    status_update = pyqtSignal(str)

    def __init__(self, chapa_largura, chapa_altura, offset, margin, method, grouped_df, parent=None):
        super().__init__(parent)
        self.chapa_largura = chapa_largura
        self.chapa_altura = chapa_altura
        self.offset = offset
        self.grouped_df = grouped_df
        self.method = method
        self.margin = margin

    def run(self):
        try:
            logging.info("Thread de cálculo iniciada.")
            for espessura, group in self.grouped_df:
                # --- INÍCIO: LÓGICA DE OFFSET E MARGEM DINÂMICOS ---
                is_guillotine = self.method == "Guilhotina"

                if is_guillotine:
                    logging.info(f"Usando método Guilhotina para espessura {espessura}mm.")
                    current_offset = 0
                    refila = 2 * espessura
                    sheet_width_for_calc = self.chapa_largura - refila
                    sheet_height_for_calc = self.chapa_altura
                    effective_margin = 0
                else: # Plasma/Laser
                    logging.info(f"Usando método Plasma/Laser para espessura {espessura}mm.")
                    current_offset, _ = self._get_dynamic_offset_and_margin(espessura, self.offset, self.margin)
                    # A margem efetiva é calculada para garantir 10mm da borda até a peça real.
                    effective_margin = 10 - (current_offset / 2)
                    sheet_width_for_calc = self.chapa_largura
                    sheet_height_for_calc = self.chapa_altura

                logging.info(f"Para espessura {espessura}mm, usando Offset: {current_offset}mm. Margem efetiva: {effective_margin}mm.")

                pecas_para_calcular = []
                # --- INÍCIO: LÓGICA PARA INCLUIR CÍRCULOS NO CÁLCULO ---
                for _, row in group.iterrows():
                    if row['forma'] == 'rectangle' and row['largura'] > 0 and row['altura'] > 0:
                        pecas_para_calcular.append({
                            'forma': 'rectangle',
                            'largura': row['largura'] + current_offset,
                            'altura': row['altura'] + current_offset,
                            'quantidade': int(row['qtd']),
                            'furos': row.get('furos', [])
                        })
                    elif row['forma'] == 'circle' and row['diametro'] > 0:
                        pecas_para_calcular.append({
                            'forma': 'circle',
                            'largura': row['diametro'] + current_offset, # Bounding box
                            'altura': row['diametro'] + current_offset, # Bounding box
                            'diametro': row['diametro'], # Diâmetro original
                            'quantidade': int(row['qtd']),
                            'furos': row.get('furos', [])
                        })
                    elif row['forma'] == 'right_triangle' and row['rt_base'] > 0 and row['rt_height'] > 0:
                        # Para otimizar o encaixe, tratamos pares de triângulos como um único retângulo.
                        # Isso permite que o algoritmo de nesting (que trabalha com retângulos) os posicione
                        # de forma eficiente, garantindo o offset correto inclusive na hipotenusa.
                        num_pecas = int(row['qtd'])
                        num_pares = num_pecas // 2
                        # Adiciona os pares como retângulos
                        if num_pares > 0:
                            pecas_para_calcular.append({'forma': 'paired_triangle', 'largura': row['rt_base'] + current_offset, 'altura': row['rt_height'] + current_offset, 'quantidade': num_pares, 'furos': []})
                        # Adiciona a peça ímpar sobrando, se houver
                        if num_pecas % 2 == 1:
                            pecas_para_calcular.append({'forma': 'right_triangle', 'largura': row['rt_base'] + current_offset, 'altura': row['rt_height'] + current_offset, 'quantidade': 1, 'furos': []})

                    elif row['forma'] == 'trapezoid' and row['trapezoid_large_base'] > 0 and row['trapezoid_height'] > 0:
                        pecas_para_calcular.append({
                            'forma': 'trapezoid',
                            'largura': row['trapezoid_large_base'] + current_offset, # Bounding box
                            'altura': row['trapezoid_height'] + current_offset, # Bounding box
                            'small_base': row['trapezoid_small_base'] + current_offset,
                            'quantidade': int(row['qtd']),
                            'furos': row.get('furos', [])
                        })
                    elif row['forma'] == 'dxf_shape' and row['largura'] > 0 and row['altura'] > 0:
                        pecas_para_calcular.append({
                            'forma': 'dxf_shape',
                            'largura': row['largura'] + current_offset,
                            'altura': row['altura'] + current_offset,
                            'dxf_path': row['dxf_path'],
                            'quantidade': int(row['qtd']),
                            'furos': row.get('furos', [])
                        })
                # --- FIM: LÓGICA PARA INCLUIR CÍRCULOS ---

                if not pecas_para_calcular:
                    continue
                
                logging.debug(f"Iniciando cálculo para espessura {espessura} com {len(pecas_para_calcular)} tipos de peças.")
                # Chama a função de cálculo pesada
                resultado = orquestrar_planos_de_corte(sheet_width_for_calc, sheet_height_for_calc, pecas_para_calcular, current_offset, effective_margin, espessura, is_guillotine, status_signal_emitter=self.status_update)
                logging.debug(f"Cálculo para espessura {espessura} concluído. Emitindo resultado.")
                self.result_ready.emit(espessura, resultado)
        except Exception as e:
            logging.error(f"Erro na thread de cálculo: {e}", exc_info=True) # exc_info=True para logar o traceback
            self.error.emit(f"Erro no Cálculo (Espessura {espessura}mm)", str(e))
        finally:
            logging.info("Thread de cálculo finalizada.")
            self.finished.emit()

# --- INÍCIO: FUNÇÃO PARA GERAR CORES DISTINTAS ---
def generate_distinct_colors(n):
    """Gera N cores visualmente distintas."""
    colors = []
    for i in range(n):
        hue = i / n
        # Usamos saturação e valor altos para cores vibrantes, mas não totalmente saturadas para não cansar a vista.
        saturation = 0.85
        value = 0.9
        rgb = colorsys.hsv_to_rgb(hue, saturation, value)
        colors.append(QColor(int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255)))
    return colors
# --- FIM: FUNÇÃO PARA GERAR CORES DISTINTAS ---

def _draw_dxf_entities(painter, dxf_path, offset_x, offset_y, scale):
    """Lê um arquivo DXF e desenha suas entidades usando QPainter."""
    try:
        doc = ezdxf.readfile(dxf_path)
        msp = doc.modelspace()
        
        path = QPainterPath()
        for entity in msp:
            if entity.dxftype() == 'LWPOLYLINE':
                points = [(p[0] * scale + offset_x, p[1] * scale + offset_y) for p in entity.get_points('xy')]
                if points:
                    path.moveTo(points[0][0], points[0][1])
                    for i in range(1, len(points)):
                        path.lineTo(points[i][0], points[i][1])
                if entity.is_closed:
                    path.closeSubpath()
            elif entity.dxftype() == 'CIRCLE':
                center = entity.dxf.center
                radius = entity.dxf.radius
                cx, cy, r = center.x * scale + offset_x, center.y * scale + offset_y, radius * scale
                path.addEllipse(cx - r, cy - r, r * 2, r * 2)
        painter.drawPath(path)
    except (IOError, ezdxf.DXFStructureError) as e:
        logging.error(f"Erro ao ler ou desenhar DXF '{dxf_path}': {e}")

# Classe para desenhar a visualização do plano de corte
class CuttingPlanWidget(QWidget):
    def __init__(self, chapa_largura, chapa_altura, plano, color_map, parent=None):
        super().__init__(parent)
        self.chapa_largura, self.chapa_altura = chapa_largura, chapa_altura
        self.plano = plano
        self.color_map = color_map
        self.setMinimumSize(400, 600)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # Escala para caber no widget
        scale_w = self.width() / self.chapa_largura
        scale_h = self.height() / self.chapa_altura
        scale = min(scale_w, scale_h) * 0.95 # 5% de margem

        # Centraliza o desenho
        offset_x = (self.width() - (self.chapa_largura * scale)) / 2
        offset_y = (self.height() - (self.chapa_altura * scale)) / 2

        # 1. Desenha a chapa
        painter.setPen(QPen(QColor("#CCCCCC")))
        painter.setBrush(QBrush(QColor("#FFFFFF")))
        painter.drawRect(int(offset_x), int(offset_y), int(self.chapa_largura * scale), int(self.chapa_altura * scale))

        # 2. Desenha as peças
        painter.setPen(QPen(QColor("#333333")))
        for peca in self.plano:
            x, y, w, h, tipo_key = peca['x'], peca['y'], peca['largura'], peca['altura'], peca['tipo_key']
            # As coordenadas X, Y vêm com a origem no canto superior esquerdo.
            rect_x = int(offset_x + x * scale)
            rect_y = int(offset_y + y * scale)
            rect_w = int(w * scale)
            rect_h = int(h * scale)

            # --- INÍCIO: DESENHO CONDICIONAL (RETÂNGULO OU CÍRCULO) ---
            forma = peca.get('forma', 'rectangle')
            cor_peca = self.color_map.get(tipo_key, QColor("#A94442"))
            painter.setBrush(QBrush(cor_peca))

            if forma == 'circle':
                # Para círculos, desenha o círculo original dentro do seu bounding box alocado
                diametro_original = peca.get('diametro', 0)
                raio_desenhado = (diametro_original * scale) / 2
                centro_x = rect_x + rect_w / 2
                centro_y = rect_y + rect_h / 2
                painter.drawEllipse(int(centro_x - raio_desenhado), int(centro_y - raio_desenhado), int(raio_desenhado * 2), int(raio_desenhado * 2))
            elif forma == 'paired_triangle':
                # Desenha dois triângulos opostos para representar o par
                path1 = QPainterPath()
                path1.moveTo(rect_x, rect_y)
                path1.lineTo(rect_x + rect_w, rect_y)
                path1.lineTo(rect_x, rect_y + rect_h)
                path1.closeSubpath()
                path2 = QPainterPath()
                path2.moveTo(rect_x + rect_w, rect_y + rect_h)
                path2.lineTo(rect_x, rect_y + rect_h)
                path2.lineTo(rect_x + rect_w, rect_y)
                path2.closeSubpath()
                painter.drawPath(path1); painter.drawPath(path2)
            elif forma == 'paired_trapezoid':
                orig_dims = peca.get('orig_dims')
                if orig_dims:
                    large_base_scaled = orig_dims['large_base'] * scale
                    small_base_scaled = orig_dims['small_base'] * scale
                    height_scaled = orig_dims['height'] * scale
                    offset_x_trap = (large_base_scaled - small_base_scaled) / 2

                    # Trapézio 1 (normal)
                    path1 = QPainterPath()
                    path1.moveTo(rect_x, rect_y); path1.lineTo(rect_x + large_base_scaled, rect_y); path1.lineTo(rect_x + large_base_scaled - offset_x_trap, rect_y + height_scaled); path1.lineTo(rect_x + offset_x_trap, rect_y + height_scaled); path1.closeSubpath()
                    
                    # Trapézio 2 (rotacionado 180 graus e deslocado)
                    # --- CORREÇÃO DA LÓGICA DE DESENHO DO SEGUNDO TRAPÉZIO ---
                    path2 = QPainterPath()
                    # Ponto inferior esquerdo do 2º trapézio (coincide com inferior direito do 1º)
                    path2.moveTo(rect_x + large_base_scaled, rect_y)
                    # Ponto inferior direito do 2º trapézio (canto do bounding box)
                    path2.lineTo(rect_x + rect_w, rect_y)
                    # Ponto superior direito do 2º trapézio
                    path2.lineTo(rect_x + rect_w - offset_x_trap, rect_y + height_scaled)
                    # Ponto superior esquerdo do 2º trapézio (coincide com superior direito do 1º)
                    path2.lineTo(rect_x + large_base_scaled - offset_x_trap, rect_y + height_scaled)
                    path2.closeSubpath()
                    painter.drawPath(path1); painter.drawPath(path2)
            elif forma == 'dxf_shape':
                _draw_dxf_entities(painter, peca['dxf_path'], rect_x, rect_y, scale)
            else: # 'rectangle'
                painter.drawRect(rect_x, rect_y, rect_w, rect_h)
            # --- FIM: DESENHO CONDICIONAL ---

            # 3. Desenha os furos dentro da peça (se não for DXF, pois já podem estar no arquivo)
            furos = peca.get('furos', [])
            if furos:
                painter.setBrush(QBrush(QColor("#FFFFFF"))) # Furos brancos
                for furo in furos:
                    furo_x = rect_x + (furo['x'] * scale)
                    furo_y = rect_y + (furo['y'] * scale)
                    furo_diam = furo['diam'] * scale
                    # Desenha o círculo do furo centralizado na sua coordenada
                    painter.drawEllipse(int(furo_x - furo_diam / 2), int(furo_y - furo_diam / 2), int(furo_diam), int(furo_diam))

        # 4. Desenha as sobras
        sobras = self.parent().plano_sobras # Pega as sobras do diálogo pai
        if sobras:
            font = painter.font()
            font.setPointSize(7)
            painter.setFont(font)

            for sobra in sobras:
                if sobra.get('tipo_sobra') == 'aproveitavel':
                    painter.setBrush(QBrush(QColor(100, 100, 100, 150))) # Cinza escuro semi-transparente
                    painter.setPen(QPen(QColor("#333333"), 1, Qt.DashLine))
                else:
                    painter.setBrush(QBrush(QColor(230, 230, 230, 120))) # Cinza claro semi-transparente
                    painter.setPen(QPen(QColor("#888888"), 1, Qt.DashLine))
                # As coordenadas da sobra já vêm com a origem no topo, igual às peças.
                rect_x, rect_y = int(offset_x + sobra['x'] * scale), int(offset_y + sobra['y'] * scale)
                rect_w, rect_h = int(sobra['largura'] * scale), int(sobra['altura'] * scale)
                painter.drawRect(rect_x, rect_y, rect_w, rect_h)
                painter.drawText(rect_x + 4, rect_y + 12, f"{sobra['largura']:.0f}x{sobra['altura']:.0f}")

class PlanVisualizationDialog(QDialog):
    def __init__(self, chapa_largura, chapa_altura, plano_info, offset, color_map, parent=None):
        super().__init__(parent)
        self.chapa_largura = chapa_largura
        self.chapa_altura = chapa_altura
        self.plano = plano_info['plano']
        self.repeticoes = plano_info['repeticoes']
        # --- CORREÇÃO: Acessa a lista de sobras corretamente ---
        # O cálculo agora retorna uma lista de dicionários, não um dicionário com chaves.
        sobras_raw = plano_info.get('sobras', [])
        self.plano_sobras = sobras_raw if isinstance(sobras_raw, list) else []
        self.resumo_pecas = plano_info['resumo_pecas']
        self.offset = offset
        self.color_map = color_map
        self.setWindowTitle("Visualização Detalhada do Plano de Corte")
        self.setMinimumSize(600, 800)
        
        layout = QVBoxLayout(self)

        # --- INÍCIO: CONTAINER PARA DETALHES ---
        # Agrupa todos os QGroupBox de detalhes em um único widget para facilitar o toggle
        self.details_container = QWidget()
        details_layout = QVBoxLayout(self.details_container)
        details_layout.setContentsMargins(0, 0, 0, 0)
        # --- INÍCIO: NOVOS LABELS DE INFORMAÇÃO ---
        info_group = QGroupBox("Detalhes do Plano")
        info_layout = QVBoxLayout()

        chapa_label = QLabel(f"<b>Dimensões da Chapa:</b> {self.chapa_largura} x {self.chapa_altura} mm")
        info_layout.addWidget(chapa_label)

        repeticoes_label = QLabel(f"<b>Repetições deste Plano:</b> {self.repeticoes}x")
        info_layout.addWidget(repeticoes_label)

        pecas_label_titulo = QLabel("<b>Peças neste plano:</b>")
        info_layout.addWidget(pecas_label_titulo)

        for item in self.resumo_pecas:
            # --- CORREÇÃO: Exibe a identificação da peça diretamente ---
            texto_peca = f"- {item['qtd']}x de {item['tipo']}"
            info_layout.addWidget(QLabel(texto_peca))
        info_group.setLayout(info_layout)
        details_layout.addWidget(info_group)

        # --- INÍCIO: SEPARAÇÃO DAS SOBRAS NA VISUALIZAÇÃO ---
        if self.plano_sobras:
            sobras_aproveitaveis = [s for s in self.plano_sobras if s.get('tipo_sobra') == 'aproveitavel']
            sobras_sucata = [s for s in self.plano_sobras if s.get('tipo_sobra') != 'aproveitavel']

            if sobras_aproveitaveis:
                aprov_group = QGroupBox("Sobras Aproveitáveis")
                aprov_layout = QVBoxLayout()
                for i, sobra in enumerate(sobras_aproveitaveis):
                    aprov_layout.addWidget(QLabel(f"- Retalho {i+1}: {sobra['largura']:.0f} x {sobra['altura']:.0f} mm"))
                aprov_group.setLayout(aprov_layout)
                details_layout.addWidget(aprov_group)
            
            if sobras_sucata:
                sucata_group = QGroupBox("Sobras (Sucata)")
                sucata_layout = QVBoxLayout()
                for i, sobra in enumerate(sobras_sucata):
                    sucata_layout.addWidget(QLabel(f"- Retalho {i+1}: {sobra['largura']:.0f} x {sobra['altura']:.0f} mm"))
                sucata_group.setLayout(sucata_layout)
                details_layout.addWidget(sucata_group)
        # --- FIM: NOVOS LABELS DE INFORMAÇÃO ---

        layout.addWidget(self.details_container)
        # --- FIM: CONTAINER PARA DETALHES ---

        cutting_widget = CuttingPlanWidget(chapa_largura, chapa_altura, self.plano, color_map)
        layout.addWidget(cutting_widget)

        buttons_layout = QHBoxLayout()
        btn_export_pdf = QPushButton("Exportar para PDF")
        btn_export_pdf.clicked.connect(self.export_to_pdf)

        # --- INÍCIO: BOTÃO PARA OCULTAR/MOSTRAR DETALHES ---
        self.toggle_details_btn = QPushButton("Ocultar Detalhes")
        self.toggle_details_btn.clicked.connect(self.toggle_details_visibility)
        # --- FIM: BOTÃO PARA OCULTAR/MOSTRAR DETALHES ---

        btn_close = QPushButton("Fechar")
        btn_close.clicked.connect(self.accept)

        buttons_layout.addWidget(self.toggle_details_btn)
        buttons_layout.addWidget(btn_export_pdf)
        buttons_layout.addStretch()
        buttons_layout.addWidget(btn_close)
        layout.addLayout(buttons_layout)

        # --- MELHORIA: Oculta os detalhes por padrão ---
        self.details_container.setVisible(False)
        self.toggle_details_btn.setText("Mostrar Detalhes")

    def toggle_details_visibility(self):
        is_visible = self.details_container.isVisible()
        self.details_container.setVisible(not is_visible)
        self.toggle_details_btn.setText("Mostrar Detalhes" if is_visible else "Ocultar Detalhes")

    def export_to_pdf(self):
        default_path = os.path.join(os.path.expanduser("~"), "Downloads", "Plano_de_Corte.pdf")
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar PDF do Plano de Corte", default_path, "PDF Files (*.pdf)")
        if save_path:
            c = canvas.Canvas(save_path, pagesize=A4)
            pdf_generator.gerar_pdf_plano_de_corte(c, self.chapa_largura, self.chapa_altura, self.plano, self.color_map)
            c.save()
            QMessageBox.information(self, "Sucesso", f"PDF do plano de corte salvo em:\n{save_path}")


class NestingDialog(QDialog):
    def __init__(self, dataframe, parent=None):
        super().__init__(parent)
        self.df = dataframe
        self.calculation_results = None # Armazena os resultados completos
        self.color_map = {} # Armazena o mapa de cores por tipo de peça
        self.setWindowTitle("Cálculo de Aproveitamento de Chapa")
        self.setMinimumWidth(600)
        self.resize(800, 700) # Define um tamanho inicial maior

        # Layout principal
        self.main_layout = QVBoxLayout(self)

        # --- INÍCIO: BOTÃO DE EXPANSÃO ---
        title_bar_layout = QHBoxLayout()
        title_bar_layout.addStretch()
        self.toggle_fullscreen_btn = QPushButton("🗖") # Símbolo de maximizar/restaurar
        self.toggle_fullscreen_btn.setFixedSize(30, 30)
        self.toggle_fullscreen_btn.setToolTip("Maximizar / Restaurar Janela")
        self.toggle_fullscreen_btn.clicked.connect(self.toggle_fullscreen)
        title_bar_layout.addWidget(self.toggle_fullscreen_btn)
        self.main_layout.addLayout(title_bar_layout)
        # --- FIM: BOTÃO DE EXPANSÃO ---

        # --- Grupo de Inputs do Usuário ---
        # ... (código existente)
        input_group = QGroupBox("Parâmetros")
        form_layout = QFormLayout()
        self.chapa_largura_input = QLineEdit("3000")
        self.chapa_altura_input = QLineEdit("1500")
        self.offset_input = QLineEdit("8")
        self.margin_input = QLineEdit("10")
        self.method_combo = QComboBox()
        self.method_combo.addItems(["Plasma/Laser", "Guilhotina"])
        form_layout.addRow("Largura da Chapa (mm):", self.chapa_largura_input)
        form_layout.addRow("Altura da Chapa (mm):", self.chapa_altura_input)
        form_layout.addRow("Método de Corte:", self.method_combo)
        form_layout.addRow("Offset entre Peças (mm):", self.offset_input)
        form_layout.addRow("Margem da Chapa (mm):", self.margin_input) # <<< NOVA LINHA
        input_group.setLayout(form_layout)
        self.main_layout.addWidget(input_group)
        
        # Botões de Ação
        action_layout = QHBoxLayout()
        self.calculate_btn = QPushButton("Calcular")
        self.calculate_btn.clicked.connect(self.run_calculation)
        self.export_report_btn = QPushButton("Exportar Relatório (PDF)")
        self.export_dxf_btn = QPushButton("Exportar Planos (DXF)") # Novo botão
        self.export_report_btn.clicked.connect(self.export_full_report_to_pdf)
        self.export_report_btn.setEnabled(False) # Desabilitado até o cálculo ser feito
        self.export_dxf_btn.setEnabled(False) # Desabilitado até o cálculo ser feito
        self.export_dxf_btn.clicked.connect(self.export_layouts_to_dxf)
        action_layout.addWidget(self.calculate_btn)
        action_layout.addWidget(self.export_report_btn)
        action_layout.addWidget(self.export_dxf_btn)
        self.main_layout.addLayout(action_layout)

        # --- Área de Resultados ---
        self.status_label = QLabel("Clique em 'Calcular' para iniciar.")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-style: italic; color: #888888;")
        self.main_layout.addWidget(self.status_label)

        results_group = QGroupBox("Resultados")
        self.results_layout = QVBoxLayout()
        results_group.setLayout(self.results_layout)
        
        # Adiciona uma área de rolagem para os resultados
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        self.results_scroll_layout = QVBoxLayout(scroll_content)
        scroll.setWidget(scroll_content)

        self.main_layout.addWidget(results_group)
        self.main_layout.addWidget(scroll)

    def toggle_fullscreen(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def run_calculation(self):
        try:
            chapa_largura = float(self.chapa_largura_input.text())
            chapa_altura = float(self.chapa_altura_input.text())
            offset = float(self.offset_input.text())
            margin = float(self.margin_input.text())
            method = self.method_combo.currentText()
        except ValueError:
            QMessageBox.critical(self, "Erro de Entrada", "Por favor, insira valores numéricos válidos.")
            return

        # Filtra apenas peças retangulares e agrupa por espessura
        # --- MUDANÇA: Inclui círculos no filtro ---
        valid_shapes_df = self.df[self.df['forma'].isin(['rectangle', 'circle', 'right_triangle', 'trapezoid', 'dxf_shape'])].copy()
        valid_shapes_df['espessura'] = valid_shapes_df['espessura'].astype(float)
        grouped = valid_shapes_df.groupby('espessura')

        if len(grouped) == 0: # Atualize a mensagem de erro
            QMessageBox.information(self, "Nenhuma Peça", "Nenhuma peça válida (retângulo, círculo, triângulo, trapézio ou DXF) encontrada para o cálculo.")
            return

        # --- INÍCIO: LÓGICA DA THREAD ---
        # 1. Prepara a UI para o cálculo
        self.prepare_for_calculation()

        # 2. Cria e inicia a thread
        self.thread = CalculationThread(chapa_largura, chapa_altura, offset, margin, method, grouped)
        self.thread.result_ready.connect(self.on_result_ready)
        self.thread.finished.connect(self.on_calculation_finished)
        self.thread.error.connect(self.on_calculation_error)
        self.thread.status_update.connect(self.on_status_update) # Conecta o novo sinal
        self.thread.start()
        # --- FIM: LÓGICA DA THREAD ---

    def prepare_for_calculation(self):
        """Limpa a UI e a prepara para receber novos resultados."""
        # Limpa resultados anteriores
        for i in reversed(range(self.results_scroll_layout.count())): 
            widget = self.results_scroll_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        
        self.calculation_results = {}
        self.export_report_btn.setEnabled(False)
        self.calculate_btn.setEnabled(False)
        self.calculate_btn.setText("Calculando...")
        self.status_label.setText("Calculando, por favor aguarde...")
        self.status_label.setStyleSheet("font-style: normal; color: #FFFFFF;")

    def on_result_ready(self, espessura, resultado):
        """Slot para receber o resultado de uma espessura e atualizar a UI."""
        chapa_largura = float(self.chapa_largura_input.text())
        chapa_altura = float(self.chapa_altura_input.text())
        self.calculation_results[espessura] = resultado
        self.display_results_for_thickness(espessura, resultado, chapa_largura, chapa_altura)

    def on_calculation_error(self, title, message):
        """Slot para exibir uma mensagem de erro vinda da thread."""
        QMessageBox.critical(self, title, message)

    def on_status_update(self, message):
        """Slot para receber atualizações de status da thread de cálculo."""
        self.status_label.setText(message)

    def on_calculation_finished(self):
        """Slot chamado quando todos os cálculos terminam."""
        self.calculate_btn.setEnabled(True)
        self.calculate_btn.setText("Calcular")
        self.status_label.setText("Cálculo concluído.")
        self.status_label.setStyleSheet("font-style: italic; color: #4CAF50;") # Verde
        # Habilita o botão de exportar se houver resultados
        if self.calculation_results:
            self.export_report_btn.setEnabled(True)
            self.export_dxf_btn.setEnabled(True)


    def export_full_report_to_pdf(self):
        if not self.calculation_results:
            QMessageBox.warning(self, "Sem Dados", "Nenhum resultado de cálculo para exportar. Por favor, clique em 'Calcular' primeiro.")
            return

        # --- INÍCIO: CÁLCULO DA ÁREA REAL DAS PEÇAS ---
        import math
        df_copy = self.df.copy()
        df_copy['espessura'] = df_copy['espessura'].astype(float)
        grouped_by_thickness = df_copy.groupby('espessura')

        total_piece_areas = {}
        for espessura, group in grouped_by_thickness:
            total_area = 0
            for _, row in group.iterrows():
                forma = row.get('forma', 'rectangle')
                qtd = row.get('qtd', 1)
                
                piece_area = 0
                if forma == 'rectangle':
                    piece_area = row.get('largura', 0) * row.get('altura', 0)
                elif forma == 'circle':
                    piece_area = math.pi * (row.get('diametro', 0) / 2)**2
                elif forma == 'right_triangle':
                    piece_area = (row.get('rt_base', 0) * row.get('rt_height', 0)) / 2
                elif forma == 'trapezoid':
                    b1 = row.get('trapezoid_large_base', 0)
                    b2 = row.get('trapezoid_small_base', 0)
                    h = row.get('trapezoid_height', 0)
                    piece_area = ((b1 + b2) * h) / 2
                elif forma == 'dxf_shape':
                    piece_area = row.get('largura', 0) * row.get('altura', 0)
                    
                total_area += piece_area * qtd
            total_piece_areas[espessura] = total_area
        # --- FIM: CÁLCULO DA ÁREA REAL DAS PEÇAS ---

        default_path = os.path.join(os.path.expanduser("~"), "Downloads", "Relatorio_Aproveitamento.pdf")
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório de Aproveitamento", default_path, "PDF Files (*.pdf)")
        
        if save_path:
            # --- INÍCIO: CÁLCULO DO OFFSET (LÓGICA CORRIGIDA) ---
            for espessura, resultado in self.calculation_results.items():
                area_total_chapas = resultado.get('area_total_chapas', 0)
                area_real_pecas = total_piece_areas.get(espessura, 0)
                area_sobra_aproveitavel = resultado.get('total_area_sobra_aproveitavel', 0)
                area_sobra_sucata = resultado.get('total_area_sobra_sucata', 0)

                offset_area = area_total_chapas - (area_real_pecas + area_sobra_aproveitavel + area_sobra_sucata)
                if offset_area < 0:
                    offset_area = 0

                offset_weight = (offset_area / 1_000_000) * espessura * 7.85

                resultado['offset_area'] = offset_area
                resultado['offset_weight'] = offset_weight
            # --- FIM: CÁLCULO DO OFFSET (LÓGICA CORRIGIDA) ---

            c = canvas.Canvas(save_path, pagesize=A4)
            pdf_generator.gerar_relatorio_completo_pdf(c, self.calculation_results, float(self.chapa_largura_input.text()), float(self.chapa_altura_input.text()))
            c.save()
            QMessageBox.information(self, "Sucesso", f"Relatório PDF salvo em:\n{save_path}")

    def export_layouts_to_dxf(self):
        """Exporta todos os planos de corte calculados para um único arquivo DXF."""
        if not self.calculation_results:
            QMessageBox.warning(self, "Sem Dados", "Nenhum resultado de cálculo para exportar.")
            return

        default_path = os.path.join(os.path.expanduser("~"), "Downloads", "Aproveitamento_Completo.dxf")
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Layouts em DXF", default_path, "DXF Files (*.dxf)")

        if not save_path:
            return

        try:
            doc = ezdxf.new('R2010')
            msp = doc.modelspace()
            chapa_w = float(self.chapa_largura_input.text())
            chapa_h = float(self.chapa_altura_input.text())
            margin = float(self.margin_input.text())
            x_offset = 0

            for espessura, resultado in self.calculation_results.items():
                for plano_info in resultado['planos_unicos']:
                    for i in range(plano_info['repeticoes']):
                        # Desenha o contorno da chapa
                        msp.add_lwpolyline([(x_offset, 0), (x_offset + chapa_w, 0), (x_offset + chapa_w, chapa_h), (x_offset, chapa_h)], close=True, dxfattribs={'layer': 'CONTORNO_CHAPA'})

                        # Desenha as peças
                        for peca in plano_info['plano']:
                            layer_name = peca['tipo_key'].replace(' ', '_').replace('Ø', 'D').replace('/', '_')
                            if layer_name not in doc.layers:
                                doc.layers.new(name=layer_name)
                            
                            px, py, pw, ph = peca['x'], peca['y'], peca['largura'], peca['altura']
                            
                            # Converte Y para o sistema de coordenadas do DXF (origem embaixo)
                            # A coordenada 'py' já vem com a origem no topo, então a conversão é direta.
                            py_dxf = chapa_h - py - ph 
                            forma = peca.get('forma', 'rectangle')
                            if forma == 'rectangle':
                                msp.add_lwpolyline([(x_offset + px, py_dxf), (x_offset + px + pw, py_dxf), (x_offset + px + pw, py_dxf + ph), (x_offset + px, py_dxf + ph)], close=True, dxfattribs={'layer': layer_name})
                            elif forma == 'circle':
                                cx = x_offset + px + pw / 2
                                cy = py_dxf + ph / 2
                                msp.add_circle((cx, cy), radius=peca['diametro']/2, dxfattribs={'layer': layer_name})
                            elif forma == 'paired_triangle':
                                x, y = x_offset + px, py_dxf
                                msp.add_lwpolyline([(x, y), (x + pw, y), (x, y + ph)], close=True, dxfattribs={'layer': layer_name})
                                msp.add_lwpolyline([(x + pw, y + ph), (x, y + ph), (x + pw, y)], close=True, dxfattribs={'layer': layer_name})
                            elif forma == 'paired_trapezoid':
                                dims = peca['orig_dims']
                                l_base, s_base, h = dims['large_base'], dims['small_base'], dims['height']
                                x_trap_offset = (l_base - s_base) / 2
                                x, y = x_offset + px, py_dxf
                                # Trapézio 1
                                msp.add_lwpolyline([(x, y), (x + l_base, y), (x + l_base - x_trap_offset, y + h), (x + x_trap_offset, y + h)], close=True, dxfattribs={'layer': layer_name})
                                # Trapézio 2
                                msp.add_lwpolyline([(x + l_base, y), (x + pw, y), (x + pw - x_trap_offset, y + h), (x + l_base, y + h)], close=True, dxfattribs={'layer': layer_name})

                        x_offset += chapa_w + 100 # Espaço entre as chapas

            doc.saveas(save_path)
            QMessageBox.information(self, "Sucesso", f"Layouts DXF salvos em:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro na Exportação DXF", f"Ocorreu um erro ao gerar o arquivo DXF:\n{e}")

    def display_results_for_thickness(self, espessura, resultado, chapa_w, chapa_h):
        # Cria um grupo para cada espessura
        group_box = QGroupBox(f"Espessura: {espessura} mm")
        group_layout = QVBoxLayout()
        
        # Adiciona informações gerais
        info_label = QLabel(f"Total de Chapas: {resultado['total_chapas']} | Aproveitamento Geral: {resultado['aproveitamento_geral']}")
        info_label.setStyleSheet("font-weight: bold;")
        group_layout.addWidget(info_label)

        # --- INÍCIO: CORREÇÃO DO ERRO "INDEX OUT OF RANGE" ---
        # Verifica se existem planos de corte antes de tentar acessá-los.
        # Isso acontece se nenhuma peça couber na chapa.
        if resultado.get('planos_unicos'):
            # --- INÍCIO: GERAÇÃO DO MAPA DE CORES ---
            # Gera o mapa de cores aqui, pois é específico para cada grupo de espessura
            offset = float(self.offset_input.text())
            tipos_de_peca_unicos = [p['tipo'] for p in resultado['planos_unicos'][0]['resumo_pecas']]
            cores = generate_distinct_colors(len(tipos_de_peca_unicos))
            self.color_map = {tipo: cor for tipo, cor in zip(tipos_de_peca_unicos, cores)}
            # --- CORREÇÃO: Armazena o mapa de cores DENTRO do resultado da espessura específica ---
            resultado['color_map'] = self.color_map
            # --- FIM: GERAÇÃO DO MAPA DE CORES ---

            # Lista os planos de corte únicos
            for i, plano_info in enumerate(resultado['planos_unicos']):
                plano_layout = QHBoxLayout()
                
                resumo_pecas_str = ", ".join([f"{p['qtd']}x ({p['tipo']})" for p in plano_info['resumo_pecas']])
                
                plan_label = QLabel(f"Plano {i+1}: {plano_info['repeticoes']}x | Peças: {resumo_pecas_str}")
                
                view_btn = QPushButton("Ver Detalhes")
                # --- MUDANÇA: Passa o dicionário 'plano_info' completo e o offset ---
                view_btn.clicked.connect(lambda _, p_info=plano_info, w=chapa_w, h=chapa_h: self.show_plan_visualization(p_info, w, h, self.color_map))
                
                plano_layout.addWidget(plan_label)
                plano_layout.addStretch()
                plano_layout.addWidget(view_btn)
                
                group_layout.addLayout(plano_layout)
        else:
            # Caso não haja planos, exibe uma mensagem informativa.
            no_fit_label = QLabel("Nenhuma peça coube na chapa com as dimensões fornecidas.")
            no_fit_label.setStyleSheet("color: #FDBA74; font-style: italic;") # Laranja claro
            group_layout.addWidget(no_fit_label)
        # --- FIM: CORREÇÃO ---

        group_box.setLayout(group_layout)
        self.results_scroll_layout.addWidget(group_box)

    def show_plan_visualization(self, plano_info, chapa_w, chapa_h, color_map):
        offset = float(self.offset_input.text())
        dialog = PlanVisualizationDialog(chapa_w, chapa_h, plano_info, offset, color_map, self)
        dialog.exec_()