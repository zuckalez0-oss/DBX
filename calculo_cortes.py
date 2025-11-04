import rectpack
import logging
from PyQt5.QtCore import QObject, pyqtSignal
from rectpack.maxrects import MaxRectsBssf, MaxRectsBaf, MaxRectsBlsf, MaxRectsBl
from rectpack.skyline import SkylineBl, SkylineBlWm, SkylineMwf, SkylineMwfl
import math
import os

# --- INÍCIO: CONFIGURAÇÃO DE LOGGING PARA DEBUG ---
logging.basicConfig(filename='debug_nesting.log', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s', filemode='w')
# --- FIM: CONFIGURAÇÃO DE LOGGING ---

# --- INÍCIO: CLASSE PARA EMITIR SINAIS DE STATUS ---
class StatusSignaler(QObject):
    status_update = pyqtSignal(str)

status_signaler = StatusSignaler()
# --- FIM: CLASSE PARA EMITIR SINAIS DE STATUS ---

def _merge_scraps(scraps):
    """
    Função auxiliar para fundir retângulos de sobras adjacentes.
    Esta versão é mais robusta e corrige falhas na lógica de fusão horizontal.
    """
    if not scraps:
        return []

    TOLERANCE = 1e-5 

    while True:
        merged_in_pass = False
        i = 0
        while i < len(scraps):
            j = i + 1
            while j < len(scraps):
                r1 = scraps[i]
                r2 = scraps[j]
                
                merged = False

                # Tenta fusão vertical
                if abs(r1['x'] - r2['x']) < TOLERANCE and abs(r1['largura'] - r2['largura']) < TOLERANCE:
                    if abs((r1['y'] + r1['altura']) - r2['y']) < TOLERANCE:
                        r1['altura'] += r2['altura']
                        merged = True
                    elif abs((r2['y'] + r2['altura']) - r1['y']) < TOLERANCE:
                        r1['y'] = r2['y']
                        r1['altura'] += r2['altura']
                        merged = True

                # Tenta fusão horizontal
                elif abs(r1['y'] - r2['y']) < TOLERANCE and abs(r1['altura'] - r2['altura']) < TOLERANCE:
                    if abs((r1['x'] + r1['largura']) - r2['x']) < TOLERANCE:
                        r1['largura'] += r2['largura']
                        merged = True
                    elif abs((r2['x'] + r2['largura']) - r1['x']) < TOLERANCE:
                        # r1 está à direita de r2, expande r1 para a esquerda
                        r1['x'] = r2['x']
                        r1['largura'] += r2['largura']
                        merged = True

                if merged:
                    scraps.pop(j)
                    merged_in_pass = True # Não incrementa 'j' para que o próximo item
                    # (que agora está no índice 'j') seja verificado na próxima iteração do loop.
                else:
                    j += 1
            i += 1
        
        if not merged_in_pass:
            break
            
    return scraps

def encontrar_sobras(chapa_largura, chapa_altura, pecas_alocadas, min_dim=50):
    """
    Encontra os maiores retângulos de sobra em uma chapa.
    Este método usa um algoritmo de varredura (Sweep-line) para robustez,
    seguido por uma etapa de fusão para consolidar os resultados.
    """
    logging.debug("Iniciando 'encontrar_sobras' com algoritmo de varredura (Scanline).")

    # 1. Coletar todas as coordenadas Y únicas (eventos da linha de varredura).
    y_coords = {0, chapa_altura}
    for p in pecas_alocadas:
        y_coords.add(p['y'])
        y_coords.add(p['y'] + p['altura'])
    
    sorted_y = sorted(list(y_coords))
    sobras_brutas = []

    # 2. Iterar sobre cada faixa horizontal (strip) criada pelas coordenadas Y.
    for i in range(len(sorted_y) - 1):
        y1, y2 = sorted_y[i], sorted_y[i+1]
        altura_faixa = y2 - y1

        if altura_faixa < 1e-5:
            continue

        # Encontra todas as peças que se sobrepõem a esta faixa horizontal.
        mid_y = y1 + altura_faixa / 2
        pecas_na_faixa = [p for p in pecas_alocadas if p['y'] <= mid_y < (p['y'] + p['altura'])]
        
        # 3. Ordena as peças na faixa pelo eixo X para encontrar os vãos.
        pecas_na_faixa.sort(key=lambda p: p['x'])

        # 4. Encontra os espaços vazios (vãos) no eixo X.
        ponteiro_x = 0
        for p in pecas_na_faixa:
            if p['x'] > ponteiro_x:
                largura_sobra = p['x'] - ponteiro_x
                if largura_sobra > 1e-5:
                    sobras_brutas.append({
                        'x': ponteiro_x, 'y': y1,
                        'largura': largura_sobra, 'altura': altura_faixa
                    })
            ponteiro_x = max(ponteiro_x, p['x'] + p['largura'])
        
        # 5. Verifica se há um último vão após a última peça até a borda da chapa.
        if ponteiro_x < chapa_largura:
            largura_sobra = chapa_largura - ponteiro_x
            if largura_sobra > 1e-5:
                sobras_brutas.append({
                    'x': ponteiro_x, 'y': y1,
                    'largura': largura_sobra, 'altura': altura_faixa
                })

    # 6. Fundir os retângulos de sobra adjacentes para formar peças maiores.
    sobras_fundidas = _merge_scraps(sobras_brutas)

    # 7. Filtrar pelo tamanho mínimo e classificar
    sobras_finais = []
    for s in sobras_fundidas:
        if s['largura'] >= min_dim and s['altura'] >= min_dim:
            s['y'] = chapa_altura - s['y'] - s['altura']
            
            s['tipo_sobra'] = 'aproveitavel' if min(s['largura'], s['altura']) >= 300 and max(s['largura'], s['altura']) >= 300 else 'nao_aproveitavel'
            
            # ATRIBUIÇÃO DE PONTUAÇÃO DE REUTILIZAÇÃO
            score = 0
            area = s['largura'] * s['altura']
            
            if area > (chapa_largura * chapa_altura * 0.1):
                score += 10
            
            if s['tipo_sobra'] == 'aproveitavel':
                score += 5
            
            if (s['largura'] >= chapa_largura * 0.9 and s['altura'] >= 100) or \
               (s['altura'] >= chapa_altura * 0.9 and s['largura'] >= 100):
                score += 7
            
            s['potential_reuse_score'] = score
            sobras_finais.append(s)

    logging.debug(f"Finalizado 'encontrar_sobras' (Scanline). Encontradas {len(sobras_finais)} sobras válidas.")
    return sobras_finais

def orquestrar_planos_de_corte(chapa_largura, chapa_altura, pecas, offset, margin, espessura, is_guillotine=False, peso_especifico_base=7.85, status_signal_emitter=None):
    """
    Função mestre que orquestra o processo de nesting.
    """
    logging.info(f"--- INICIANDO ORQUESTRAÇÃO DE NESTING (ESTRATÉGIA OTIMIZADA) PARA ESPESSURA {espessura}mm ---")

    pecas_ordenadas = sorted(pecas, key=lambda p: p['largura'] * p['altura'], reverse=True)
    logging.info(f"Ordenando {len(pecas_ordenadas)} tipos de peças por área para otimização.")

    bins_disponiveis = [(chapa_largura, chapa_altura, margin)] * 200

    if status_signal_emitter: status_signal_emitter.emit("Otimizando alocação de todas as peças...")
    resultado_otimizado = calcular_plano_de_corte_em_bins(
        pecas_ordenadas, 
        offset, 
        espessura, 
        is_guillotine,
        bins_disponiveis, 
        peso_especifico_base, 
        status_signal_emitter
    )

    logging.info(f"--- ORQUESTRAÇÃO OTIMIZADA FINALIZADA ---")
    return resultado_otimizado


def calcular_plano_de_corte_em_bins(pecas, offset, espessura, is_guillotine, bins, peso_especifico_base=7.85, status_signal_emitter=None):
    """
    Calcula o plano de corte, incluindo uma análise detalhada de pesos e sucatas.
    """
    logging.info(f"Iniciando cálculo de corte para {len(pecas)} tipos de peças em {len(bins)} bins disponíveis.")

    def _calc_peso(area_mm2):
        if espessura is None or espessura <= 0:
            return 0
        return (area_mm2 / 1_000_000) * espessura * peso_especifico_base
    
    # LÓGICA DE FUSÃO DE FORMAS (Triângulos, Trapézios)
    pecas_processadas = []
    trapezios = [p for p in pecas if p.get('forma') == 'trapezoid']
    outras_pecas = [p for p in pecas if p.get('forma') not in ['trapezoid']]

    from collections import defaultdict

    mapa_trapezios = defaultdict(list)
    for t in trapezios:
        mapa_trapezios[(t['largura'], t['altura'], t.get('small_base', 0))].append(t)

    for dim, lista_trapezios in mapa_trapezios.items():
        total_qtd = sum(t['quantidade'] for t in lista_trapezios)
        num_pares, num_sozinhos = divmod(total_qtd, 2)
        if num_pares > 0:
            pecas_processadas.append({
                'forma': 'paired_trapezoid', 'largura': dim[0] + dim[2], 'altura': dim[1], 'quantidade': num_pares,
                'orig_dims': {'large_base': dim[0], 'small_base': dim[2], 'height': dim[1]}
            })
        if num_sozinhos > 0:
            pecas_processadas.append(lista_trapezios[0])

    pecas_processadas.extend(outras_pecas)
    
    peca_unica_id = 1
    retangulos_para_alocar = []
    id_peca_map = {}

    for peca_proc in pecas_processadas:
        largura_com_offset = peca_proc['largura']
        altura_com_offset = peca_proc['altura']
        largura_sem_offset = largura_com_offset - offset if largura_com_offset > offset else largura_com_offset
        altura_sem_offset = altura_com_offset - offset if altura_com_offset > offset else altura_com_offset
        
        for _ in range(peca_proc['quantidade']):
            rid = str(peca_unica_id)
            retangulos_para_alocar.append((largura_com_offset, altura_com_offset, rid))
            
            id_peca_map[rid] = {
                'largura_com_offset': largura_com_offset, 'altura_com_offset': altura_com_offset,
                'largura_sem_offset': largura_sem_offset, 'altura_sem_offset': altura_sem_offset,
                'furos': peca_proc.get('furos', []),
                'forma': peca_proc.get('forma', 'rectangle'),
                'diametro': peca_proc.get('diametro', 0),
                'orig_dims': peca_proc.get('orig_dims'),
                'dxf_path': peca_proc.get('dxf_path')
            }
            peca_unica_id += 1

    # ESTRATÉGIA DE BUSCA OTIMIZADA
    todos_algoritmos = [MaxRectsBssf, MaxRectsBaf, MaxRectsBlsf, SkylineBl, SkylineMwf, SkylineBlWm, SkylineMwfl]
    melhor_resultado_final = None
    num_pecas = len(retangulos_para_alocar)

    for num_bins_to_try in range(1, len(bins) + 1):
        if status_signal_emitter: status_signal_emitter.emit(f"Tentando alocar em {num_bins_to_try} chapa(s)...")
        logging.info(f"--- TENTATIVA COM {num_bins_to_try} CHAPAS ---")

        solucoes_validas_nesta_iteracao = []

        for algo in todos_algoritmos:
            packer = rectpack.newPacker(rotation=True, pack_algo=algo)
            for r in retangulos_para_alocar:
                packer.add_rect(r[0], r[1], rid=r[2])

            for i in range(num_bins_to_try):
                b_width, b_height, b_margin = bins[i]
                nesting_width = b_width - (2 * b_margin)
                nesting_height = b_height - (2 * b_margin)
                if nesting_width > 0 and nesting_height > 0:
                    packer.add_bin(nesting_width, nesting_height, bid=(b_width, b_height, b_margin))
            
            packer.pack()

            num_alocadas = sum(len(b) for b in packer)
            if num_alocadas == num_pecas:
                logging.info(f"SUCESSO! Algoritmo '{algo.__name__}' conseguiu alocar todas as peças em {num_bins_to_try} chapas.")
                
                planos_agrupados = {}
                area_total_utilizada_com_offset = 0

                for bin_node in packer:
                    if not bin_node: continue # Pula bins vazios
                    chapa_largura, chapa_altura, margin = bin_node.bid
                    nesting_width, nesting_height = chapa_largura - (2 * margin), chapa_altura - (2 * margin)
                    chapa_alocada = [r for r in bin_node if hasattr(r, 'x')]
                    assinatura = (chapa_largura, chapa_altura) + tuple(sorted([(r.x, r.y, r.width, r.height) for r in chapa_alocada]))

                    if assinatura not in planos_agrupados:
                        plano_de_corte, pecas_contagem = [], {}
                        for r in chapa_alocada:
                            peca_info = id_peca_map[r.rid]
                            forma = peca_info.get('forma', 'rectangle')
                            if forma == 'rectangle': tipo_key = f"R {peca_info['largura_sem_offset']:.0f}x{peca_info['altura_sem_offset']:.0f}"
                            elif forma == 'circle': tipo_key = f"C Ø{peca_info['diametro']:.0f}"
                            elif forma == 'paired_trapezoid': # A lógica de trapézio pode ser mantida por enquanto
                                dims = peca_info['orig_dims']
                                tipo_key = f"2Z {dims['large_base']-offset:.0f}/{dims['small_base']-offset:.0f}x{dims['height']-offset:.0f}"
                            elif forma == 'dxf_shape': tipo_key = f"DXF: {os.path.basename(peca_info['dxf_path'])}"
                            else: tipo_key = f"{forma[0].upper()} {peca_info['largura_sem_offset']:.0f}x{peca_info['altura_sem_offset']:.0f}"
                            pecas_contagem[tipo_key] = pecas_contagem.get(tipo_key, 0) + 1

                            furos_trans = []
                            if r.width != peca_info['largura_com_offset']:
                                for furo in peca_info['furos']: furos_trans.append({'diam': furo['diam'], 'x': furo['y'], 'y': peca_info['largura_com_offset'] - furo['x']})
                            else: furos_trans = peca_info['furos']

                            plano_de_corte.append({
                                "x": margin + r.x + (offset / 2),
                                "y": (chapa_altura - margin) - (r.y + r.height) + (offset / 2), # Invertido para origem no topo
                                "largura": r.width - offset, "altura": r.height - offset,
                                "tipo_key": tipo_key, "furos": furos_trans, "forma": forma, "rid": r.rid, "diametro": peca_info['diametro'],
                                "orig_dims": peca_info.get('orig_dims'), "dxf_path": peca_info['dxf_path']
                            })
                        
                        resumo_pecas = [{"tipo": t, "qtd": q} for t, q in pecas_contagem.items()]
                        pecas_para_geometria = [{'x': r.x, 'y': r.y, 'largura': r.width, 'altura': r.height} for r in chapa_alocada]
                        sobras_na_area_nesting = encontrar_sobras(nesting_width, nesting_height, pecas_para_geometria)
                        for s in sobras_na_area_nesting: s['x'] += margin; s['y'] += margin

                        planos_agrupados[assinatura] = {
                            "plano": plano_de_corte, "repeticoes": 1, "resumo_pecas": resumo_pecas, "sobras": sobras_na_area_nesting,
                            "chapa_largura": chapa_largura, "chapa_altura": chapa_altura
                        }
                    else:
                        planos_agrupados[assinatura]["repeticoes"] += 1
                    
                    area_total_utilizada_com_offset += sum(r.width * r.height for r in chapa_alocada)
                
                # Calcula o aproveitamento para esta solução específica (para este algoritmo).
                area_total_chapas = sum(p['chapa_largura'] * p['chapa_altura'] * p['repeticoes'] for p in planos_agrupados.values())
                
                # Para calcular a área real, precisamos iterar sobre todas as peças em todos os planos únicos
                area_real_pecas = 0
                # --- INÍCIO: LÓGICA ESPECIAL PARA CÍRCULOS ---
                area_bounding_box_pecas = 0
                is_only_circles = all(id_peca_map[r['rid']].get('forma') == 'circle' for plano in planos_agrupados.values() for r in plano['plano'])

                for plano in planos_agrupados.values():
                    for r in plano['plano']:
                        peca_info = id_peca_map[r['rid']]
                        # --- OTIMIZAÇÃO: Cálculo de área real por forma ---
                        if peca_info.get('forma') == 'circle':
                            area_real_pecas += (math.pi * (peca_info['diametro'] / 2)**2) * plano['repeticoes']
                        elif peca_info.get('forma') == 'right_triangle':
                            # A área de um triângulo é (base*altura)/2.
                            area_real_pecas += (peca_info['largura_sem_offset'] * peca_info['altura_sem_offset'] * 0.5) * plano['repeticoes']
                        elif peca_info.get('forma') in ['trapezoid', 'paired_trapezoid']:
                            # A área de um trapézio é ((B+b)*h)/2. Um par tem o dobro.
                            fator = 0.5 if peca_info.get('forma') == 'trapezoid' else 1.0
                            dims = peca_info['orig_dims']
                            area_real_pecas += ((dims['large_base'] + dims['small_base']) * dims['height'] * fator) * plano['repeticoes']
                        else:
                            # Para retângulos e DXFs (usa bounding box como aproximação)
                            area_real_pecas += (peca_info['largura_sem_offset'] * peca_info['altura_sem_offset']) * plano['repeticoes']
                        # --- FIM DA OTIMIZAÇÃO ---
                        
                        # Calcula a área do bounding box para a lógica de sucata
                        area_bounding_box_pecas += (peca_info['largura_sem_offset'] * peca_info['altura_sem_offset']) * plano['repeticoes']
                
                # Define qual área de peça usar para o cálculo da sucata
                area_pecas_para_sucata = area_bounding_box_pecas if is_only_circles else area_real_pecas
                # --- FIM: LÓGICA ESPECIAL PARA CÍRCULOS ---

                aproveitamento_geral = (area_real_pecas / area_total_chapas) * 100 if area_total_chapas > 0 else 0

                # Armazena a solução válida junto com seu aproveitamento.
                solucoes_validas_nesta_iteracao.append({
                    'aproveitamento': aproveitamento_geral,
                    'planos_agrupados': planos_agrupados,
                    'area_total_chapas': area_total_chapas,
                    'area_real_pecas': area_real_pecas,
                    'area_total_utilizada_com_offset': area_total_utilizada_com_offset
                })

        if solucoes_validas_nesta_iteracao:
            logging.info(f"Encontrado o número mínimo de chapas: {num_bins_to_try}. Selecionando a melhor solução.")
            melhor_solucao_iteracao = max(solucoes_validas_nesta_iteracao, key=lambda x: x['aproveitamento'])
            
            planos_unicos = list(melhor_solucao_iteracao['planos_agrupados'].values())
            total_chapas = sum(p['repeticoes'] for p in planos_unicos)

            # CÁLCULO DETALHADO DE SUCATA E PESOS
            total_area_sobra_aproveitavel, total_area_sobra_sucata = 0, 0
            sobras_aproveitaveis_detalhado, sucatas_dimensionadas_detalhado = [], []
            for plano in planos_unicos:
                for sobra in plano.get('sobras', []):
                    area_sobra = sobra['largura'] * sobra['altura']
                    item = {'largura': sobra['largura'], 'altura': sobra['altura'], 'peso': _calc_peso(area_sobra), 'quantidade': plano['repeticoes']}
                    if sobra['tipo_sobra'] == 'aproveitavel':
                        total_area_sobra_aproveitavel += area_sobra * plano['repeticoes']
                        sobras_aproveitaveis_detalhado.append(item)
                    else:
                        total_area_sobra_sucata += area_sobra * plano['repeticoes']
                        sucatas_dimensionadas_detalhado.append(item)

            area_offset_total = melhor_solucao_iteracao['area_total_utilizada_com_offset'] - melhor_solucao_iteracao['area_real_pecas']
            # --- CORREÇÃO: Ajusta o cálculo do offset para o caso especial de círculos e outras formas não retangulares ---
            # A área de offset é a diferença entre a área do bounding box com offset e a área do bounding box sem offset.
            area_offset_total = melhor_solucao_iteracao['area_total_utilizada_com_offset'] - area_bounding_box_pecas

            area_demais_sucatas = melhor_solucao_iteracao['area_total_chapas'] - melhor_solucao_iteracao['area_real_pecas'] - total_area_sobra_aproveitavel - total_area_sobra_sucata - area_offset_total
            # --- OTIMIZAÇÃO: Garante que a área de perda de processo nunca seja negativa ---
            area_demais_sucatas = max(0, area_demais_sucatas)
            # --- FIM DA OTIMIZAÇÃO ---

            sucata_detalhada = {
                # --- INÍCIO: LÓGICA DE PERDA PARA GUILHOTINA ---
                # Se for guilhotina, a perda por offset é zero.
                "peso_offset": 0 if is_guillotine else _calc_peso(area_offset_total),
                # --- FIM: LÓGICA DE PERDA PARA GUILHOTINA ---
                "peso_offset": _calc_peso(area_offset_total),
                "sobras_aproveitaveis": sobras_aproveitaveis_detalhado,
                "sucatas_dimensionadas": sucatas_dimensionadas_detalhado,
                "peso_demais_sucatas": _calc_peso(area_demais_sucatas)
            }

            melhor_resultado_final = {
                "planos_unicos": planos_unicos,
                "total_chapas": total_chapas,
                "aproveitamento_geral": f"{melhor_solucao_iteracao['aproveitamento']:.2f}%",
                "color_map": {},
                "area_total_chapas": melhor_solucao_iteracao['area_total_chapas'],
                "area_utilizada_real": melhor_solucao_iteracao['area_real_pecas'],
                "total_area_sobra_aproveitavel": total_area_sobra_aproveitavel,
                "total_area_sobra_sucata": total_area_sobra_sucata,
                "sucata_detalhada": sucata_detalhada
            }

            peso_total_chapas = _calc_peso(melhor_resultado_final['area_total_chapas'])
            peso_sobras_aproveitaveis = _calc_peso(melhor_resultado_final['total_area_sobra_aproveitavel'])
            peso_sobras_sucata = _calc_peso(melhor_resultado_final['total_area_sobra_sucata'])
            peso_offset = sucata_detalhada.get('peso_offset', 0)
            peso_demais_sucatas = sucata_detalhada.get('peso_demais_sucatas', 0)
            peso_perda_total_sucata = peso_sobras_sucata + peso_offset + peso_demais_sucatas
            percentual_sobras_aproveitaveis = (peso_sobras_aproveitaveis / peso_total_chapas) * 100 if peso_total_chapas > 0 else 0
            percentual_perda_total_sucata = (peso_perda_total_sucata / peso_total_chapas) * 100 if peso_total_chapas > 0 else 0

            melhor_resultado_final['peso_perda_total_sucata'] = peso_perda_total_sucata
            melhor_resultado_final['percentual_sobras_aproveitaveis'] = percentual_sobras_aproveitaveis
            melhor_resultado_final['percentual_perda_total_sucata'] = percentual_perda_total_sucata
            
            break # Encontramos a melhor solução, podemos sair do loop principal.

    if melhor_resultado_final:
        logging.info(f"Cálculo finalizado. Melhor resultado: {melhor_resultado_final['total_chapas']} chapas, {melhor_resultado_final['aproveitamento_geral']} de aproveitamento.")
    else:
        logging.warning("Nenhum resultado de cálculo foi gerado. Todas as peças podem não ter cabido nas chapas disponíveis.")
        # Retorna um dicionário vazio ou None para indicar falha
        return None 
        
    return melhor_resultado_final
