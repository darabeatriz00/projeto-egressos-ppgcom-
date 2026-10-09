# -*- coding: utf-8 -*-
"""
Automação Local Lattes -> Extração Granular, Modular e Escalável em XML.

Principais Funcionalidades:
    1. Nomeação Padronizada dos Arquivos XML:
       - Formato: `<NOME_DO_EGRESSO>_<ID_LATTES>.xml` (ex: `Janaina_Vieira_de_Paula_Jordao_9936127174016398.xml`).
    2. Atuação Profissional Robusta:
       - Processamento pareado de rótulos (layout-cell-3) e conteúdos (layout-cell-9).
       - Captura e anexa corretamente "Outras informações" ao respectivo <VINCULO>.
       - Identifica Vínculo Atual (VINCULO-ATUAL="SIM"/"NAO") com datas de início e fim.
       - Estrutura cargos de gestão, comissões, ensino e serviços técnicos.
    3. Linhas de Pesquisa e Projetos de Pesquisa:
       - Módulos próprios com status (EM_ANDAMENTO / CONCLUIDO), descrição e financiadores.
    4. Simulação Humana Pré-Captcha:
       - Execução única no desafio para evitar loops.
    5. Controle em Lote via Planilha:
       - Processamento por faixas (EGRESSO_INICIAL e EGRESSO_FINAL).

Pré-requisitos:
    pip install pandas openpyxl selenium webdriver-manager beautifulsoup4 lxml
"""

import os
import re
import time
import random
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

# ==============================================================================
# CONFIGURAÇÕES E PARÂMETROS DE CONTROLE
# ==============================================================================
DIRETORIO_BASE = os.path.dirname(os.path.abspath(__file__))
DIRETORIO_SAIDA = os.path.join(DIRETORIO_BASE, "xml_lattes_estruturados")
os.makedirs(DIRETORIO_SAIDA, exist_ok=True)

NOME_ARQUIVO_EXCEL = "egressos_ppgcom_consolidado.xlsx"
CAMINHO_EXCEL = os.path.join(DIRETORIO_BASE, NOME_ARQUIVO_EXCEL)
ABA_PLANILHA = "Todos Consolidado"

# Intervalo de Registros da Planilha (Base 1)
EGRESSO_INICIAL = 3
EGRESSO_FINAL = 3

PULAR_SE_EXISTIR = True


# ==============================================================================
# ENGINE DE COMPORTAMENTO HUMANO
# ==============================================================================
def pausa_aleatoria(min_s: float = 0.4, max_s: float = 1.2):
    time.sleep(random.uniform(min_s, max_s))


def simular_movimentos_exploratorios(driver, passos: int = 3):
    try:
        actions = ActionChains(driver)
        for _ in range(passos):
            actions.move_by_offset(random.randint(-80, 80), random.randint(-50, 50))
            actions.pause(random.uniform(0.1, 0.25))
        actions.perform()
    except Exception:
        pass


def rolar_pagina_natural(driver, vezes: int = 2):
    for _ in range(vezes):
        distancia = random.randint(150, 300)
        direcao = 1 if random.random() > 0.2 else -1
        driver.execute_script(f"window.scrollBy({{top: {distancia * direcao}, behavior: 'smooth'}});")
        pausa_aleatoria(0.3, 0.6)


def simular_comportamento_pre_captcha(driver):
    print("[*] Simulando comportamento prévio (leitura, rolagem e movimento de cursor)...")
    pausa_aleatoria(0.5, 1.0)
    rolar_pagina_natural(driver, vezes=2)
    simular_movimentos_exploratorios(driver, passos=2)
    pausa_aleatoria(0.4, 0.8)


def mover_e_clicar_humanizado(driver, elemento):
    try:
        actions = ActionChains(driver)
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", elemento)
        pausa_aleatoria(0.3, 0.6)
        actions.move_to_element_with_offset(elemento, random.randint(-4, 4), random.randint(-3, 3))
        actions.pause(random.uniform(0.15, 0.3))
        actions.click()
        actions.perform()
        pausa_aleatoria(0.4, 0.8)
        return True
    except Exception:
        try:
            elemento.click()
            return True
        except Exception:
            return False


def criar_driver_stealth():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-infobars")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"}
    )
    return driver


def lidar_com_captcha_e_carregamento(driver, tempo_limite=180) -> bool:
    inicio = time.time()
    simulacao_executada = False

    while time.time() - inicio <= tempo_limite:
        if driver.find_elements(By.CSS_SELECTOR, ".infpessoa, .nome, .resumo"):
            return True

        checkbox_captcha = driver.find_elements(By.CSS_SELECTOR, ".recaptcha-checkbox-border, #recaptcha-anchor")
        botao_submit = driver.find_elements(By.CSS_SELECTOR, "#submitBtn, input[type='submit']")

        if checkbox_captcha or botao_submit:
            if not simulacao_executada:
                simular_comportamento_pre_captcha(driver)
                simulacao_executada = True

                if checkbox_captcha and checkbox_captcha[0].is_displayed():
                    print("[+] Clicando no checkbox de validação...")
                    mover_e_clicar_humanizado(driver, checkbox_captcha[0])
                    pausa_aleatoria(1.0, 2.0)

                print("\n" + "=" * 70)
                print("[!] ATENÇÃO: Conclua a validação do Captcha no navegador se solicitado.")
                print("[!] O script continuará automaticamente assim que a página for liberada.")
                print("=" * 70 + "\n")

            if botao_submit and botao_submit[0].is_enabled():
                print("[+] Botão habilitado. Submetendo formulário...")
                mover_e_clicar_humanizado(driver, botao_submit[0])
                pausa_aleatoria(2.0, 3.0)

        time.sleep(2)
    return False


# ==============================================================================
# FUNÇÕES DE EXTRAÇÃO E PARSER GRANULAR
# ==============================================================================
def sanitizar_nome_arquivo(nome: str) -> str:
    nome = re.sub(r'[\\/*?:"<>|]', "", str(nome))
    return re.sub(r"\s+", "_", nome.strip())


def extrair_id_lattes_de_url(url_ou_texto: str) -> str:
    match = re.search(r"(\d{15,20})", str(url_ou_texto))
    return match.group(1) if match else "ID_DESCONHECIDO"


def parse_periodo(texto: str):
    texto = texto.strip()
    match_intervalo = re.search(r"(\d{2}/\d{4}|\d{4})\s*-\s*(\d{2}/\d{4}|\d{4}|Atual)", texto, re.IGNORECASE)
    if match_intervalo:
        inicio = match_intervalo.group(1).strip()
        fim = match_intervalo.group(2).strip()
        atual = "SIM" if fim.lower() == "atual" else "NAO"
        return inicio, fim, atual
    match_unico = re.search(r"(\d{2}/\d{4}|\d{4})", texto)
    if match_unico:
        val = match_unico.group(1).strip()
        return val, val, "NAO"
    return "", "", "NAO"


def parse_formacao_detalhe(detalhe_txt: str):
    titulo = ""
    match_tit = re.search(r"Título:\s*([^,\n\.]+)", detalhe_txt, re.IGNORECASE)
    if match_tit:
        titulo = match_tit.group(1).strip()

    orientador = ""
    match_orient = re.search(r"Orientador:\s*([^,\n\.]+)", detalhe_txt, re.IGNORECASE)
    if match_orient:
        orientador = match_orient.group(1).strip()

    bolsa = ""
    match_bolsa = re.search(r"Bolsista do\(a\):\s*([^,\n\.]+)", detalhe_txt, re.IGNORECASE)
    if match_bolsa:
        bolsa = match_bolsa.group(1).strip()

    ch = ""
    match_ch = re.search(r"Carga Horária:\s*(\d+h?)", detalhe_txt, re.IGNORECASE)
    if match_ch:
        ch = match_ch.group(1).strip()

    palavras_chave = ""
    match_pc = re.search(r"Palavras-chave:\s*([^,\n\.]+)", detalhe_txt, re.IGNORECASE)
    if match_pc:
        palavras_chave = match_pc.group(1).strip()

    partes = [p.strip() for p in detalhe_txt.split(". ") if p.strip()]
    curso = partes[0] if len(partes) > 0 else ""
    instituicao = partes[1] if len(partes) > 1 else ""

    return {
        "CURSO": curso,
        "INSTITUICAO": instituicao,
        "TITULO": titulo,
        "ORIENTADOR": orientador,
        "BOLSA": bolsa,
        "CARGA-HORARIA": ch,
        "PALAVRAS-CHAVE": palavras_chave
    }


def extrair_e_gerar_xml_estruturado(html_content: str, url_origem: str, nome_padrao: str = "") -> tuple[str, str, str]:
    soup = BeautifulSoup(html_content, "html.parser")

    # Identificação Básica
    nome_tag = soup.find("h2", class_="nome")
    nome = nome_tag.get_text(strip=True) if nome_tag else (nome_padrao or "Não informado")

    match_id = re.search(r"http://lattes\.cnpq\.br/(\d+)", html_content)
    if match_id:
        id_lattes = match_id.group(1)
    else:
        id_lattes = extrair_id_lattes_de_url(url_origem)

    data_atualizacao = "Não informada"
    for li in soup.find_all("li"):
        txt = li.get_text(strip=True)
        if "atualização do currículo" in txt.lower():
            match_d = re.search(r"\d{2}/\d{2}/\d{4}", txt)
            if match_d:
                data_atualizacao = match_d.group(0)
            break

    distincao = ""
    for h in soup.find_all("h2"):
        if "bolsista" in h.get_text(strip=True).lower():
            distincao = h.get_text(strip=True)
            break

    citacoes = ""
    div_cit = soup.find(string=re.compile(r"Nome em citações bibliográficas"))
    if div_cit and div_cit.find_parent("div"):
        val_cit = div_cit.find_parent("div").find_next_sibling("div")
        if val_cit:
            citacoes = val_cit.get_text(strip=True)

    link_orcid = "Não encontrado"
    for a in soup.find_all("a", href=True):
        if "orcid.org" in a["href"]:
            link_orcid = a["href"].strip()
            break

    link_linkedin = "Não encontrado"
    for a in soup.find_all("a", href=True):
        if "linkedin.com" in a["href"]:
            link_linkedin = a["href"].strip()
            break
    if link_linkedin == "Não encontrado":
        match_lk = re.search(r"https?://(www\.)?linkedin\.com/[a-zA-Z0-9_\-\./]+", html_content, re.IGNORECASE)
        if match_lk:
            link_linkedin = match_lk.group(0).strip()

    resumo_texto = "Não informado"
    resumo_p = soup.find("p", class_="resumo")
    if resumo_p:
        resumo_texto = resumo_p.get_text(" ", strip=True).replace("(Texto informado pelo autor)", "").strip()

    # =========================================================================
    # ÁRVORE XML PADRONIZADA
    # =========================================================================
    root = ET.Element("CURRICULO-LATTES", {
        "ID": id_lattes,
        "DATA-ATUALIZACAO": data_atualizacao,
        "DATA-EXTRACAO": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    })

    # 1. IDENTIFICAÇÃO
    sec_ident = ET.SubElement(root, "IDENTIFICACAO")
    ET.SubElement(sec_ident, "NOME").text = nome
    ET.SubElement(sec_ident, "ID-LATTES").text = id_lattes
    if citacoes:
        ET.SubElement(sec_ident, "CITACAO-BIBLIOGRAFICA").text = citacoes
    ET.SubElement(sec_ident, "NACIONALIDADE").text = "Brasil"
    if distincao:
        ET.SubElement(sec_ident, "DISTINCAO").text = distincao

    # 2. LINKS E REDES
    sec_links = ET.SubElement(root, "LINKS-E-REDES")
    ET.SubElement(sec_links, "LINK", {"TIPO": "LATTES"}).text = f"http://lattes.cnpq.br/{id_lattes}" if id_lattes != "ID_DESCONHECIDO" else "Não encontrado"
    ET.SubElement(sec_links, "LINK", {"TIPO": "ORCID"}).text = link_orcid
    ET.SubElement(sec_links, "LINK", {"TIPO": "LINKEDIN"}).text = link_linkedin

    # 3. RESUMO PROFISSIONAL
    sec_resumo = ET.SubElement(root, "RESUMO-PROFISSIONAL")
    ET.SubElement(sec_resumo, "TEXTO").text = resumo_texto

    # 4. ENDEREÇO PROFISSIONAL
    sec_end = ET.SubElement(root, "ENDERECO-PROFISSIONAL")
    end_tag = soup.find("a", attrs={"name": "Endereco"})
    if end_tag and end_tag.parent:
        end_bloco = end_tag.parent.find_next("div", class_="data-cell")
        if end_bloco:
            end_texto = end_bloco.get_text("\n", strip=True)
            linhas_end = [l.strip() for l in end_texto.split("\n") if l.strip() and "endereço" not in l.lower()]

            inst_end = linhas_end[0] if len(linhas_end) > 0 else ""
            campus_end = linhas_end[1] if len(linhas_end) > 1 else ""
            bairro_end = linhas_end[2] if len(linhas_end) > 2 else ""

            cep_m = re.search(r"(\d{8}|\d{5}-\d{3})", end_texto)
            tel_m = re.search(r"Telefone:\s*([^,\n]+)", end_texto)
            home_m = re.search(r"(http[s]?://\S+)", end_texto)

            ET.SubElement(sec_end, "INSTITUICAO").text = inst_end
            ET.SubElement(sec_end, "CAMPUS").text = campus_end
            ET.SubElement(sec_end, "BAIRRO").text = bairro_end
            ET.SubElement(sec_end, "CIDADE").text = "Goiânia" if "Goiânia" in end_texto else ""
            ET.SubElement(sec_end, "UF").text = "GO" if "GO" in end_texto else ""
            ET.SubElement(sec_end, "PAIS").text = "Brasil"
            ET.SubElement(sec_end, "CEP").text = cep_m.group(1) if cep_m else ""
            ET.SubElement(sec_end, "TELEFONE").text = tel_m.group(1).strip() if tel_m else ""
            ET.SubElement(sec_end, "HOMEPAGE").text = home_m.group(1).strip() if home_m else ""

    # 5. FORMAÇÃO ACADÊMICA
    sec_form = ET.SubElement(root, "FORMACAO-ACADEMICA")
    for a_tag in soup.find_all("a", attrs={"name": re.compile(r"FormacaoAcademica.*")}):
        secao = a_tag.find_next("div", class_="data-cell")
        if not secao:
            continue
        p_divs = secao.find_all("div", class_=re.compile(r"layout-cell-3"))
        d_divs = secao.find_all("div", class_=re.compile(r"layout-cell-9"))

        for p_div, d_div in zip(p_divs, d_divs):
            p_txt = p_div.get_text(strip=True)
            d_txt = d_div.get_text(" ", strip=True)
            if not p_txt or not d_txt:
                continue

            ano_ini, ano_fim, _ = parse_periodo(p_txt)
            parsed = parse_formacao_detalhe(d_txt)

            tag_nome = "TITULACAO"
            low = d_txt.lower()
            if "doutorado" in low:
                tag_nome = "DOUTORADO"
            elif "mestrado" in low:
                tag_nome = "MESTRADO"
            elif "especialização" in low:
                tag_nome = "ESPECIALIZACAO"
            elif "graduação" in low:
                tag_nome = "GRADUACAO"
            elif "pós-doutorado" in low:
                tag_nome = "POS-DOUTORADO"

            node_item = ET.SubElement(sec_form, tag_nome, {
                "ANO-INICIO": ano_ini,
                "ANO-CONCLUSAO": ano_fim
            })

            if parsed["CURSO"]:
                ET.SubElement(node_item, "CURSO").text = parsed["CURSO"]
            if parsed["INSTITUICAO"]:
                ET.SubElement(node_item, "INSTITUICAO").text = parsed["INSTITUICAO"]
            if parsed["CARGA-HORARIA"]:
                ET.SubElement(node_item, "CARGA-HORARIA").text = parsed["CARGA-HORARIA"]
            if parsed["TITULO"]:
                ET.SubElement(node_item, "TITULO-TRABALHO").text = parsed["TITULO"]
            if parsed["ORIENTADOR"]:
                ET.SubElement(node_item, "ORIENTADOR").text = parsed["ORIENTADOR"]
            if parsed["BOLSA"]:
                ET.SubElement(node_item, "AGENCIA-FOMENTO").text = parsed["BOLSA"]
            if parsed["PALAVRAS-CHAVE"]:
                ET.SubElement(node_item, "PALAVRAS-CHAVE").text = parsed["PALAVRAS-CHAVE"]

    # Cursos Complementares
    form_comp_tag = soup.find("a", attrs={"name": "FormacaoComplementar"})
    if form_comp_tag and form_comp_tag.parent:
        cell_comp = form_comp_tag.parent.find_next("div", class_="data-cell")
        if cell_comp:
            node_comp_parent = ET.SubElement(sec_form, "CURSOS-COMPLEMENTARES")
            p_comp = cell_comp.find_all("div", class_=re.compile(r"layout-cell-3"))
            d_comp = cell_comp.find_all("div", class_=re.compile(r"layout-cell-9"))
            for p, d in zip(p_comp, d_comp):
                ano_c = p.get_text(strip=True)
                det_c = d.get_text(" ", strip=True)
                ch_m = re.search(r"Carga horária:\s*(\d+h?)", det_c, re.IGNORECASE)
                partes_c = [x.strip() for x in det_c.split(".") if x.strip()]
                nome_c = partes_c[0] if len(partes_c) > 0 else ""
                inst_c = partes_c[1] if len(partes_c) > 1 else ""

                node_c = ET.SubElement(node_comp_parent, "CURSO", {
                    "ANO": ano_c,
                    "CARGA-HORARIA": ch_m.group(1) if ch_m else ""
                })
                ET.SubElement(node_c, "NOME").text = nome_c
                if inst_c:
                    ET.SubElement(node_c, "INSTITUICAO").text = inst_c

    # 6. ATUAÇÃO PROFISSIONAL DETALHADA
    sec_atuacao = ET.SubElement(root, "ATUACAO-PROFISSIONAL")
    div_atuacao = soup.find("a", attrs={"name": "AtuacaoProfissional"})
    if div_atuacao and div_atuacao.parent:
        cell_atuacao = div_atuacao.parent.find_next("div", class_="data-cell")
        if cell_atuacao:
            instituicao_atual = None
            bloco_vinculo_atual = None

            divs_filhos = [c for c in cell_atuacao.children if getattr(c, "name", None) == "div"]
            idx_d = 0

            while idx_d < len(divs_filhos):
                d_item = divs_filhos[idx_d]
                classes_item = d_item.get("class", [])

                if "inst_back" in classes_item:
                    nome_inst = d_item.get_text(strip=True)
                    instituicao_atual = ET.SubElement(sec_atuacao, "INSTITUICAO", {"NOME": nome_inst})
                    bloco_vinculo_atual = None
                    idx_d += 1
                    continue

                if "layout-cell-3" in classes_item:
                    rotulo_txt = d_item.get_text(" ", strip=True)

                    if idx_d + 1 < len(divs_filhos) and "layout-cell-9" in divs_filhos[idx_d + 1].get("class", []):
                        valor_div = divs_filhos[idx_d + 1]
                        valor_txt = valor_div.get_text(" ", strip=True)
                        idx_d += 2

                        if not valor_txt or instituicao_atual is None:
                            continue

                        rotulo_low = rotulo_txt.lower()

                        if "outras informações" in rotulo_low or "outras informacoes" in rotulo_low:
                            if bloco_vinculo_atual is not None:
                                ET.SubElement(bloco_vinculo_atual, "OUTRAS-INFORMACOES").text = valor_txt
                            else:
                                ET.SubElement(instituicao_atual, "OUTRAS-INFORMACOES").text = valor_txt
                            continue

                        ini, fim, atual = parse_periodo(rotulo_txt)

                        if "vínculo:" in valor_txt.lower() or "enquadramento funcional:" in valor_txt.lower():
                            v_m = re.search(r"Vínculo:\s*([^,]+)", valor_txt)
                            eq_m = re.search(r"Enquadramento Funcional:\s*([^,]+)", valor_txt)
                            ch_m = re.search(r"Carga horária:\s*(\d+)", valor_txt)
                            reg_m = re.search(r"Regime:\s*([^,\.]+)", valor_txt)

                            bloco_vinculo_atual = ET.SubElement(instituicao_atual, "VINCULO", {
                                "PERIODO-INICIO": ini,
                                "PERIODO-FIM": fim,
                                "VINCULO-ATUAL": atual
                            })
                            if v_m:
                                ET.SubElement(bloco_vinculo_atual, "TIPO-VINCULO").text = v_m.group(1).strip()
                            if eq_m:
                                ET.SubElement(bloco_vinculo_atual, "ENQUADRAMENTO-FUNCIONAL").text = eq_m.group(1).strip()
                            if ch_m:
                                ET.SubElement(bloco_vinculo_atual, "CARGA-HORARIA-SEMANAL").text = ch_m.group(1).strip()
                            if reg_m:
                                ET.SubElement(bloco_vinculo_atual, "REGIME-DE-TRABALHO").text = reg_m.group(1).strip()

                        else:
                            bloco_atv = ET.SubElement(instituicao_atual, "ATIVIDADE", {
                                "TIPO": rotulo_txt if not ini else "ATIVIDADE_ESPECIFICA",
                                "PERIODO-INICIO": ini,
                                "PERIODO-FIM": fim,
                                "EM-ANDAMENTO": atual
                            })

                            if "cargo ou função" in valor_txt.lower():
                                funcao_txt = valor_txt.replace("Cargo ou função", "").strip()
                                ET.SubElement(bloco_atv, "CARGO-FUNCAO").text = funcao_txt
                            elif "disciplinas ministradas" in valor_txt.lower():
                                discs = [d.strip() for d in valor_txt.replace("Disciplinas ministradas", "").split("<br>") if d.strip()]
                                if len(discs) == 1 and ";" in discs[0]:
                                    discs = [d.strip() for d in discs[0].split(";") if d.strip()]
                                node_d = ET.SubElement(bloco_atv, "DISCIPLINAS-MINISTRADAS")
                                for d in discs:
                                    ET.SubElement(node_d, "DISCIPLINA").text = d
                            elif "atividade realizada" in valor_txt.lower():
                                ET.SubElement(bloco_atv, "DESCRICAO-ATIVIDADE").text = valor_txt.replace("Atividade realizada", "").strip()
                            elif "serviço realizado" in valor_txt.lower():
                                ET.SubElement(bloco_atv, "SERVICO-TECNICO").text = valor_txt.replace("Serviço realizado", "").strip()
                            elif "estágio realizado" in valor_txt.lower():
                                ET.SubElement(bloco_atv, "ESTAGIO").text = valor_txt.replace("Estágio realizado", "").strip()
                            else:
                                ET.SubElement(bloco_atv, "DESCRICAO").text = valor_txt
                    else:
                        idx_d += 1
                else:
                    idx_d += 1

    # 7. LINHAS DE PESQUISA
    sec_lp = ET.SubElement(root, "LINHAS-DE-PESQUISA")
    div_lp = soup.find("a", attrs={"name": "LinhaPesquisa"})
    if div_lp and div_lp.parent:
        cell_lp = div_lp.parent.find_next("div", class_="data-cell")
        if cell_lp:
            for item_lp in cell_lp.find_all("div", class_=re.compile(r"layout-cell-9")):
                txt_lp = item_lp.get_text(" ", strip=True)
                if txt_lp:
                    ET.SubElement(sec_lp, "LINHA-DE-PESQUISA").text = txt_lp

    # 8. PROJETOS DE PESQUISA
    sec_proj = ET.SubElement(root, "PROJETOS-DE-PESQUISA")
    div_proj = soup.find("a", attrs={"name": "ProjetosPesquisa"})
    if div_proj and div_proj.parent:
        cell_proj = div_proj.parent.find_next("div", class_="data-cell")
        if cell_proj:
            p_divs = cell_proj.find_all("div", class_=re.compile(r"layout-cell-3"))
            d_divs = cell_proj.find_all("div", class_=re.compile(r"layout-cell-9"))

            for p_tag, d_tag in zip(p_divs, d_divs):
                p_txt = p_tag.get_text(strip=True)
                d_txt = d_tag.get_text(" ", strip=True)
                if not d_txt:
                    continue

                ini, fim, atual = parse_periodo(p_txt)

                nome_proj = d_txt.split("Descrição:")[0].split("Situação:")[0].strip()
                desc_m = re.search(r"Descrição:\s*(.*?)(?=Situação:|$)", d_txt)
                desc_proj = desc_m.group(1).strip() if desc_m else ""

                sit_m = re.search(r"Situação:\s*([^;]+)", d_txt)
                status_proj = sit_m.group(1).strip().upper().replace(" ", "_") if sit_m else ("EM_ANDAMENTO" if atual == "SIM" else "CONCLUIDO")

                integ_m = re.search(r"Integrantes:\s*(.*?)(?=Financiador|$)", d_txt)
                integrantes = integ_m.group(1).strip() if integ_m else ""

                fianc_m = re.search(r"Financiador\(es\):\s*(.*?)(?=\.|$)", d_txt)
                fomento = fianc_m.group(1).strip() if fianc_m else ""

                node_p = ET.SubElement(sec_proj, "PROJETO", {
                    "ANO-INICIO": ini,
                    "ANO-FIM": fim,
                    "SITUACAO": status_proj
                })
                ET.SubElement(node_p, "NOME-PROJETO").text = nome_proj
                if desc_proj:
                    ET.SubElement(node_p, "DESCRICAO").text = desc_proj
                if integrantes:
                    ET.SubElement(node_p, "INTEGRANTES").text = integrantes
                if fomento:
                    ET.SubElement(node_p, "AGENCIA-FOMENTO").text = fomento

    xml_final = minidom.parseString(ET.tostring(root, encoding="utf-8")).toprettyxml(indent="  ")
    return xml_final, nome, id_lattes


# ==============================================================================
# FLUXO DE EXECUÇÃO EM LOTE
# ==============================================================================
def executar_lote_lattes():
    print("=" * 85)
    print("AUTOMAÇÃO LATTES -> PROCESSAMENTO MODULAR E ESCALÁVEL")
    print(f"Arquivo Base: {CAMINHO_EXCEL}")
    print(f"Intervalo: Egressos {EGRESSO_INICIAL} a {EGRESSO_FINAL}")
    print("=" * 85)

    if not os.path.exists(CAMINHO_EXCEL):
        print(f"[X] Arquivo não encontrado: {CAMINHO_EXCEL}")
        return

    df = pd.read_excel(CAMINHO_EXCEL, sheet_name=ABA_PLANILHA)
    df.columns = [str(c).strip() for c in df.columns]
    total_registros = len(df)

    idx_inicio = max(0, EGRESSO_INICIAL - 1)
    idx_fim = min(total_registros, EGRESSO_FINAL)

    df_lote = df.iloc[idx_inicio:idx_fim].copy()
    print(f"[+] Processando {len(df_lote)} registros (do #{EGRESSO_INICIAL} ao #{idx_fim})...\n")

    driver = criar_driver_stealth()

    sucessos = 0
    ignorados = 0
    erros = 0

    try:
        for idx, (_, linha) in enumerate(df_lote.iterrows(), start=idx_inicio + 1):
            nome_planilha = str(linha.get("Discente", "")).strip()
            url_lattes = str(linha.get("Currículo Lattes", "")).strip()
            id_lattes_planilha = extrair_id_lattes_de_url(url_lattes)

            print("-" * 85)
            print(f"[{idx}/{idx_fim}] Egresso: {nome_planilha}")
            print(f"URL: {url_lattes}")

            if not url_lattes or "lattes.cnpq.br" not in url_lattes:
                print("[-] Sem link Lattes válido. Pulando...")
                ignorados += 1
                continue

            # Nome do arquivo no padrão: <NOME_EGRESSO>_<ID_LATTES>.xml
            nome_arquivo = f"{sanitizar_nome_arquivo(nome_planilha)}_{id_lattes_planilha}.xml"
            caminho_xml = os.path.join(DIRETORIO_SAIDA, nome_arquivo)

            if PULAR_SE_EXISTIR and os.path.exists(caminho_xml):
                print(f"[i] XML já existe ({nome_arquivo}). Pulando...")
                ignorados += 1
                continue

            try:
                driver.get(url_lattes)
                pausa_aleatoria(1.0, 2.0)

                if not lidar_com_captcha_e_carregamento(driver):
                    print(f"[X] Tempo limite esgotado para '{nome_planilha}'")
                    erros += 1
                    continue

                html_pagina = driver.page_source
                xml_conteudo, nome_lattes, id_extraido = extrair_e_gerar_xml_estruturado(
                    html_pagina, url_lattes, nome_padrao=nome_planilha
                )

                # Atualiza o nome do arquivo se o ID ou Nome tiver sido refinado do HTML
                id_final = id_extraido if id_extraido != "ID_DESCONHECIDO" else id_lattes_planilha
                nome_final = nome_lattes if nome_lattes != "Não informado" else nome_planilha
                nome_arquivo_final = f"{sanitizar_nome_arquivo(nome_final)}_{id_final}.xml"
                caminho_xml_final = os.path.join(DIRETORIO_SAIDA, nome_arquivo_final)

                with open(caminho_xml_final, "w", encoding="utf-8") as f:
                    f.write(xml_conteudo)

                print(f"[✓] XML salvo com sucesso: {nome_arquivo_final}")
                sucessos += 1
                pausa_aleatoria(1.5, 3.0)

            except Exception as erro_item:
                print(f"[X] Erro ao processar '{nome_planilha}': {erro_item}")
                erros += 1

    finally:
        pausa_aleatoria(0.8, 1.5)
        driver.quit()
        print("\n" + "=" * 85)
        print("RESUMO DO PROCESSAMENTO:")
        print(f"  • Intervalo: {EGRESSO_INICIAL} a {idx_fim}")
        print(f"  • Sucessos: {sucessos}")
        print(f"  • Ignorados/Existentes: {ignorados}")
        print(f"  • Erros: {erros}")
        print(f"  • Pasta de Saída: {DIRETORIO_SAIDA}")
        print("=" * 85)


if __name__ == "__main__":
    executar_lote_lattes()
