import time
import requests
import pandas as pd

# ==============================================================================
# CONFIGURAÇÕES
# ==============================================================================
# Planilha gerada a partir da extração dos arquivos XML do Lattes
ARQUIVO_ENTRADA = 'egressos_orcid_extraido_lattes.xlsx'
ABA_ENTRADA = 0 

ARQUIVO_SAIDA = 'publicacoes_egressos_openalex2.xlsx'

# ==============================================================================
# FUNÇÃO DE EXTRAÇÃO COMPLETA DA API OPENALEX
# ==============================================================================
def extrair_dados_completos_openalex(orcid):
    """Extrai todos os campos e metadados estratégicos da API de Works do OpenAlex."""
    orcid_clean = str(orcid).strip()
    
    # Tratamento de valores nulos ou vazios
    if not orcid_clean or orcid_clean.lower() in ['nan', 'não encontrado', 'none', '']:
        return []

    if not orcid_clean.startswith('https://orcid.org/'):
        orcid_clean = f"https://orcid.org/{orcid_clean}"

    trabalhos = []
    page = 1

    while True:
        url = f"https://api.openalex.org/works?filter=author.orcid:{orcid_clean}&per_page=200&page={page}"

        try:
            response = requests.get(url, timeout=15)
            if response.status_code != 200:
                break

            data = response.json()
            results = data.get('results', [])
            if not results:
                break

            for item in results:
                # 1. Local de Publicação e Fonte
                primary_loc = item.get('primary_location') or {}
                source = primary_loc.get('source') or {}

                # 2. Dados de Acesso Aberto (OA)
                oa_info = item.get('open_access') or {}

                # 3. Autoria, Coautores e Instituições Filiadas
                authorships = item.get('authorships') or []
                coautores = []
                instituicoes = set()
                paises = set()

                for auth in authorships:
                    author_obj = auth.get('author') or {}
                    if author_obj.get('display_name'):
                        coautores.append(author_obj.get('display_name'))

                    for inst in auth.get('institutions') or []:
                        if inst.get('display_name'):
                            instituicoes.add(inst.get('display_name'))
                        if inst.get('country_code'):
                            paises.add(inst.get('country_code'))

                # 4. Taxonomia e Classificação Acadêmica
                primary_topic = item.get('primary_topic') or {}
                subfield = primary_topic.get('subfield') or {}
                field = primary_topic.get('field') or {}
                domain = primary_topic.get('domain') or {}

                # 5. Conceitos / Palavras-chave (Top 5 por Relevância)
                concepts = item.get('concepts') or []
                concepts_sorted = sorted(concepts, key=lambda x: x.get('score', 0), reverse=True)
                top_concepts = [c.get('display_name') for c in concepts_sorted[:5] if c.get('display_name')]

                # 6. Histórico de Citações por Ano
                counts_by_year = item.get('counts_by_year') or []
                hist_citacoes = [f"{c.get('year')}:{c.get('cited_by_count')}" for c in counts_by_year]

                # 7. Financiadores e Agências de Fomento
                grants = item.get('grants') or []
                financiadores = [g.get('funder_display_name') for g in grants if g.get('funder_display_name')]

                # 8. Objetivos de Desenvolvimento Sustentável (ODS/SDG da ONU)
                sdgs = item.get('sustainable_development_goals') or []
                ods_lista = [s.get('display_name') for s in sdgs if s.get('display_name')]

                # Consolidação dos Dados da Publicação
                trabalhos.append({
                    # Identificadores Básicos
                    'ID_OpenAlex': item.get('id'),
                    'DOI': item.get('doi'),
                    'Titulo_Trabalho': item.get('title'),
                    'Ano_Publicacao': item.get('publication_year'),
                    'Data_Exata_Publicacao': item.get('publication_date'),
                    'Tipo_Documento': item.get('type'),
                    'Idioma': item.get('language'),

                    # Métricas de Citação
                    'Citacoes_Totais': item.get('cited_by_count', 0),
                    'Historico_Citacoes_Ano': " | ".join(hist_citacoes),
                    'Qtd_Referencias_Citadas': len(item.get('referenced_works') or []),
                    'Qtd_Obras_Relacionadas': len(item.get('related_works') or []),

                    # Acesso Aberto (Open Access)
                    'Acesso_Aberto': 'Sim' if oa_info.get('is_oa') else 'Não',
                    'Status_Acesso_Aberto': oa_info.get('oa_status'),
                    'URL_PDF_Aberto': primary_loc.get('pdf_url'),
                    'URL_Pagina_Publicacao': primary_loc.get('landing_page_url'),

                    # Veículo / Periódico / Evento
                    'Nome_Veiculo': source.get('display_name'),
                    'Tipo_Veiculo': source.get('type'),
                    'Editora_Publisher': source.get('host_organization_name'),
                    'ISSN': source.get('issn_l'),

                    # Rede de Colaboração e Instituições
                    'Total_Autores': len(authorships),
                    'Lista_Coautores': " ; ".join(coautores),
                    'Instituicoes_Filiadas': " ; ".join(instituicoes),
                    'Paises_Instituicoes': " ; ".join(paises),

                    # Áreas do Conhecimento (Classificação)
                    'Topico_Principal': primary_topic.get('display_name'),
                    'Subcampo_Conhecimento': subfield.get('display_name'),
                    'Campo_Conhecimento': field.get('display_name'),
                    'Grande_Area_Conhecimento': domain.get('display_name'),
                    'Palavras_Chave_Conceitos': " | ".join(top_concepts),

                    # Fomento e Impacto Social
                    'Agencias_Financiadoras': " ; ".join(set(financiadores)),
                    'ODS_ONU_Associados': " ; ".join(ods_lista)
                })

            if len(results) < 200:
                break
            page += 1

        except Exception as e:
            print(f"   [Erro] Falha ao processar ORCID {orcid}: {e}")
            break

    return trabalhos

# ==============================================================================
# EXECUÇÃO DO PIPELINE
# ==============================================================================
print(f"1. Carregando planilha '{ARQUIVO_ENTRADA}'...")
df_egressos = pd.read_excel(ARQUIVO_ENTRADA, sheet_name=ABA_ENTRADA)

# Filtra apenas egressos que possuem ORCID extraído dos XMLs do Lattes
df_validos = df_egressos[
    df_egressos['Tem_ORCID'].astype(str).str.lower() == 'sim'
].copy()

todas_publicacoes = []
resumo_egressos = []

print(f"2. Coletando dados completos para {len(df_validos)} egressos com ORCID...\n")

for idx, (_, row) in enumerate(df_validos.iterrows(), 1):
    discente = row['Nome']
    orcid = row['ORCID']
    id_lattes = row['ID_Lattes']
    xml_origem = row.get('Nome_Arquivo_XML', '')

    print(f"[{idx}/{len(df_validos)}] Coletando: {discente} (ORCID: {orcid})...")

    pubs = extrair_dados_completos_openalex(orcid)

    # Adiciona os dados do egresso a cada publicação obtida
    for p in pubs:
        registro = {
            'Discente_UFG': discente,
            'ID_Lattes': id_lattes,
            'ORCID_Egresso': orcid,
            'Arquivo_XML_Origem': xml_origem,
            **p
        }
        todas_publicacoes.append(registro)

    # Mapeamento do perfil consolidado do egresso
    resumo_egressos.append({
        'Discente': discente,
        'ID_Lattes': id_lattes,
        'ORCID': orcid,
        'Total_Publicacoes_OpenAlex': len(pubs),
        'Total_Citacoes_Acumuladas': sum(p['Citacoes_Totais'] for p in pubs),
        'Qtd_Publicacoes_Acesso_Aberto': sum(1 for p in pubs if p['Acesso_Aberto'] == 'Sim'),
        'Paises_Colaboracao': " ; ".join(set(p['Paises_Instituicoes'] for p in pubs if p['Paises_Instituicoes']))
    })

    time.sleep(0.2)

# ==============================================================================
# GRAVAÇÃO DOS DADOS
# ==============================================================================
print("\n3. Salvando dados estruturados no Excel...")
df_pubs = pd.DataFrame(todas_publicacoes)
df_resumo = pd.DataFrame(resumo_egressos)

with pd.ExcelWriter(ARQUIVO_SAIDA, engine='openpyxl') as writer:
    df_resumo.to_excel(writer, sheet_name='Resumo de Indicadores', index=False)
    df_pubs.to_excel(writer, sheet_name='Publicações Detalhadas', index=False)

print(f"✅ Processo concluído! Arquivo salvo em: '{ARQUIVO_SAIDA}'")
