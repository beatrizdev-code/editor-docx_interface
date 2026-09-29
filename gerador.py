from docx import Document

def montar_aspectos_gerais (dados:dict) -> str:
    nome = dados ['nome_primeiro']
    idade = dados ['idade']

    if dados['tem_filhos']:
        qtd = dados.get('qtd_filhos','').strip()
        if qtd:
            parte_filhos= f'{qtd} filho(s)'
        else:
            parte_filhos = 'tem filhos'
    else:
        parte_filhos = 'não tem filhos'

    frase_1 = f'{nome} tem {idade} anos e {parte_filhos}.'

    #experiencia profissional
    experiencia = dados['experiencia'].strip()
    if experiencia:
        frase_2= f' Relatou que possui experiência como {experiencia}'      
    else:
        frase_2 = f'Relatou não possuir experiência profissonal anterior.'

    escolaridade = dados ['escolaridade'].lower()
    curso = dados ['curso'].strip()

    if 'cursando' in escolaridade:
        frase_3 = f'. Possui graduação em {curso}.'
    else:
        frase_3 = f'Possui graduação em {curso}'    

    frase_4= (
        f" Nos testes de conhecimentos gerais que realizou, obteve "
        f"{dados['percentual_portugues']}% de aproveitamento em Português, "
        f"{dados['percentual_matematica']}% na prova de Matemática e "
        f"na prova de Raciocínio Lógico alcançou o total de "
        f"{dados['percentual_raciocinio']}%."
    )
    return ''.join([frase_1,frase_2,frase_3,frase_4])

def substituir_marcadores (documento:Document, dados: dict) -> None:
    for paragrafo in documento.paragraphs:
        for run in paragrafo.runs:
            for chave,valor in dados.items():
                marcador = '{{' + chave.upper() + '}}'
                if marcador in run.text:
                    run.text = run.text.replace (marcador,str(valor))

def gerar_documento (dados_formulario: dict, caminho_template: str, caminho_saida: str) -> str:
    documento = Document(caminho_template)

    aspectos_gerais = montar_aspectos_gerais(dados_formulario)

    dados_para_substituir = {
        "INSTITUICAO": dados_formulario["instituicao"],
        "FINALIDADE": dados_formulario["finalidade"],
        "NOME": dados_formulario["nome_completo"],
        "NOME_PRIMEIRO": dados_formulario["nome_primeiro"],
        "IDADE": dados_formulario["idade"],
        "SEXO": dados_formulario["sexo"],
        "ESTADO_CIVIL": dados_formulario["estado_civil"],
        "DATA_AVALIACAO": dados_formulario["data_avaliacao"],
        "CARGO": dados_formulario["cargo"],
        "ESCOLARIDADE": dados_formulario["escolaridade"],
        "FUNCAO": dados_formulario["funcao"],
        "ASPECTOS_GERAIS": aspectos_gerais,
        "RESULTADO": dados_formulario["resultado"],
        "DATA_ASSINATURA": dados_formulario["data_assinatura"],
    }

    substituir_marcadores (documento,dados_para_substituir)
    documento.save(caminho_saida)
    return caminho_saida
