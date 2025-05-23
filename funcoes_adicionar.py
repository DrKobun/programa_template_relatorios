from docxtpl import DocxTemplate

def adicionar_nome_executor(v):
    global nome_executor
    nome_executor = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['nome_executor'] = nome_executor
    doc.render(context)
    doc.save(nome_documento)

def adicionar_titulo_objeto(v):
    global titulo_objeto
    titulo_objeto = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['titulo_objeto'] = titulo_objeto
    doc.render(context)
    doc.save(nome_documento)

def adicionar_detalhamento_objeto(v):
    global detalhamento_objeto
    detalhamento_objeto = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['detalhamento_objeto'] = detalhamento_objeto
    doc.render(context)
    doc.save(nome_documento)

def adicionar_objetivo_principal(v):
    global objetivo_principal
    objetivo_principal = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['objetivo_principal'] = objetivo_principal
    doc.render(context)
    doc.save(nome_documento)

def adicionar_composicao_objeto(v):
    global composicao_objeto
    composicao_objeto = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['composicao_objeto'] = composicao_objeto
    doc.render(context)
    doc.save(nome_documento)

def adicionar_tipo_intervencao(v):
    global tipo_intervencao
    tipo_intervencao = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['tipo_intervencao'] = tipo_intervencao
    doc.render(context)
    doc.save(nome_documento)

def adicionar_tipo_obra(v):
    global tipo_obra
    tipo_obra = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['tipo_obra'] = tipo_obra
    doc.render(context)
    doc.save(nome_documento)

def adicionar_meta_fisica(v):
    global meta_fisica
    meta_fisica = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['meta_fisica'] = meta_fisica
    doc.render(context)
    doc.save(nome_documento)

def adicionar_unidade_medida(v):
    global unidade_medida
    unidade_medida = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['unidade_medida'] = unidade_medida
    doc.render(context)
    doc.save(nome_documento)

def adicionar_novo_pac(v):
    global novo_pac
    novo_pac = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['novo_pac'] = novo_pac
    doc.render(context)
    doc.save(nome_documento)

def adicionar_id_novo_pac(v):
    global id_novo_pac
    id_novo_pac = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['id_novo_pac'] = id_novo_pac
    doc.render(context)
    doc.save(nome_documento)

def adicionar_modalidade(v):
    global modalidade
    modalidade = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['modalidade'] = modalidade
    doc.render(context)
    doc.save(nome_documento)

def adicionar_normativo_principal(v):
    global normativo_principal
    normativo_principal = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['normativo_principal'] = normativo_principal
    doc.render(context)
    doc.save(nome_documento)

def adicionar_regime_simplificado(v):
    global regime_simplificado
    regime_simplificado = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['regime_simplificado'] = regime_simplificado
    doc.render(context)
    doc.save(nome_documento)

def adicionar_numero_instrumento(v):
    global numero_instrumento
    numero_instrumento = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['numero_instrumento'] = numero_instrumento
    doc.render(context)
    doc.save(nome_documento)

def adicionar_siafi(v):
    global siafi
    siafi = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['siafi'] = siafi
    doc.render(context)
    doc.save(nome_documento)

def adicionar_numero_proposta(v):
    global numero_proposta
    numero_proposta = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['numero_proposta'] = numero_proposta
    doc.render(context)
    doc.save(nome_documento)

def adicionar_status_instrumento(v):
    global status_instrumento
    status_instrumento = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['status_instrumento'] = status_instrumento
    doc.render(context)
    doc.save(nome_documento)

def adicionar_status_complementar(v):
    global status_complementar
    status_complementar = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['status_complementar'] = status_complementar
    doc.render(context)
    doc.save(nome_documento)

def adicionar_providencias(v):
    global providencias
    providencias = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['providencias'] = providencias
    doc.render(context)
    doc.save(nome_documento)

def adicionar_data_previdencias(v):
    global data_previdencias
    data_previdencias = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['data_previdencias'] = data_previdencias
    doc.render(context)
    doc.save(nome_documento)

def adicionar_restricoes(v):
    global restricoes
    restricoes = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['restricoes'] = restricoes
    doc.render(context)
    doc.save(nome_documento)

def adicionar_resultados(v):
    global resultados
    resultados = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['resultados'] = resultados
    doc.render(context)
    doc.save(nome_documento)

def adicionar_estagio_objeto(v):
    global estagio_objeto
    estagio_objeto = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['estagio_objeto'] = estagio_objeto
    doc.render(context)
    doc.save(nome_documento)

def adicionar_execucao_fisica(v):
    global execucao_fisica
    execucao_fisica = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['execucao_fisica'] = execucao_fisica
    doc.render(context)
    doc.save(nome_documento)

def adicionar_detalhamento_execucao(v):
    global detalhamento_execucao
    detalhamento_execucao = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['detalhamento_execucao'] = detalhamento_execucao
    doc.render(context)
    doc.save(nome_documento)

def adicionar_motivo_paralisacao(v):
    global motivo_paralisacao
    motivo_paralisacao = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['motivo_paralisacao'] = motivo_paralisacao
    doc.render(context)
    doc.save(nome_documento)

def adicionar_status_funcionalidade(v):
    global status_funcionalidade
    status_funcionalidade = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['status_funcionalidade'] = status_funcionalidade
    doc.render(context)
    doc.save(nome_documento)

def adicionar_execucao_adequada_pt(v):
    global execucao_adequada_pt
    execucao_adequada_pt = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['execucao_adequada_pt'] = execucao_adequada_pt
    doc.render(context)
    doc.save(nome_documento)

def adicionar_descricao_problemas(v):
    global descricao_problemas
    descricao_problemas = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['descricao_problemas'] = descricao_problemas
    doc.render(context)
    doc.save(nome_documento)

def adicionar_data_referencia_execucao(v):
    global data_referencia_execucao
    data_referencia_execucao = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['data_referencia_execucao'] = data_referencia_execucao
    doc.render(context)
    doc.save(nome_documento)

def adicionar_fonte_informacao_execucao(v):
    global fonte_informacao_execucao
    fonte_informacao_execucao = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['fonte_informacao_execucao'] = fonte_informacao_execucao
    doc.render(context)
    doc.save(nome_documento)

def adicionar_previsao_inicio(v):
    global previsao_inicio
    previsao_inicio = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['previsao_inicio'] = previsao_inicio
    doc.render(context)
    doc.save(nome_documento)

def adicionar_previsao_termino(v):
    global previsao_termino
    previsao_termino = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['previsao_termino'] = previsao_termino
    doc.render(context)
    doc.save(nome_documento)

def adicionar_inicio_objeto(v):
    global inicio_objeto
    inicio_objeto = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['inicio_objeto'] = inicio_objeto
    doc.render(context)
    doc.save(nome_documento)

def adicionar_referencia_inicio(v):
    global referencia_inicio
    referencia_inicio = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['referencia_inicio'] = referencia_inicio
    doc.render(context)
    doc.save(nome_documento)

def adicionar_termino_objeto(v):
    global termino_objeto
    termino_objeto = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['termino_objeto'] = termino_objeto
    doc.render(context)
    doc.save(nome_documento)

def adicionar_referencia_termino(v):
    global referencia_termino
    referencia_termino = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['referencia_termino'] = referencia_termino
    doc.render(context)
    doc.save(nome_documento)

def adicionar_inauguracao_objeto(v):
    global inauguracao_objeto
    inauguracao_objeto = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['inauguracao_objeto'] = inauguracao_objeto
    doc.render(context)
    doc.save(nome_documento)

def adicionar_data_recebimento_definitivo(v):
    global data_recebimento_definitivo
    data_recebimento_definitivo = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['data_recebimento_definitivo'] = data_recebimento_definitivo
    doc.render(context)
    doc.save(nome_documento)

def adicionar_visita_preliminar(v):
    global visita_preliminar
    visita_preliminar = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['visita_preliminar'] = visita_preliminar
    doc.render(context)
    doc.save(nome_documento)

def adicionar_visita_final(v):
    global visita_final
    visita_final = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['visita_final'] = visita_final
    doc.render(context)
    doc.save(nome_documento)

def adicionar_unidade_departamento(v):
    global unidade_departamento
    unidade_departamento = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['unidade_departamento'] = unidade_departamento
    doc.render(context)
    doc.save(nome_documento)

def adicionar_analista(v):
    global analista
    analista = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['analista'] = analista
    doc.render(context)
    doc.save(nome_documento)

def adicionar_data_assinatura(v):
    global data_assinatura
    data_assinatura = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['data_assinatura'] = data_assinatura
    doc.render(context)
    doc.save(nome_documento)

def adicionar_data_publicacao(v):
    global data_publicacao
    data_publicacao = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['data_publicacao'] = data_publicacao
    doc.render(context)
    doc.save(nome_documento)

def adicionar_inicio_vigencia(v):
    global inicio_vigencia
    inicio_vigencia = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['inicio_vigencia'] = inicio_vigencia
    doc.render(context)
    doc.save(nome_documento)

def adicionar_fim_vigencia_atual(v):
    global fim_vigencia_atual
    fim_vigencia_atual = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['fim_vigencia_atual'] = fim_vigencia_atual
    doc.render(context)
    doc.save(nome_documento)

# def adicionar_fim_vigencia_atual(v):
#     global fim_vigencia_atual
#     fim_vigencia_atual= v
#     nome_documento = 'teste_funcao.docx'
#     caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
#     doc = DocxTemplate(caminho_template)
#     if 'context' not in globals():
#         global context
#         context = {}
#     context['fim_vigencia_atual_(dias)'] = fim_vigencia_atual(dias)
#     doc.render(context)
#     doc.save(nome_documento)

def adicionar_fim_vigencia_original(v):
    global fim_vigencia_original
    fim_vigencia_original = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['fim_vigencia_original'] = fim_vigencia_original
    doc.render(context)
    doc.save(nome_documento)

def adicionar_valor_total_inicial(v):
    global valor_total_inicial
    valor_total_inicial = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['valor_total_inicial'] = valor_total_inicial
    doc.render(context)
    doc.save(nome_documento)

def adicionar_valor_total_atual(v):
    global valor_total_atual
    valor_total_atual = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['valor_total_atual'] = valor_total_atual
    doc.render(context)
    doc.save(nome_documento)

def adicionar_repasse_previsto_inicial(v):
    global repasse_previsto_inicial
    repasse_previsto_inicial = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['repasse_previsto_inicial'] = repasse_previsto_inicial
    doc.render(context)
    doc.save(nome_documento)

def adicionar_repasse_previsto_atual(v):
    global repasse_previsto_atual
    repasse_previsto_atual = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['repasse_previsto_atual'] = repasse_previsto_atual
    doc.render(context)
    doc.save(nome_documento)

def adicionar_contrapartida_inicial(v):
    global contrapartida_inicial
    contrapartida_inicial = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['contrapartida_inicial'] = contrapartida_inicial
    doc.render(context)
    doc.save(nome_documento)

def adicionar_contrapartida_atual(v):
    global contrapartida_atual
    contrapartida_atual = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['contrapartida_atual'] = contrapartida_atual
    doc.render(context)
    doc.save(nome_documento)

def adicionar_valor_referencia_repasse(v):
    global valor_referencia_repasse
    valor_referencia_repasse = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['valor_referencia_repasse'] = valor_referencia_repasse
    doc.render(context)
    doc.save(nome_documento)

def adicionar_total_empenhado(v):
    global total_empenhado
    total_empenhado = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['total_empenhado'] = total_empenhado
    doc.render(context)
    doc.save(nome_documento)

def adicionar_valor_a_empenhar(v):
    global valor_a_empenhar
    valor_a_empenhar = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['valor_a_empenhar'] = valor_a_empenhar
    doc.render(context)
    doc.save(nome_documento)

def adicionar_total_repassado(v):
    global total_repassado
    total_repassado = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['total_repassado'] = total_repassado
    doc.render(context)
    doc.save(nome_documento)

def adicionar_valor_a_repassar(v):
    global valor_a_repassar
    valor_a_repassar = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['valor_a_repassar'] = valor_a_repassar
    doc.render(context)
    doc.save(nome_documento)

def adicionar_total_aprovado_repasse(v):
    global total_aprovado_repasse
    total_aprovado_repasse = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['total_aprovado_repasse'] = total_aprovado_repasse
    doc.render(context)
    doc.save(nome_documento)

def adicionar_saldo_empenho(v):
    global saldo_empenho
    saldo_empenho = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['saldo_empenho'] = saldo_empenho
    doc.render(context)
    doc.save(nome_documento)

def adicionar_saldo_aprovado_repasse(v):
    global saldo_aprovado_repasse
    saldo_aprovado_repasse = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['saldo_aprovado_repasse'] = saldo_aprovado_repasse
    doc.render(context)
    doc.save(nome_documento)

def adicionar_valor_aporte_adicional(v):
    global valor_aporte_adicional
    valor_aporte_adicional = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['valor_aporte_adicional'] = valor_aporte_adicional
    doc.render(context)
    doc.save(nome_documento)

def adicionar_descricao_aporte_adicional(v):
    global descricao_aporte_adicional
    descricao_aporte_adicional = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['descricao_aporte_adicional'] = descricao_aporte_adicional
    doc.render(context)
    doc.save(nome_documento)

def adicionar_ano(v):
    global ano
    ano = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['ano'] = ano
    doc.render(context)
    doc.save(nome_documento)

def adicionar_programa(v):
    global programa
    programa = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['programa'] = programa
    doc.render(context)
    doc.save(nome_documento)

def adicionar_acao_ultima(v):
    global acao_ultima
    acao_ultima = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['acao_ultima'] = acao_ultima
    doc.render(context)
    doc.save(nome_documento)

def adicionar_municipios_beneficiados(v):
    global municipios_beneficiados
    municipios_beneficiados = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['municipios_beneficiados'] = municipios_beneficiados
    doc.render(context)
    doc.save(nome_documento)

def adicionar_qtd_municipios(v):
    global qtd_municipios
    qtd_municipios = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['qtd_municipios'] = qtd_municipios
    doc.render(context)
    doc.save(nome_documento)

def adicionar_populacao_beneficiada(v):
    global populacao_beneficiada
    populacao_beneficiada = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['populacao_beneficiada'] = populacao_beneficiada
    doc.render(context)
    doc.save(nome_documento)

def adicionar_descricao_beneficios(v):
    global descricao_beneficios
    descricao_beneficios = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['descricao_beneficios'] = descricao_beneficios
    doc.render(context)
    doc.save(nome_documento)

def adicionar_latitude(v):
    global latitude
    latitude = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['latitude'] = latitude
    doc.render(context)
    doc.save(nome_documento)

def adicionar_longitude(v):
    global longitude
    longitude = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['longitude'] = longitude
    doc.render(context)
    doc.save(nome_documento)

def adicionar_pi_130_2013(v):
    global pi_130_2013
    pi_130_2013 = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['pi_130_2013'] = pi_130_2013
    doc.render(context)
    doc.save(nome_documento)

def adicionar_prazo_suspensiva_original(v):
    global prazo_suspensiva_original
    prazo_suspensiva_original = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['prazo_suspensiva_original'] = prazo_suspensiva_original
    doc.render(context)
    doc.save(nome_documento)

def adicionar_prazo_suspensiva_vigente(v):
    global prazo_suspensiva_vigente
    prazo_suspensiva_vigente = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['prazo_suspensiva_vigente'] = prazo_suspensiva_vigente
    doc.render(context)
    doc.save(nome_documento)

def adicionar_data_abertura_processo(v):
    global data_abertura_processo
    data_abertura_processo = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['data_abertura_processo'] = data_abertura_processo
    doc.render(context)
    doc.save(nome_documento)

def adicionar_data_apresentacao_proposta(v):
    global data_apresentacao_proposta
    data_apresentacao_proposta = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['data_apresentacao_proposta'] = data_apresentacao_proposta
    doc.render(context)
    doc.save(nome_documento)

def adicionar_data_retirada_suspensiva(v):
    global data_retirada_suspensiva
    data_retirada_suspensiva = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['data_retirada_suspensiva'] = data_retirada_suspensiva
    doc.render(context)
    doc.save(nome_documento)

def adicionar_data_inicio_supervisao(v):
    global data_inicio_supervisao
    data_inicio_supervisao = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['data_inicio_supervisao'] = data_inicio_supervisao
    doc.render(context)
    doc.save(nome_documento)

def adicionar_data_analise_pcf(v):
    global data_analise_pcf
    data_analise_pcf = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['data_analise_pcf'] = data_analise_pcf
    doc.render(context)
    doc.save(nome_documento)

def adicionar_doh(v):
    global doh
    doh = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['doh'] = doh
    doc.render(context)
    doc.save(nome_documento)

def adicionar_ativo(v):
    global ativo
    ativo = v
    nome_documento = 'teste_funcao.docx'
    caminho_template = 'C:\\Users\\walyson.ferreira\\Desktop\\openpy\\template\\template.docx'
    doc = DocxTemplate(caminho_template)
    if 'context' not in globals():
        global context
        context = {}
    context['ativo'] = ativo
    doc.render(context)
    doc.save(nome_documento)