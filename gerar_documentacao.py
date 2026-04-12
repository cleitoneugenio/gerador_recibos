#!/usr/bin/env python3
"""
Gera a documentação do projeto gerar_recibos.py em PDF.
Uso: python gerar_documentacao.py
"""

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------

def criar_estilos() -> dict:
    return {
        'capa_titulo': ParagraphStyle(
            'capa_titulo',
            fontName='Helvetica-Bold',
            fontSize=22,
            alignment=TA_CENTER,
            textColor=colors.HexColor('#1a1a2e'),
            spaceAfter=10,
        ),
        'capa_sub': ParagraphStyle(
            'capa_sub',
            fontName='Helvetica',
            fontSize=13,
            alignment=TA_CENTER,
            textColor=colors.HexColor('#444444'),
            spaceAfter=6,
        ),
        'secao': ParagraphStyle(
            'secao',
            fontName='Helvetica-Bold',
            fontSize=14,
            textColor=colors.HexColor('#1a1a2e'),
            spaceBefore=18,
            spaceAfter=6,
            borderPad=4,
        ),
        'subsecao': ParagraphStyle(
            'subsecao',
            fontName='Helvetica-Bold',
            fontSize=11,
            textColor=colors.HexColor('#2e4057'),
            spaceBefore=10,
            spaceAfter=4,
        ),
        'corpo': ParagraphStyle(
            'corpo',
            fontName='Helvetica',
            fontSize=10,
            alignment=TA_JUSTIFY,
            leading=16,
            spaceAfter=6,
        ),
        'codigo': ParagraphStyle(
            'codigo',
            fontName='Courier',
            fontSize=9,
            leading=14,
            backColor=colors.HexColor('#f4f4f4'),
            leftIndent=12,
            rightIndent=12,
            spaceBefore=4,
            spaceAfter=4,
        ),
        'topico': ParagraphStyle(
            'topico',
            fontName='Helvetica',
            fontSize=10,
            leading=16,
            leftIndent=14,
            spaceAfter=3,
        ),
        'rodape': ParagraphStyle(
            'rodape',
            fontName='Helvetica',
            fontSize=8,
            alignment=TA_CENTER,
            textColor=colors.grey,
        ),
    }


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def divisor(cor=colors.HexColor('#cccccc')) -> list:
    return [
        Spacer(1, 0.3 * cm),
        HRFlowable(width='100%', thickness=0.5, color=cor),
        Spacer(1, 0.3 * cm),
    ]


def bloco_codigo(linhas: list, s: dict) -> list:
    elementos = []
    for linha in linhas:
        elementos.append(Paragraph(linha, s['codigo']))
    return elementos


# ---------------------------------------------------------------------------
# Seções do documento
# ---------------------------------------------------------------------------

def secao_capa(s: dict) -> list:
    return [
        Spacer(1, 2 * cm),
        Paragraph('DOCUMENTAÇÃO TÉCNICA', s['capa_titulo']),
        Paragraph('Gerador de Recibos de Prestação de Serviço', s['capa_sub']),
        Spacer(1, 0.4 * cm),
        HRFlowable(width='60%', thickness=2, color=colors.HexColor('#1a1a2e'), hAlign='CENTER'),
        Spacer(1, 0.6 * cm),
        Paragraph('Arquivo: <b>gerar_recibos.py</b>', s['capa_sub']),
        Paragraph('Linguagem: <b>Python 3</b>', s['capa_sub']),
        Paragraph('Bibliotecas: <b>pandas · ReportLab · num2words</b>', s['capa_sub']),
        Spacer(1, 3 * cm),
        Paragraph(
            'Este documento explica passo a passo como o programa foi construído, '
            'o motivo de cada escolha técnica e como utilizá-lo no dia a dia.',
            s['corpo'],
        ),
    ]


def secao_objetivo(s: dict) -> list:
    return [
        *divisor(),
        Paragraph('1. OBJETIVO DO PROGRAMA', s['secao']),
        Paragraph(
            'O programa nasceu da necessidade de eliminar o trabalho manual de criar '
            'recibos de pagamento toda semana para os funcionários da empresa '
            '<b>GF MUNIZ ARTEFACTOS DE CERAMICA EIRELI</b>.',
            s['corpo'],
        ),
        Paragraph(
            'Toda semana existe uma planilha Excel (<b>.xlsx</b>) com os nomes dos '
            'funcionários e os valores de diária de segunda a sábado. O programa lê '
            'essa planilha automaticamente e gera um único arquivo PDF com todos os '
            'recibos prontos para impressão e assinatura.',
            s['corpo'],
        ),
        Paragraph('<b>Problema resolvido:</b>', s['subsecao']),
        Paragraph('• Criar recibo por recibo manualmente levava muito tempo.', s['topico']),
        Paragraph('• Risco de erro humano ao copiar nomes e valores.', s['topico']),
        Paragraph('• Dificuldade em escrever o valor por extenso corretamente.', s['topico']),
        Paragraph('<b>Solução entregue:</b>', s['subsecao']),
        Paragraph('• Um único comando gera os ~20 recibos em segundos.', s['topico']),
        Paragraph('• Valor por extenso gerado automaticamente em português.', s['topico']),
        Paragraph('• Data preenchida automaticamente com o dia da geração.', s['topico']),
        Paragraph('• Funcionários sem trabalho na semana são ignorados automaticamente.', s['topico']),
    ]


def secao_bibliotecas(s: dict) -> list:
    dados_tabela = [
        ['Biblioteca', 'Para que serve', 'Por que foi escolhida'],
        ['pandas', 'Ler e processar o arquivo .xlsx', 'Biblioteca padrão para dados em Python; lida com Excel de forma simples e robusta'],
        ['openpyxl', 'Motor de leitura de arquivos .xlsx', 'Exigida pelo pandas para abrir arquivos Excel modernos'],
        ['ReportLab', 'Gerar o arquivo PDF', 'Biblioteca profissional para criação de PDFs em Python; permite controle total do layout'],
        ['num2words', 'Converter número em texto por extenso', 'Suporte nativo ao português brasileiro (pt_BR); preciso e fácil de usar'],
        ['os / sys', 'Lidar com arquivos e argumentos da linha de comando', 'Módulos nativos do Python; sem instalação extra'],
        ['datetime', 'Obter a data atual do sistema', 'Módulo nativo do Python; fornece a data de geração do recibo automaticamente'],
    ]

    tabela = Table(
        dados_tabela,
        colWidths=[3.2 * cm, 4.5 * cm, 8.8 * cm],
        repeatRows=1,
    )
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a1a2e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f4f8')]),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cccccc')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('WORDWRAP', (0, 0), (-1, -1), True),
    ]))

    return [
        *divisor(),
        Paragraph('2. BIBLIOTECAS UTILIZADAS', s['secao']),
        Paragraph(
            'Antes de rodar o programa pela primeira vez, é necessário instalar as '
            'bibliotecas externas com o comando abaixo no terminal:',
            s['corpo'],
        ),
        *bloco_codigo(
            ['pip install pandas openpyxl reportlab num2words'],
            s,
        ),
        Spacer(1, 0.4 * cm),
        tabela,
    ]


def secao_estrutura(s: dict) -> list:
    return [
        *divisor(),
        Paragraph('3. ESTRUTURA DO PROGRAMA', s['secao']),
        Paragraph(
            'O programa é dividido em funções, cada uma com uma responsabilidade '
            'específica. Essa separação torna o código mais fácil de entender, '
            'testar e modificar no futuro.',
            s['corpo'],
        ),

        Paragraph('3.1 — valor_por_extenso(valor)', s['subsecao']),
        Paragraph(
            'Recebe um número (ex: <b>395.0</b>) e retorna uma string formatada com '
            'o valor em reais e o texto por extenso em português.',
            s['corpo'],
        ),
        *bloco_codigo([
            'def valor_por_extenso(valor: float) -> str:',
            '    reais = int(valor)',
            '    centavos = round((valor - reais) * 100)',
            "    extenso = num2words(reais, lang='pt_BR') + ' reais'",
            '    if centavos > 0:',
            "        extenso += ' e ' + num2words(centavos, lang='pt_BR') + ' centavos'",
            "    valor_fmt = f'R$ {valor:,.2f}'.replace(',','X').replace('.',',').replace('X','.')",
            "    return f'{valor_fmt} ({extenso})'",
        ], s),
        Paragraph(
            '<b>Por que assim?</b> O formato brasileiro usa vírgula como separador decimal '
            '(R$ 395,00), mas o Python usa ponto. Por isso fazemos a substituição manualmente: '
            'primeiro trocamos a vírgula por "X" (temporário), depois o ponto por vírgula, '
            'e por fim o "X" por ponto — resultando no padrão correto.',
            s['corpo'],
        ),

        Paragraph('3.2 — formatar_data(data)', s['subsecao']),
        Paragraph(
            'Converte um objeto <b>datetime</b> para o formato por extenso em português, '
            'como: <b>11 de abril de 2026</b>.',
            s['corpo'],
        ),
        *bloco_codigo([
            'def formatar_data(data: datetime) -> str:',
            "    meses = ['janeiro', 'fevereiro', 'março', ...]",
            "    return f'{data.day} de {meses[data.month - 1]} de {data.year}'",
        ], s),
        Paragraph(
            '<b>Por que assim?</b> O Python não tem suporte nativo a nomes de meses em '
            'português sem configurar o locale do sistema operacional (o que pode falhar '
            'dependendo do Windows). A lista manual garante que funciona em qualquer máquina.',
            s['corpo'],
        ),

        Paragraph('3.3 — ler_funcionarios(arquivo)', s['subsecao']),
        Paragraph(
            'Lê o arquivo <b>.xlsx</b> com o pandas e retorna uma lista com nome e total '
            'de cada funcionário que trabalhou na semana.',
            s['corpo'],
        ),
        *bloco_codigo([
            'def ler_funcionarios(arquivo: str) -> list:',
            '    df = pd.read_excel(arquivo, header=0)',
            '    for _, row in df.iterrows():',
            '        try:',
            '            int(row.iloc[0])   # col 0 = número do funcionário',
            '        except (ValueError, TypeError):',
            '            continue           # pula cabeçalho e linha TOTAL',
            '        nome  = str(row.iloc[1]).strip()   # col 1 = nome',
            '        total = float(row.iloc[10])        # col 10 = total',
            '        if total <= 0: continue',
            "        funcionarios.append({'nome': nome, 'total': total})",
        ], s),
        Paragraph(
            '<b>Por que header=0?</b> A planilha tem uma linha de título no topo '
            '("CONTROLE DE PONTO...") que o pandas usa automaticamente como nome das colunas. '
            'Isso nos permite identificar as linhas de dados pelo conteúdo da primeira coluna.',
            s['corpo'],
        ),
        Paragraph(
            '<b>Por que tentar converter col[0] para int?</b> As linhas de dados têm '
            'número de funcionário (1, 2, 3...). A linha de cabeçalho tem "#" e a linha '
            'final tem "TOTAL" — ambas falham no int() e são ignoradas automaticamente.',
            s['corpo'],
        ),
        Paragraph(
            '<b>Por que ignorar total <= 0?</b> Funcionários que não trabalharam nenhum dia '
            'na semana ficam com total zero. Não faz sentido gerar recibo de R$ 0,00.',
            s['corpo'],
        ),

        Paragraph('3.4 — montar_recibo(nome, total, data_str, styles)', s['subsecao']),
        Paragraph(
            'Monta a lista de elementos visuais de um recibo individual: título, '
            'texto do corpo, linha de assinatura e nome do funcionário.',
            s['corpo'],
        ),
        *bloco_codigo([
            'def montar_recibo(...) -> list:',
            '    return [',
            "        Paragraph('<u><b>RECIBO DE PRESTAÇÃO DE SERVIÇO</b></u>', titulo),",
            '        Spacer(1, 0.5 * cm),',
            '        Paragraph(corpo_texto, corpo),',
            '        Spacer(1, 0.4 * cm),',
            "        Paragraph(f'Bela Cruz, {data_str}.', corpo),",
            '        Spacer(1, 1.8 * cm),',
            "        HRFlowable(width='55%', ...),   # linha de assinatura",
            '        Spacer(1, 0.25 * cm),',
            '        Paragraph(nome, assinatura),',
            '    ]',
        ], s),
        Paragraph(
            '<b>Por que retornar uma lista?</b> O ReportLab trabalha com uma lista de '
            '"flowables" (elementos que fluem pela página). Retornar a lista permite '
            'envolvê-la em um <b>KeepTogether</b>, que garante que os elementos de um '
            'mesmo recibo nunca sejam separados entre duas páginas.',
            s['corpo'],
        ),
        Paragraph(
            '<b>Por que &lt;u&gt; e &lt;b&gt; no título?</b> O ReportLab aceita '
            'marcação XML dentro do texto dos parágrafos. A tag &lt;u&gt; aplica '
            'sublinhado e &lt;b&gt; aplica negrito — exatamente o formato pedido para o título.',
            s['corpo'],
        ),

        Paragraph('3.5 — gerar_pdf(arquivo_xlsx, arquivo_pdf)', s['subsecao']),
        Paragraph(
            'Função principal: lê os funcionários, monta todos os recibos e constrói o PDF.',
            s['corpo'],
        ),
        *bloco_codigo([
            'def gerar_pdf(arquivo_xlsx, arquivo_pdf):',
            '    hoje = datetime.today()',
            '    funcionarios = ler_funcionarios(arquivo_xlsx)',
            '    doc = SimpleDocTemplate(arquivo_pdf, pagesize=A4, ...)',
            '    story = []',
            '    for i, func in enumerate(funcionarios):',
            '        recibo = montar_recibo(...)',
            '        story.append(KeepTogether(recibo))     # mantém recibo unido',
            '        if i < len(funcionarios) - 1:',
            '            story.append(HRFlowable(dash=(6,4)))  # linha tracejada',
            '    doc.build(story)',
        ], s),
        Paragraph(
            '<b>Por que KeepTogether?</b> Sem ele, um recibo poderia ser cortado no meio '
            'entre uma página e outra. O KeepTogether move o bloco inteiro para a página '
            'seguinte caso não caiba no espaço restante.',
            s['corpo'],
        ),
        Paragraph(
            '<b>Por que a linha tracejada (dash)?</b> Serve como guia visual para separar '
            'os recibos ao imprimir e recortar, facilitando a distribuição para cada funcionário.',
            s['corpo'],
        ),

        Paragraph('3.6 — main()', s['subsecao']),
        Paragraph(
            'Ponto de entrada do programa. Lê os argumentos passados na linha de comando '
            'e chama a função gerar_pdf.',
            s['corpo'],
        ),
        *bloco_codigo([
            'def main():',
            '    arquivo_xlsx = sys.argv[1]',
            '    arquivo_pdf  = sys.argv[2] if len(sys.argv) >= 3 else',
            "                   os.path.splitext(arquivo_xlsx)[0] + '_recibos.pdf'",
            '    gerar_pdf(arquivo_xlsx, arquivo_pdf)',
        ], s),
        Paragraph(
            '<b>Por que sys.argv?</b> Permite passar o nome do arquivo diretamente no '
            'terminal, tornando o programa flexível para trabalhar com qualquer planilha '
            '— não só a da semana atual.',
            s['corpo'],
        ),
    ]


def secao_layout(s: dict) -> list:
    return [
        *divisor(),
        Paragraph('4. LAYOUT E ESTILOS DO PDF', s['secao']),
        Paragraph(
            'O ReportLab utiliza o conceito de <b>ParagraphStyle</b> para definir a '
            'aparência de cada bloco de texto. Três estilos foram criados:',
            s['corpo'],
        ),
        Paragraph('• <b>titulo</b> — Helvetica-Bold, 13pt, centralizado.', s['topico']),
        Paragraph('• <b>corpo</b> — Helvetica, 10pt, justificado, entrelinha 15pt.', s['topico']),
        Paragraph('• <b>assinatura</b> — Helvetica, 10pt, centralizado.', s['topico']),
        Spacer(1, 0.3 * cm),
        Paragraph(
            'A página usa margens de <b>2 cm</b> nas laterais e <b>1,5 cm</b> no topo e '
            'rodapé, aproveitando bem o espaço do A4 para caber o maior número possível '
            'de recibos por página (em média 2 a 3, dependendo do comprimento do nome).',
            s['corpo'],
        ),
    ]


def secao_uso(s: dict) -> list:
    return [
        *divisor(),
        Paragraph('5. COMO USAR O PROGRAMA', s['secao']),

        Paragraph('Passo 1 — Instalar as dependências (apenas uma vez)', s['subsecao']),
        *bloco_codigo(['pip install pandas openpyxl reportlab num2words'], s),

        Paragraph('Passo 2 — Abrir o terminal na pasta do projeto', s['subsecao']),
        Paragraph(
            'No Windows Explorer, navegue até a pasta do projeto, clique na barra de '
            'endereço, digite <b>cmd</b> e pressione Enter.',
            s['corpo'],
        ),

        Paragraph('Passo 3 — Executar o programa', s['subsecao']),
        Paragraph('<b>Opção A</b> — O PDF é gerado com o mesmo nome da planilha:', s['topico']),
        *bloco_codigo(['python gerar_recibos.py ponto_cedan_semana.xlsx'], s),
        Paragraph('<b>Opção B</b> — Definindo o nome do PDF de saída:', s['topico']),
        *bloco_codigo(['python gerar_recibos.py ponto_cedan_semana.xlsx recibos_semana1.pdf'], s),

        Paragraph('Passo 4 — Verificar o resultado', s['subsecao']),
        Paragraph(
            'O PDF é salvo na mesma pasta da planilha. Abra-o, confira os recibos '
            'e imprima normalmente.',
            s['corpo'],
        ),

        Paragraph('Regras automáticas do programa', s['subsecao']),
        Paragraph('• Funcionários com total R$ 0,00 são ignorados (não gera recibo).', s['topico']),
        Paragraph('• A data é sempre a do dia em que o programa é executado.', s['topico']),
        Paragraph('• O valor por extenso é calculado automaticamente em português.', s['topico']),
        Paragraph('• Se um recibo não couber na página atual, vai para a próxima.', s['topico']),
    ]


def secao_estrutura_planilha(s: dict) -> list:
    dados_tabela = [
        ['Coluna', 'Índice', 'Conteúdo esperado'],
        ['#', '0', 'Número do funcionário (inteiro). Usado para identificar linhas de dados.'],
        ['FUNCIONÁRIOS', '1', 'Nome completo do funcionário.'],
        ['SEGUNDA', '2', 'Valor da diária ou "F" (falta).'],
        ['TERÇA', '3', 'Valor da diária ou "F" (falta).'],
        ['QUARTA', '4', 'Valor da diária ou "F" (falta).'],
        ['QUINTA', '5', 'Valor da diária ou "F" (falta).'],
        ['SEXTA', '6', 'Valor da diária ou "F" (falta).'],
        ['SÁBADO', '7', 'Valor da diária ou "F" (falta).'],
        ['DIAS', '8', 'Total de dias trabalhados.'],
        ['BÔNUS', '9', 'Bônus (quando aplicável).'],
        ['TOTAL', '10', 'Total a receber. É este valor que aparece no recibo.'],
    ]

    tabela = Table(
        dados_tabela,
        colWidths=[2.8 * cm, 1.8 * cm, 12 * cm],
        repeatRows=1,
    )
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2e4057')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f4f8')]),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cccccc')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))

    return [
        *divisor(),
        Paragraph('6. ESTRUTURA ESPERADA DA PLANILHA', s['secao']),
        Paragraph(
            'O programa foi desenvolvido para ler planilhas no seguinte formato:',
            s['corpo'],
        ),
        tabela,
        Spacer(1, 0.4 * cm),
        Paragraph(
            '<b>Atenção:</b> a planilha deve ter uma linha de título mesclado no topo '
            '(ex: "CONTROLE DE PONTO") e uma linha de cabeçalho com os nomes das colunas. '
            'A última linha com "TOTAL" é ignorada automaticamente.',
            s['corpo'],
        ),
    ]


def secao_gui(s: dict) -> list:
    dados_componentes = [
        ['Componente', 'Widget tkinter', 'Justificativa'],
        ['Rótulos', 'tk.Label', 'Texto estático para identificar cada campo da tela.'],
        ['Campos de texto', 'tk.Entry + StringVar', 'Permitem leitura e escrita do caminho do arquivo. StringVar separa modelo da visão.'],
        ['Botões de arquivo', 'tk.Button + filedialog', 'Abrem o diálogo nativo do Windows — familiar ao usuário, sem necessidade de digitar o caminho.'],
        ['Barra de progresso', 'ttk.Progressbar', 'Indica que o programa está trabalhando. Modo indeterminate correto pois não há percentual real.'],
        ['Label de status', 'tk.Label + StringVar', 'Feedback textual contínuo sem interromper o fluxo com caixas de diálogo.'],
        ['Botão principal', 'tk.Button', 'Cor verde (#2e7d32) destaca a ação principal. Desabilitado durante execução para evitar cliques duplos.'],
        ['Diálogo de conclusão', 'messagebox.askyesno', 'Pergunta se o usuário quer abrir o PDF. Não abre automaticamente para não interromper quem gera vários arquivos.'],
    ]

    tabela = Table(
        dados_componentes,
        colWidths=[3.5 * cm, 3.5 * cm, 9.5 * cm],
        repeatRows=1,
    )
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2e7d32')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f4f8')]),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cccccc')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))

    return [
        *divisor(),
        Paragraph('7. INTERFACE GRÁFICA (GUI) — IMPLEMENTAÇÃO PASSO A PASSO', s['secao']),
        Paragraph(
            'A interface gráfica foi adicionada ao script para eliminar a necessidade do terminal. '
            'O comportamento é preservado: sem argumentos → abre a janela; com argumentos → modo '
            'CLI original. Toda a lógica de geração de PDF permanece inalterada.',
            s['corpo'],
        ),

        # ── PASSO 1 ──
        Paragraph('Passo 1 — Escolha da biblioteca: tkinter', s['subsecao']),
        Paragraph(
            'Foram avaliadas quatro opções de biblioteca gráfica para Python:',
            s['corpo'],
        ),
        Paragraph(
            '• <b>tkinter</b> ✔ — já vem embutida no Python; sem instalação; usa controles '
            'nativos do Windows; estável e bem documentada.',
            s['topico'],
        ),
        Paragraph(
            '• <b>PyQt / PySide6</b> ✘ — exige instalação (~60 MB); licença LGPL com '
            'restrições; complexidade desnecessária para um formulário simples.',
            s['topico'],
        ),
        Paragraph(
            '• <b>wxPython</b> ✘ — instalação adicional; menos popular; documentação mais escassa.',
            s['topico'],
        ),
        Paragraph(
            '• <b>Dear PyGui</b> ✘ — voltado para dashboards e jogos; API inadequada para '
            'formulários de escritório.',
            s['topico'],
        ),
        Paragraph(
            '<b>Decisão:</b> tkinter foi escolhido por custo zero de instalação e suficiência '
            'total para o caso de uso.',
            s['corpo'],
        ),

        # ── PASSO 2 ──
        Paragraph('Passo 2 — Criação da janela principal', s['subsecao']),
        *bloco_codigo([
            'root = tk.Tk()',
            'root.title("Gerador de Recibos")',
            'root.resizable(False, False)',
        ], s),
        Paragraph(
            '<b>resizable(False, False):</b> A janela tem largura fixa definida pelos Entry de '
            'largura 50 caracteres. Permitir redimensionamento sem lógica de responsividade '
            'distorceria os widgets — fixar o tamanho é mais simples e correto para este layout.',
            s['corpo'],
        ),

        # ── PASSO 3 ──
        Paragraph('Passo 3 — Campos de entrada com StringVar', s['subsecao']),
        *bloco_codigo([
            'var_xlsx = tk.StringVar()',
            'entry_xlsx = tk.Entry(root, textvariable=var_xlsx, width=50)',
        ], s),
        Paragraph(
            '<b>Por que StringVar?</b> O tkinter segue o padrão MVC: a <i>StringVar</i> é o '
            'modelo e o <i>Entry</i> é a visão. Ler ou escrever o valor via '
            '<i>var_xlsx.get() / var_xlsx.set()</i> evita referenciar diretamente o widget, '
            'tornando o código mais limpo e testável.',
            s['corpo'],
        ),

        # ── PASSO 4 ──
        Paragraph('Passo 4 — Diálogos nativos de arquivo', s['subsecao']),
        *bloco_codigo([
            '# Abrir planilha:',
            "caminho = filedialog.askopenfilename(",
            "    title='Selecionar planilha',",
            "    filetypes=[('Excel', '*.xlsx *.xls'), ('Todos', '*.*')],",
            ')',
            '',
            '# Salvar PDF:',
            "caminho = filedialog.asksaveasfilename(",
            "    title='Salvar PDF como',",
            "    defaultextension='.pdf',",
            "    filetypes=[('PDF', '*.pdf')],",
            ')',
        ], s),
        Paragraph(
            '<b>askopenfilename vs asksaveasfilename:</b> O primeiro só aceita arquivos '
            'existentes; o segundo permite digitar um nome novo. O parâmetro '
            '<i>defaultextension=".pdf"</i> garante a extensão correta mesmo que o usuário '
            'não a escreva.',
            s['corpo'],
        ),
        Paragraph(
            '<b>Preenchimento automático do PDF:</b> Ao selecionar a planilha, o caminho do '
            'PDF é preenchido automaticamente (mesmo nome, extensão _recibos.pdf) se o campo '
            'ainda estiver vazio. Reduz o número de interações necessárias no caso mais comum.',
            s['corpo'],
        ),

        # ── PASSO 5 ──
        Paragraph('Passo 5 — Barra de progresso no modo indeterminate', s['subsecao']),
        *bloco_codigo([
            "progress = ttk.Progressbar(root, mode='indeterminate', length=400)",
            '# Iniciar animação:',
            'progress.start(10)   # intervalo de 10ms entre quadros',
            '# Parar:',
            'progress.stop()',
        ], s),
        Paragraph(
            '<b>Por que mode="indeterminate"?</b> A geração do PDF não reporta etapas '
            'intermediárias — só se sabe quando começa e quando termina. O modo '
            '<i>determinate</i> (barra que vai de 0% a 100%) seria semanticamente errado '
            'pois induziria o usuário a inferir um percentual inexistente.',
            s['corpo'],
        ),
        Paragraph(
            '<b>Por que ttk e não tk?</b> O widget <i>tk.Progressbar</i> não existe. '
            'O <i>ttk.Progressbar</i> é a implementação oficial, com visual integrado ao '
            'tema do Windows.',
            s['corpo'],
        ),

        # ── PASSO 6 ──
        Paragraph('Passo 6 — Execução em thread separada (ponto crítico)', s['subsecao']),
        Paragraph(
            'Este é o componente mais importante da implementação. O tkinter roda em uma '
            'única thread e processa todos os eventos (cliques, redesenhos, animações) '
            'em seu <i>event loop</i>. Se <b>gerar_pdf()</b> fosse chamada diretamente '
            'no botão, a janela congelaria completamente até o término — a barra de '
            'progresso pararia e o Windows marcaria o programa como "Não respondendo".',
            s['corpo'],
        ),
        *bloco_codigo([
            'def executar():',
            '    # validações de entrada...',
            '    btn_gerar.config(state="disabled")  # bloqueia novo clique',
            '    progress.start(10)',
            '    def tarefa():',
            '        try:',
            '            gerar_pdf(xlsx, pdf)',
            '            root.after(0, lambda: concluido(pdf))',
            '        except Exception as exc:',
            '            root.after(0, lambda: erro(str(exc)))',
            '    threading.Thread(target=tarefa, daemon=True).start()',
        ], s),
        Paragraph(
            '<b>threading.Thread:</b> Módulo da biblioteca padrão. Suficiente para tarefas '
            'I/O-bound como escrita de arquivo — não exige asyncio nem concurrent.futures.',
            s['corpo'],
        ),
        Paragraph(
            '<b>daemon=True:</b> A thread é marcada como daemon para encerrar automaticamente '
            'caso o usuário feche a janela antes do término, evitando processos zumbis.',
            s['corpo'],
        ),
        Paragraph(
            '<b>root.after(0, callback):</b> O tkinter não é thread-safe — widgets '
            'NUNCA devem ser atualizados de dentro de outra thread. O método '
            '<i>root.after(0, callback)</i> agenda a execução do callback no event loop '
            'principal, tornando a atualização segura.',
            s['corpo'],
        ),
        Paragraph(
            '<b>try/except dentro da thread:</b> Erros em threads secundárias não '
            'propagam para a thread principal — eles simplesmente desaparecem. O bloco '
            'try/except captura a exceção e a envia ao event loop via root.after para '
            'exibição ao usuário.',
            s['corpo'],
        ),

        # ── PASSO 7 ──
        Paragraph('Passo 7 — Diálogo pós-conclusão e abertura do arquivo', s['subsecao']),
        *bloco_codigo([
            "if messagebox.askyesno('Concluído',",
            "                       f'PDF gerado:\\n{pdf}\\n\\nDeseja abrir o arquivo?'):",
            '    os.startfile(pdf)',
        ], s),
        Paragraph(
            '<b>messagebox.askyesno:</b> Preferido sobre abrir o arquivo automaticamente '
            'para respeitar o fluxo do usuário — ele pode gerar vários PDFs em sequência '
            'e não querer interrupções a cada um.',
            s['corpo'],
        ),
        Paragraph(
            '<b>os.startfile:</b> Abre o PDF com o programa padrão do sistema (Reader, '
            'Edge, etc.) sem hardcodar um aplicativo específico. É a abordagem idiomática '
            'no Windows.',
            s['corpo'],
        ),

        # ── PASSO 8 ──
        Paragraph('Passo 8 — Compatibilidade com o modo CLI', s['subsecao']),
        *bloco_codigo([
            'def main():',
            '    if len(sys.argv) < 2:',
            '        iniciar_gui()   # sem argumentos → abre a janela',
            '        return',
            '    # com argumentos → comportamento CLI original inalterado',
        ], s),
        Paragraph(
            'A GUI é uma camada adicionada <b>sobre</b> a lógica existente, não uma '
            'substituição. Scripts de automação, agendamentos via Task Scheduler e outros '
            'usos programáticos continuam funcionando sem nenhuma alteração.',
            s['corpo'],
        ),

        # ── TABELA RESUMO ──
        Spacer(1, 0.4 * cm),
        Paragraph('Resumo dos componentes', s['subsecao']),
        tabela,
    ]


def secao_executavel(s: dict) -> list:
    dados_flags = [
        ['Flag / Opção', 'O que faz', 'Por que foi usada'],
        ['--onefile', 'Empacota tudo em um único .exe', 'O usuário recebe um só arquivo para copiar e usar — sem pastas auxiliares, sem risco de esquecer alguma DLL.'],
        ['--windowed', 'Suprime a janela de terminal (cmd)', 'Como o programa tem GUI, o terminal aberto junto seria distração visual e assustaria usuários não técnicos.'],
        ['--name "GeradorRecibos"', 'Define o nome do executável gerado', 'Sem esta flag o nome seria gerar_recibos.exe (nome do script). O nome amigável fica mais claro para o usuário final.'],
    ]

    tabela_flags = Table(
        dados_flags,
        colWidths=[3.8 * cm, 4.2 * cm, 8.5 * cm],
        repeatRows=1,
    )
    tabela_flags.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a237e')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f4f8')]),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cccccc')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))

    dados_ferramentas = [
        ['Ferramenta', 'Abordagem', 'Veredicto'],
        ['PyInstaller', 'Empacota o interpretador Python + dependências + script em um exe', '✔ Escolhido — maduro, amplamente documentado, suporte nativo a tkinter/pandas/reportlab.'],
        ['cx_Freeze', 'Gera uma pasta com exe + DLLs (não onefile por padrão)', '✘ Onefile mais trabalhoso; menos hooks automáticos para bibliotecas científicas.'],
        ['Nuitka', 'Compila Python para C e gera exe nativo', '✘ Requer compilador C instalado (MSVC/MinGW); build muito mais lento; complexidade desnecessária para este caso.'],
        ['py2exe', 'Clássico empacotador Windows', '✘ Abandonado para Python 3.9+; sem suporte ao Python 3.14 usado no projeto.'],
    ]

    tabela_ferramentas = Table(
        dados_ferramentas,
        colWidths=[2.8 * cm, 5.5 * cm, 8.2 * cm],
        repeatRows=1,
    )
    tabela_ferramentas.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#37474f')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f4f8')]),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cccccc')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))

    return [
        *divisor(),
        Paragraph('8. GERAÇÃO DO EXECUTÁVEL WINDOWS (.exe) — PASSO A PASSO', s['secao']),
        Paragraph(
            'Para que o programa funcione em qualquer computador Windows sem precisar instalar '
            'o Python, ele foi convertido em um arquivo executável (.exe) usando a ferramenta '
            '<b>PyInstaller</b>. Esta seção documenta cada decisão tomada durante esse processo.',
            s['corpo'],
        ),

        # ── PASSO 1 ──
        Paragraph('Passo 1 — Escolha da ferramenta de empacotamento', s['subsecao']),
        Paragraph(
            'Existem quatro ferramentas principais para converter scripts Python em executáveis '
            'Windows. Todas foram avaliadas:',
            s['corpo'],
        ),
        tabela_ferramentas,
        Spacer(1, 0.3 * cm),
        Paragraph(
            '<b>Por que PyInstaller?</b> É a única opção com suporte testado e hooks automáticos '
            'para as três bibliotecas do projeto (pandas, reportlab, tkinter) no Python 3.14, sem '
            'exigir configuração manual adicional.',
            s['corpo'],
        ),

        # ── PASSO 2 ──
        Paragraph('Passo 2 — Instalação do PyInstaller', s['subsecao']),
        Paragraph(
            'O PyInstaller não vem com o Python — precisa ser instalado uma única vez na '
            'máquina de desenvolvimento:',
            s['corpo'],
        ),
        *bloco_codigo(['pip install pyinstaller'], s),
        Paragraph(
            '<b>Nota:</b> O PyInstaller é uma ferramenta de <i>desenvolvimento</i>. Ela '
            'não precisa estar instalada na máquina do usuário final — apenas na máquina '
            'onde o .exe é gerado.',
            s['corpo'],
        ),

        # ── PASSO 3 ──
        Paragraph('Passo 3 — Comando de geração do executável', s['subsecao']),
        Paragraph('O comando completo utilizado foi:', s['corpo']),
        *bloco_codigo([
            'python -m PyInstaller --onefile --windowed --name "GeradorRecibos" gerar_recibos.py',
        ], s),
        Paragraph('Detalhamento de cada flag:', s['corpo']),
        tabela_flags,
        Spacer(1, 0.3 * cm),
        Paragraph(
            '<b>Por que python -m PyInstaller e não pyinstaller direto?</b> O pip instalou '
            'o executável pyinstaller.exe em uma pasta fora do PATH do sistema. Usar '
            '<i>python -m PyInstaller</i> garante que o módulo correto seja chamado, '
            'independente de configuração de PATH.',
            s['corpo'],
        ),

        # ── PASSO 4 ──
        Paragraph('Passo 4 — O que acontece internamente durante o build', s['subsecao']),
        Paragraph(
            'O PyInstaller executa as seguintes etapas automaticamente:',
            s['corpo'],
        ),
        Paragraph(
            '<b>1. Análise de dependências:</b> Percorre todas as importações do script '
            '(<i>import pandas</i>, <i>import reportlab</i>, etc.) e mapeia recursivamente '
            'todos os módulos necessários, incluindo dependências de dependências.',
            s['topico'],
        ),
        Paragraph(
            '<b>2. Hooks automáticos:</b> Para bibliotecas complexas como pandas e reportlab, '
            'o PyInstaller aplica "hooks" — scripts que incluem arquivos de dados extras '
            'que a análise estática não detectaria (ex: fontes do reportlab, dados de fuso '
            'horário do pandas).',
            s['topico'],
        ),
        Paragraph(
            '<b>3. Empacotamento em PYZ:</b> Os módulos Python são compilados para bytecode '
            '(.pyc) e comprimidos em um arquivo <i>PYZ-00.pyz</i>.',
            s['topico'],
        ),
        Paragraph(
            '<b>4. Bootloader:</b> Um pequeno programa C (runw.exe para modo windowed) é '
            'usado como ponto de entrada. Ao ser executado, ele extrai os arquivos para uma '
            'pasta temporária e inicializa o interpretador Python embutido.',
            s['topico'],
        ),
        Paragraph(
            '<b>5. Empacotamento final:</b> Bootloader + PYZ + DLLs + dados são costurados '
            'em um único <i>GeradorRecibos.exe</i> (~40 MB).',
            s['topico'],
        ),

        # ── PASSO 5 ──
        Paragraph('Passo 5 — Estrutura de arquivos gerada', s['subsecao']),
        *bloco_codigo([
            'recibo_de_pagamento/',
            '├── build/                  ← arquivos intermediários do build (ignorar)',
            '│   └── GeradorRecibos/',
            '│       ├── Analysis-00.toc',
            '│       ├── PYZ-00.pyz',
            '│       ├── warn-GeradorRecibos.txt',
            '│       └── ...',
            '├── dist/                   ← pasta de saída',
            '│   └── GeradorRecibos.exe  ← ✔ este é o arquivo para distribuir',
            '└── GeradorRecibos.spec     ← receita de build gerada automaticamente',
        ], s),
        Paragraph(
            '<b>build/:</b> Arquivos temporários usados durante o processo. Podem ser '
            'apagados após o build sem problema — o PyInstaller os recria na próxima execução.',
            s['corpo'],
        ),
        Paragraph(
            '<b>dist/GeradorRecibos.exe:</b> O único arquivo necessário para distribuir. '
            'Pode ser copiado para qualquer computador Windows e executado com dois cliques.',
            s['corpo'],
        ),
        Paragraph(
            '<b>GeradorRecibos.spec:</b> Arquivo de configuração gerado automaticamente. '
            'Registra todas as opções usadas. Pode ser editado para builds avançados '
            '(adicionar ícone, incluir arquivos extras, etc.).',
            s['corpo'],
        ),

        # ── PASSO 6 ──
        Paragraph('Passo 6 — .gitignore: o que versionar e o que ignorar', s['subsecao']),
        Paragraph(
            'Os arquivos de build <b>não devem</b> ser versionados no git — são pesados, '
            'gerados automaticamente e específicos da máquina de desenvolvimento. '
            'O arquivo <i>.gitignore</i> foi atualizado com:',
            s['corpo'],
        ),
        *bloco_codigo([
            '# PyInstaller',
            'build/',
            'dist/',
            '*.spec',
        ], s),
        Paragraph(
            '<b>Por que ignorar o .spec?</b> O .spec é gerado a partir do comando com as '
            'flags e pode ser recriado a qualquer momento. Versionar arquivos gerados '
            'automaticamente cria ruído desnecessário no histórico git.',
            s['corpo'],
        ),

        # ── PASSO 7 ──
        Paragraph('Passo 7 — Como regenerar o executável após mudanças', s['subsecao']),
        Paragraph(
            'Sempre que o código do <i>gerar_recibos.py</i> for alterado, o executável '
            'precisa ser gerado novamente. O comando é sempre o mesmo:',
            s['corpo'],
        ),
        *bloco_codigo([
            'python -m PyInstaller --onefile --windowed --name "GeradorRecibos" gerar_recibos.py',
        ], s),
        Paragraph(
            'O PyInstaller detecta os arquivos de build anteriores e reutiliza o que '
            'não mudou, tornando builds subsequentes mais rápidos que o primeiro.',
            s['corpo'],
        ),

        # ── PASSO 8 ──
        Paragraph('Passo 8 — Como distribuir para outro computador', s['subsecao']),
        Paragraph(
            'O arquivo <b>dist\\GeradorRecibos.exe</b> é completamente autossuficiente. '
            'Para instalar em outro computador Windows:',
            s['corpo'],
        ),
        Paragraph(
            '1. Copie o arquivo <i>GeradorRecibos.exe</i> para o computador de destino '
            '(pen drive, e-mail, OneDrive, etc.).',
            s['topico'],
        ),
        Paragraph(
            '2. Coloque-o em qualquer pasta (ex: Área de Trabalho ou Documentos).',
            s['topico'],
        ),
        Paragraph(
            '3. Dê dois cliques para abrir. Na primeira execução, o Windows Defender pode '
            'exibir um aviso de "aplicativo desconhecido" — clique em <i>Mais informações</i> '
            'e depois em <i>Executar assim mesmo</i>. Isso acontece porque o executável não '
            'tem assinatura digital (certificado pago).',
            s['topico'],
        ),
        Paragraph(
            '4. Não é necessário instalar Python, pip ou qualquer dependência.',
            s['topico'],
        ),

        # ── AVISO IMPORTANTE ──
        Spacer(1, 0.3 * cm),
        Paragraph(
            '<b>Importante — compatibilidade de sistema operacional:</b> O executável gerado '
            'no Windows 64-bit roda <b>somente</b> em Windows 64-bit. Para rodar em Windows '
            '32-bit seria necessário gerar o .exe em uma máquina 32-bit. Para macOS ou Linux '
            'seria necessário gerar o executável nesses sistemas operacionais separadamente.',
            s['corpo'],
        ),

        # ── PASSO 9 ──
        Paragraph('Passo 9 — Ícone na barra de tarefas do Windows', s['subsecao']),
        Paragraph(
            'Após gerar o executável com <i>--icon</i>, o ícone aparece corretamente no '
            'Explorer e na Área de Trabalho, mas a barra de tarefas continuava exibindo '
            'o ícone genérico da pena do Tk. Isso acontece porque existem <b>dois ícones '
            'independentes</b> no Windows:',
            s['corpo'],
        ),
        Paragraph(
            '• <b>Ícone do arquivo .exe</b> — definido pela flag <i>--icon</i> do PyInstaller. '
            'Aparece no Explorer, Área de Trabalho e Iniciar. É estático, embutido no '
            'cabeçalho PE do executável.',
            s['topico'],
        ),
        Paragraph(
            '• <b>Ícone da janela em execução</b> — controlado pelo tkinter via '
            '<i>root.iconbitmap()</i>. É o que aparece na barra de tarefas e no canto '
            'superior esquerdo da janela enquanto o programa está aberto.',
            s['topico'],
        ),
        Paragraph(
            'Para corrigir, foram necessárias três mudanças:',
            s['corpo'],
        ),
        Paragraph('<b>1. Função resource_path()</b>', s['subsecao']),
        Paragraph(
            'Dentro de um executável gerado pelo PyInstaller (<i>--onefile</i>), os arquivos '
            'empacotados não ficam na pasta do .exe — são extraídos para uma pasta temporária '
            'em <i>C:\\Users\\...\\AppData\\Local\\Temp\\_MEIxxxxxx\\</i> cujo caminho fica em '
            '<i>sys._MEIPASS</i>. Ao rodar como script normal, esse atributo não existe.',
            s['corpo'],
        ),
        *bloco_codigo([
            'def resource_path(filename: str) -> str:',
            '    """Caminho correto para recursos no .exe (PyInstaller) ou no script."""',
            '    base = getattr(sys, "_MEIPASS",',
            '                   os.path.dirname(os.path.abspath(__file__)))',
            '    return os.path.join(base, filename)',
        ], s),
        Paragraph(
            '<b>getattr(sys, "_MEIPASS", fallback):</b> Se o atributo não existir '
            '(execução como script), retorna o fallback — a pasta onde o .py está. '
            'Isso torna o código compatível com os dois modos sem if/else explícito.',
            s['corpo'],
        ),
        Paragraph('<b>2. root.iconbitmap() na inicialização da janela</b>', s['subsecao']),
        *bloco_codigo([
            'def iniciar_gui():',
            '    root = tk.Tk()',
            '    root.title("Gerador de Recibos")',
            '    root.resizable(False, False)',
            '',
            '    ico = resource_path("recibo.ico")',
            '    if os.path.exists(ico):',
            '        root.iconbitmap(ico)',
        ], s),
        Paragraph(
            '<b>Por que iconbitmap e não iconphoto?</b> No Windows, <i>iconbitmap()</i> com '
            'arquivo <i>.ico</i> é a forma mais confiável de definir o ícone da barra de '
            'tarefas. O <i>iconphoto()</i> aceita PNG mas pode apresentar inconsistências '
            'de renderização dependendo da versão do Windows e do Tcl/Tk.',
            s['corpo'],
        ),
        Paragraph(
            '<b>Por que o if os.path.exists()?</b> Garante que, se o arquivo não for '
            'encontrado por qualquer motivo, o programa continua funcionando normalmente '
            'em vez de lançar uma exceção na inicialização.',
            s['corpo'],
        ),
        Paragraph('<b>3. --add-data no comando do PyInstaller</b>', s['subsecao']),
        Paragraph(
            'Sem esta flag, o <i>recibo.ico</i> não seria empacotado dentro do .exe e '
            '<i>resource_path()</i> nunca encontraria o arquivo em <i>sys._MEIPASS</i>:',
            s['corpo'],
        ),
        *bloco_codigo([
            'python -m PyInstaller --onefile --windowed \\',
            '    --name "GeradorRecibos" \\',
            '    --icon "recibo.ico" \\',
            '    --add-data "recibo.ico;." \\',
            '    gerar_recibos.py',
        ], s),
        Paragraph(
            'O formato <i>"recibo.ico;."</i> significa: inclua o arquivo <i>recibo.ico</i> '
            'e extraia-o na raiz (<i>.</i>) da pasta temporária <i>_MEIPASS</i>. '
            'O separador é <b>ponto-e-vírgula</b> no Windows (no Linux/macOS seria dois-pontos).',
            s['corpo'],
        ),
        Paragraph(
            '<b>Nota sobre cache de ícones do Windows:</b> Se após atualizar o .exe o ícone '
            'antigo ainda aparecer na barra de tarefas, é o cache de ícones do Windows. '
            'Solução: clique com o botão direito no ícone antigo na barra → '
            '<i>Desafixar da barra de tarefas</i> → abra o novo .exe novamente.',
            s['corpo'],
        ),
    ]


def secao_futuro(s: dict) -> list:
    return [
        *divisor(),
        Paragraph('9. PRÓXIMOS PASSOS PREVISTOS', s['secao']),
        Paragraph(
            'O programa já conta com interface gráfica, ícone personalizado e executável Windows. '
            'As melhorias abaixo são sugestões para evoluções futuras.',
            s['corpo'],
        ),
        Paragraph(
            '• <b>Pré-visualização:</b> mostrar os recibos na tela antes de salvar o PDF.',
            s['topico'],
        ),
        Paragraph(
            '• <b>Personalização:</b> permitir alterar dados da empresa (CNPJ, endereço) '
            'direto pela interface sem editar o código.',
            s['topico'],
        ),
        Paragraph(
            '• <b>Histórico:</b> salvar um registro de quais recibos foram gerados e quando.',
            s['topico'],
        ),
    ]


# ---------------------------------------------------------------------------
# Geração do PDF
# ---------------------------------------------------------------------------

def gerar_documentacao():
    arquivo_pdf = (
        r'C:\Users\cleit\OneDrive\Documentos\recibo_de_pagamento'
        r'\documentacao_gerar_recibos.pdf'
    )

    doc = SimpleDocTemplate(
        arquivo_pdf,
        pagesize=A4,
        leftMargin=2.2 * cm,
        rightMargin=2.2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
    )

    s = criar_estilos()

    story = []
    story += secao_capa(s)
    story += secao_objetivo(s)
    story += secao_bibliotecas(s)
    story += secao_estrutura(s)
    story += secao_layout(s)
    story += secao_uso(s)
    story += secao_estrutura_planilha(s)
    story += secao_gui(s)
    story += secao_executavel(s)
    story += secao_futuro(s)
    story += divisor()
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph('Fim da documentação.', s['rodape']))

    doc.build(story)
    print(f'Documentação gerada: {arquivo_pdf}')


if __name__ == '__main__':
    gerar_documentacao()
