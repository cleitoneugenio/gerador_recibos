#!/usr/bin/env python3
"""
Gerador de Recibos de Prestação de Serviço
Uso: python gerar_recibos.py <arquivo.xlsx> [saida.pdf]
"""

import os
import sys
from datetime import datetime

import pandas as pd
from num2words import num2words
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
)


def valor_por_extenso(valor: float) -> str:
    reais = int(valor)
    centavos = round((valor - reais) * 100)
    extenso = num2words(reais, lang='pt_BR') + ' reais'
    if centavos > 0:
        extenso += ' e ' + num2words(centavos, lang='pt_BR') + ' centavos'
    valor_fmt = f'R$ {valor:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
    return f'{valor_fmt} ({extenso})'


def formatar_data(data: datetime) -> str:
    meses = [
        'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
        'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro',
    ]
    return f'{data.day} de {meses[data.month - 1]} de {data.year}'


def ler_funcionarios(arquivo: str) -> list:
    df = pd.read_excel(arquivo, header=0)

    funcionarios = []
    for _, row in df.iterrows():
        # Col 0 = número do funcionário; deve ser inteiro positivo
        try:
            int(row.iloc[0])
        except (ValueError, TypeError):
            continue  # pula cabeçalho e linha TOTAL

        nome = str(row.iloc[1]).strip()
        if not nome or nome.lower() == 'nan':
            continue

        try:
            total = float(row.iloc[10])
        except (ValueError, TypeError):
            continue

        if total <= 0:
            continue

        funcionarios.append({'nome': nome, 'total': total})

    return funcionarios


def montar_recibo(nome: str, total: float, data_str: str, styles: dict) -> list:
    valor_str = valor_por_extenso(total)

    corpo_texto = (
        'Recebi da empresa <b>GF MUNIZ ARTEFACTOS DE CERAMICA EIRELI</b>, pessoa jurídica de '
        'direito privado, inscrita no CNPJ sob o n 12.509.424/0001-65, com sede na cidade de '
        'Bela Cruz - CE ROD CE 179 - KM 12 - S/N - Zona Rural, a quantia de '
        f'<b>{valor_str}</b>. '
        f'Referente a serviço de diárias no dia {data_str}. '
        'dando-lhe por este recibo a devida quitação.'
    )

    return [
        Paragraph(
            '<u><b>RECIBO DE PRESTAÇÃO DE SERVIÇO</b></u>',
            styles['titulo'],
        ),
        Spacer(1, 0.5 * cm),
        Paragraph(corpo_texto, styles['corpo']),
        Spacer(1, 0.4 * cm),
        Paragraph(f'Bela Cruz, {data_str}.', styles['corpo']),
        Spacer(1, 1.8 * cm),
        HRFlowable(width='55%', thickness=1, color=colors.black, hAlign='CENTER'),
        Spacer(1, 0.25 * cm),
        Paragraph(nome, styles['assinatura']),
    ]


def gerar_pdf(arquivo_xlsx: str, arquivo_pdf: str) -> None:
    hoje = datetime.today()
    data_str = formatar_data(hoje)

    funcionarios = ler_funcionarios(arquivo_xlsx)
    if not funcionarios:
        print('Nenhum funcionário encontrado na planilha.')
        sys.exit(1)

    doc = SimpleDocTemplate(
        arquivo_pdf,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
    )

    styles = {
        'titulo': ParagraphStyle(
            'titulo',
            fontName='Helvetica-Bold',
            fontSize=13,
            alignment=TA_CENTER,
        ),
        'corpo': ParagraphStyle(
            'corpo',
            fontName='Helvetica',
            fontSize=10,
            alignment=TA_JUSTIFY,
            leading=15,
        ),
        'assinatura': ParagraphStyle(
            'assinatura',
            fontName='Helvetica',
            fontSize=10,
            alignment=TA_CENTER,
        ),
    }

    story = []

    for i, func in enumerate(funcionarios):
        recibo = montar_recibo(func['nome'], func['total'], data_str, styles)
        story.append(KeepTogether(recibo))

        if i < len(funcionarios) - 1:
            story.append(Spacer(1, 0.5 * cm))
            story.append(
                HRFlowable(
                    width='100%',
                    thickness=0.5,
                    color=colors.grey,
                    dash=(6, 4),
                    hAlign='CENTER',
                )
            )
            story.append(Spacer(1, 0.5 * cm))

    doc.build(story)
    print(f'PDF gerado: {arquivo_pdf} ({len(funcionarios)} recibos)')


def main():
    if len(sys.argv) < 2:
        print('Uso: python gerar_recibos.py <arquivo.xlsx> [saida.pdf]')
        sys.exit(1)

    arquivo_xlsx = sys.argv[1]
    if not os.path.exists(arquivo_xlsx):
        print(f'Arquivo não encontrado: {arquivo_xlsx}')
        sys.exit(1)

    if len(sys.argv) >= 3:
        arquivo_pdf = sys.argv[2]
    else:
        arquivo_pdf = os.path.splitext(arquivo_xlsx)[0] + '_recibos.pdf'

    gerar_pdf(arquivo_xlsx, arquivo_pdf)


if __name__ == '__main__':
    main()
