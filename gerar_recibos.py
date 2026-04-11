#!/usr/bin/env python3
"""
Gerador de Recibos de Prestação de Serviço
Uso: python gerar_recibos.py <arquivo.xlsx> [saida.pdf]
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
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


def iniciar_gui():
    root = tk.Tk()
    root.title('Gerador de Recibos')
    root.resizable(False, False)

    padding = {'padx': 10, 'pady': 5}

    # --- Planilha ---
    tk.Label(root, text='Planilha (.xlsx):').grid(row=0, column=0, sticky='w', **padding)

    var_xlsx = tk.StringVar()
    entry_xlsx = tk.Entry(root, textvariable=var_xlsx, width=50)
    entry_xlsx.grid(row=0, column=1, **padding)

    def selecionar_xlsx():
        caminho = filedialog.askopenfilename(
            title='Selecionar planilha',
            filetypes=[('Excel', '*.xlsx *.xls'), ('Todos', '*.*')],
        )
        if caminho:
            var_xlsx.set(caminho)
            if not var_pdf.get():
                var_pdf.set(os.path.splitext(caminho)[0] + '_recibos.pdf')

    tk.Button(root, text='Procurar...', command=selecionar_xlsx).grid(row=0, column=2, **padding)

    # --- PDF de saída ---
    tk.Label(root, text='PDF de saída:').grid(row=1, column=0, sticky='w', **padding)

    var_pdf = tk.StringVar()
    entry_pdf = tk.Entry(root, textvariable=var_pdf, width=50)
    entry_pdf.grid(row=1, column=1, **padding)

    def selecionar_pdf():
        caminho = filedialog.asksaveasfilename(
            title='Salvar PDF como',
            defaultextension='.pdf',
            filetypes=[('PDF', '*.pdf')],
        )
        if caminho:
            var_pdf.set(caminho)

    tk.Button(root, text='Salvar como...', command=selecionar_pdf).grid(row=1, column=2, **padding)

    # --- Barra de progresso e status ---
    progress = ttk.Progressbar(root, mode='indeterminate', length=400)
    progress.grid(row=2, column=0, columnspan=3, padx=10, pady=(10, 0))

    var_status = tk.StringVar(value='Aguardando...')
    tk.Label(root, textvariable=var_status, fg='gray').grid(
        row=3, column=0, columnspan=3, pady=(2, 5)
    )

    # --- Botão gerar ---
    btn_gerar = tk.Button(root, text='Gerar Recibos', width=20, bg='#2e7d32', fg='white',
                          font=('Helvetica', 10, 'bold'))
    btn_gerar.grid(row=4, column=0, columnspan=3, pady=(5, 15))

    def executar():
        xlsx = var_xlsx.get().strip()
        pdf = var_pdf.get().strip()

        if not xlsx:
            messagebox.showerror('Erro', 'Selecione a planilha (.xlsx).')
            return
        if not os.path.exists(xlsx):
            messagebox.showerror('Erro', f'Arquivo não encontrado:\n{xlsx}')
            return
        if not pdf:
            pdf = os.path.splitext(xlsx)[0] + '_recibos.pdf'
            var_pdf.set(pdf)

        btn_gerar.config(state='disabled')
        progress.start(10)
        var_status.set('Gerando PDF...')

        def tarefa():
            try:
                gerar_pdf(xlsx, pdf)
                root.after(0, lambda: concluido(pdf))
            except Exception as exc:
                root.after(0, lambda: erro(str(exc)))

        threading.Thread(target=tarefa, daemon=True).start()

    def concluido(pdf):
        progress.stop()
        btn_gerar.config(state='normal')
        var_status.set('PDF gerado com sucesso!')
        if messagebox.askyesno('Concluído', f'PDF gerado:\n{pdf}\n\nDeseja abrir o arquivo?'):
            os.startfile(pdf)

    def erro(msg):
        progress.stop()
        btn_gerar.config(state='normal')
        var_status.set('Erro ao gerar PDF.')
        messagebox.showerror('Erro', msg)

    btn_gerar.config(command=executar)

    root.mainloop()


def main():
    if len(sys.argv) < 2:
        iniciar_gui()
        return

    arquivo_xlsx = sys.argv[1]
    if not os.path.exists(arquivo_xlsx):
        print(f'Arquivo não encontrado: {arquivo_xlsx}')
        sys.exit(1)

    arquivo_pdf = sys.argv[2] if len(sys.argv) >= 3 else os.path.splitext(arquivo_xlsx)[0] + '_recibos.pdf'
    gerar_pdf(arquivo_xlsx, arquivo_pdf)


if __name__ == '__main__':
    main()
