# Gerador de Recibos de Prestação de Serviço

Programa que lê uma planilha Excel semanal e gera automaticamente um PDF com todos os recibos de prestação de serviço prontos para impressão e assinatura.

---

## Funcionalidades

- Lê planilha `.xlsx` com nomes e valores dos funcionários
- Gera um PDF com um recibo por funcionário, separados por linha tracejada para recorte
- Valor por extenso em português gerado automaticamente (ex: *R$ 150,00 (cento e cinquenta reais)*)
- Data preenchida automaticamente com o dia da geração
- Funcionários com total R$ 0,00 são ignorados automaticamente
- Interface gráfica para uso sem terminal
- Executável Windows `.exe` — não precisa instalar Python

---

## Como usar

### Opção 1 — Interface gráfica (recomendado)

Execute o `GeradorRecibos.exe` com dois cliques. A janela abre automaticamente.

1. Clique em **Procurar...** e selecione a planilha `.xlsx`
2. O nome do PDF de saída é preenchido automaticamente
3. Clique em **Gerar Recibos**
4. Ao concluir, o programa pergunta se deseja abrir o PDF

### Opção 2 — Linha de comando

```bash
# PDF gerado com o mesmo nome da planilha
python gerar_recibos.py ponto_cedan_semana.xlsx

# Definindo o nome do PDF de saída
python gerar_recibos.py ponto_cedan_semana.xlsx recibos_semana1.pdf
```

---

## Estrutura esperada da planilha

| Coluna | Índice | Conteúdo |
|--------|--------|----------|
| # | 0 | Número do funcionário (inteiro) |
| FUNCIONÁRIOS | 1 | Nome completo |
| SEGUNDA | 2 | Valor da diária ou "F" (falta) |
| TERÇA | 3 | Valor da diária ou "F" (falta) |
| QUARTA | 4 | Valor da diária ou "F" (falta) |
| QUINTA | 5 | Valor da diária ou "F" (falta) |
| SEXTA | 6 | Valor da diária ou "F" (falta) |
| SÁBADO | 7 | Valor da diária ou "F" (falta) |
| DIAS | 8 | Total de dias trabalhados |
| BÔNUS | 9 | Bônus (quando aplicável) |
| **TOTAL** | **10** | **Total a receber — valor usado no recibo** |

> A planilha deve ter uma linha de título mesclado no topo e uma linha de cabeçalho. A última linha com "TOTAL" é ignorada automaticamente.

---

## Instalação (modo script Python)

### 1. Instalar dependências

```bash
pip install pandas openpyxl reportlab num2words
```

### 2. Executar

```bash
python gerar_recibos.py planilha.xlsx
```

---

## Gerar o executável Windows

Para distribuir o programa sem precisar instalar Python:

```bash
# Instalar PyInstaller (apenas uma vez)
pip install pyinstaller

# Converter o ícone (requer Pillow)
python -c "from PIL import Image; img=Image.open('recibo.png').convert('RGBA'); img.save('recibo.ico',format='ICO',sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)])"

# Gerar o executável
python -m PyInstaller --onefile --windowed --name "GeradorRecibos" --icon "recibo.ico" --add-data "recibo.ico;." gerar_recibos.py
```

O executável gerado fica em `dist/GeradorRecibos.exe`.

> O executável gerado no Windows 64-bit roda somente em Windows 64-bit.

---

## Estrutura do projeto

```
recibo_de_pagamento/
├── gerar_recibos.py          # programa principal
├── gerar_documentacao.py     # gera a documentação técnica em PDF
├── recibo.png                # ícone do programa
├── .gitignore
└── README.md
```

---

## Bibliotecas utilizadas

| Biblioteca | Finalidade |
|------------|-----------|
| `pandas` | Leitura e processamento da planilha `.xlsx` |
| `openpyxl` | Motor de leitura de arquivos Excel modernos |
| `reportlab` | Geração do PDF com layout completo |
| `num2words` | Conversão de valores numéricos para texto em português |

---

## Empresa

Desenvolvido para uso interno da **GF MUNIZ ARTEFACTOS DE CERAMICA EIRELI**
CNPJ: 12.509.424/0001-65 — Bela Cruz - CE
