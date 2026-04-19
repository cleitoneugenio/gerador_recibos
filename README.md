# Gerador de Recibos de Prestação de Serviço

Programa que lê uma planilha Excel semanal e gera automaticamente um PDF com todos os recibos de prestação de serviço prontos para impressão e assinatura.

---

## Funcionalidades

- Lê planilha `.xlsx` com nomes e valores dos funcionários
- Suporte a planilhas com **múltiplas abas** — selecione a semana desejada na interface
- **Filtro por funcionário** com checkboxes — escolha exatamente quem incluir no PDF
- **Painel de prévia** com navegação ← → entre os recibos selecionados
- Gera um PDF com um recibo por funcionário, separados por linha tracejada para recorte
- Valor por extenso em português gerado automaticamente (ex: *R$ 150,00 (cento e cinquenta reais)*)
- Data preenchida automaticamente com o dia da geração
- Funcionários com total R$ 0,00 são ignorados automaticamente
- Interface gráfica redimensionável para uso sem terminal
- Executável Windows `.exe` — não precisa instalar Python
- **Log persistente** em `%APPDATA%\GeradorRecibos\gerador.log` para diagnóstico

---

## Como usar

### Opção 1 — Interface gráfica (recomendado)

Execute o `GeradorRecibos.exe` com dois cliques. A janela abre automaticamente.

1. Clique em **⋯** ao lado de **PLANILHA** e selecione o arquivo `.xlsx`
2. No campo **ABA**, selecione a semana desejada — preenchido automaticamente com a primeira aba
3. O nome do PDF de saída é preenchido automaticamente
4. No painel **FUNCIONÁRIOS**, marque ou desmarque quem deve entrar no PDF (todos marcados por padrão)
5. Use **Selecionar todos** / **Desmarcar todos** para seleção rápida
6. O painel **PRÉVIA** à direita mostra o recibo do funcionário selecionado — navegue com ← →
7. Clique em **Gerar Recibos**
8. Ao concluir, o programa pergunta se deseja abrir o PDF

> A janela é redimensionável. A prévia se atualiza automaticamente ao redimensionar.

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
pip install -r requirements.txt
```

Ou manualmente:

```bash
pip install customtkinter==5.2.2 num2words==0.5.14 pandas==3.0.2 Pillow==12.2.0 reportlab==4.4.10
```

> `openpyxl` é instalado automaticamente como dependência do `pandas`.

### 2. Executar

```bash
python gerar_recibos.py
```

---

## Testes

A suite de testes cobre todas as funções críticas do programa.

```bash
# Instalar pytest (apenas uma vez)
pip install pytest openpyxl

# Rodar todos os testes
python -m pytest test_gerar_recibos.py -v
```

**Cobertura atual: 29 testes, todos passando.**

| Classe de teste | O que testa |
|-----------------|-------------|
| `TestFmtValor` | Formatação monetária brasileira |
| `TestValorPorExtenso` | Conversão de valores para texto |
| `TestFormatarData` | Formatação de datas em português |
| `TestLerFuncionarios` | Leitura da planilha xlsx, filtros, erros |
| `TestGerarPdfDeLista` | Geração de PDF, múltiplos funcionários, XML injection |
| `TestGerarPreviewPil` | Renderização da prévia, thread safety, caracteres especiais |
| `TestResourcePath` | Localização de recursos dentro e fora do PyInstaller |

---

## Log do sistema

O programa grava um log rotativo em:

```
%APPDATA%\GeradorRecibos\gerador.log
```

- Tamanho máximo: 1 MB por arquivo, 3 arquivos de histórico mantidos
- Nível: `INFO` e acima para eventos de negócio; `WARNING`/`ERROR` para falhas
- Quando executado em terminal (`isatty`): exibe `INFO` também no console

Útil para diagnóstico quando o programa não se comporta como esperado.

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
├── test_gerar_recibos.py     # suite de testes (pytest)
├── requirements.txt          # dependências com versões fixas
├── recibo.png                # ícone do programa (ignorado pelo git)
├── .gitignore
└── README.md
```

---

## Bibliotecas utilizadas

| Biblioteca | Versão | Finalidade |
|------------|--------|-----------|
| `customtkinter` | 5.2.2 | Interface gráfica moderna (dark mode) |
| `pandas` | 3.0.2 | Leitura e processamento da planilha `.xlsx` |
| `openpyxl` | — | Motor de leitura de arquivos Excel modernos (dep. do pandas) |
| `Pillow` | 12.2.0 | Renderização da prévia do recibo como imagem |
| `reportlab` | 4.4.10 | Geração do PDF com layout completo |
| `num2words` | 0.5.14 | Conversão de valores numéricos para texto em português |

---

## Decisões técnicas e bugs conhecidos resolvidos

### Prévia em thread separada com token de cancelamento

A renderização da prévia (`gerar_preview_pil`) roda em `threading.Thread` para não travar a UI. Um mecanismo de token (`_preview_token`) garante que apenas o resultado mais recente seja aplicado — renders anteriores descartados se o usuário mudar de aba durante o processamento.

### Bug: CTkLabel não limpa imagem no widget tkinter ao receber `image=None`

`CTkLabel._update_image()` não chama `self._label.configure(image='')` quando `self._image` passa a ser `None`. O widget tkinter interno mantém a referência ao nome do `PhotoImage` (`"pyimage2"`, etc.). Quando o `CTkImage` Python é coletado pelo GC, o `PhotoImage` é deletado do Tcl — mas o label ainda referencia o nome. Qualquer `configure(text=...)` subsequente lança `TclError: image "pyimage2" doesn't exist`.

**Fix aplicado:** helper `_clear_preview_image()` limpa o widget tkinter diretamente via `prev_label._label.configure(image='')` antes de liberar a referência Python, contornando o bug interno do customtkinter.

### Bug: trace de BooleanVar causava TclError ao trocar de aba

Ao trocar de aba, `CTkCheckBox.destroy()` tentava remover seu trace interno do `BooleanVar`. Se outro trace (o nosso, `_atualizar_summary`) fosse removido erroneamente antes, o `destroy()` falhava com `TclError`. A falha silenciosa deixava `list_frame[0]` apontando para um frame destruído, e a próxima carga de funcionários tentava criar widgets filhos nele — cascata de erros.

**Fix aplicado:** `checks_vars` armazena 3-tuplas `(var, nome, trace_id)`. Em `_clear_scroll()`, apenas o trace da aplicação é removido pelo ID exato (`var.trace_remove('write', trace_id)`), preservando o trace interno do `CTkCheckBox`. O `_make_list_frame()` tem `try/except` no `destroy()` para garantir que o novo frame seja sempre criado mesmo se a destruição do anterior falhar.

---

## Empresa

Desenvolvido para uso interno da **GF MUNIZ ARTEFACTOS DE CERAMICA EIRELI**
CNPJ: 12.509.424/0001-65 — Bela Cruz - CE
