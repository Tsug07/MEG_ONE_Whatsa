<div align="center">

<!-- LOGO: substitua pelo caminho do seu logo -->
<!-- <img src="logo.png" alt="M.E.G ONE Logo" width="160"> -->

# M.E.G ONE — Main Excel Generator ONE

**Gerador automatizado de planilhas Excel para envio de mensagens via WhatsApp**

![Version](https://img.shields.io/badge/version-2.0-blue?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-Proprietary-red?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Production-brightgreen?style=for-the-badge)

---

*Ferramenta interna de RPA que cruza dados de PDFs e planilhas Excel com a base de contatos, gerando o Excel final pronto para automação de envio via WhatsApp.*

</div>

---

## Funcionalidades

- Interface gráfica moderna com **CustomTkinter** (tema escuro)
- **8 modelos** de processamento para diferentes fluxos de trabalho
- Cruzamento inteligente de dados por **código**, **nome** ou **similaridade** (>= 80%)
- Extração automática de dados de **PDFs** (pdfplumber)
- Formatação automática de **CNPJ/CPF**
- Log em tempo real com barra de progresso
- Compatível com **PyInstaller** para distribuição como `.exe`

---

## Modelos Disponíveis

| Modelo | Entrada | Saída |
|:--|:--|:--|
| **ONE** | Pasta de PDFs + Excel Contatos | `Codigo` · `Nome` · `Numero` · `Caminho` |
| **Cobranca** | PDF de cobrança + Excel Contatos | `Codigo` · `Nome` · `Numero` · `Valor da Parcela` · `Data de Vencimento` · `Carta de Aviso` |
| **ComuniCertificado** | Excel Base + Excel Contatos | `Codigo` · `Nome` · `Numero` · `CNPJ` · `Vencimento` · `Carta de Aviso` |
| **Contato** | Excel Base + Excel Contatos | `Codigo` · `Nome` · `Contato` · `Grupo` · `Telefone` · `CNPJ` |
| **ALL** | Excel Origem + Excel Contatos | `Codigo` · `Empresa` · `Contato Onvio` · `Grupo Onvio` · `CNPJ` · `Telefone` |
| **ALL_info** | Excel Origem + Excel Contatos | `Codigo` · `Nome` · `Numero` · `CNPJ` · `Competencia` |
| **DomBot_GMS** | Excel Base | `Nº` · `EMPRESAS` · `Periodo` · `Salvar Como` · `Competencia` · `Caminho` |
| **DomBot_Econsig** | PDF de empréstimos | `Nº` · `EMPRESAS` · `Data Inicial` · `Data Final` · `Salvar Como` |

---

## Formato do Excel de Contatos

O arquivo Excel de contatos deve seguir obrigatoriamente este formato:

| Coluna A | Coluna B | Coluna C | Coluna D | Coluna E | Coluna F |
|:--------:|:--------:|:--------:|:--------:|:--------:|:--------:|
| Codigo | Empresa | Contato Onvio | Grupo Onvio | CNPJ | Telefone |

> O campo **Telefone** (coluna F) é o número utilizado para envio via WhatsApp e é mapeado como `Numero` nos modelos ONE, Cobranca, ComuniCertificado e ALL_info.

---

## Instalação

### Pre-requisitos

```
Python >= 3.10
```

### Dependências

```bash
pip install pandas openpyxl pdfplumber customtkinter Pillow
```

### Execução

```bash
python M.E.G_ONE_Whatsa.py
```

### Build (executável)

```bash
pyinstaller --onefile --windowed --add-data "logo.png;." --add-data "logoIcon.ico;." --icon=logoIcon.ico M.E.G_ONE_Whatsa.py
```

---

## Estrutura do Projeto

```
MEG_ONE_Whatsa/
├── M.E.G_ONE_Whatsa.py   # Código principal (GUI + processadores)
├── logo.png              # Logo exibido na interface
├── logoIcon.ico          # Ícone da janela
└── README.md
```

---

## Como Usar

1. Abra a aplicação
2. Selecione o **Modelo** desejado no dropdown
3. Preencha os campos de entrada (PDF, Excel Base, Excel Contatos)
4. Defina o caminho do **Excel de saída**
5. Clique em **Processar Relatórios**
6. Acompanhe o progresso no log em tempo real

---

## Carta de Aviso (Cobranca / ComuniCertificado)

Os modelos de cobrança e comunicado classificam automaticamente a urgência:

**Cobranca** — dias após vencimento:

| Carta | Condição |
|:-----:|:---------|
| 1 | Até 6 dias |
| 2 | 7 – 14 dias |
| 3 | 15 – 19 dias |
| 4 | 20 – 24 dias |
| 5 | 25 – 30 dias |
| 6 | Mais de 30 dias |

**ComuniCertificado** — dias até vencimento:

| Carta | Condição |
|:-----:|:---------|
| 1 | Mais de 5 dias restantes |
| 2 | 1 – 5 dias restantes |
| 3 | Vence hoje |
| 4 | Vencido |

---

<div align="center">

Desenvolvido por **Hugo** · 2025

</div>
