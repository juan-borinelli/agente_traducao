# 🤖 Agente de Tradução de Textos Acadêmicos

Agente de IA para tradução de textos acadêmicos de até 500 páginas, com saída formatada automaticamente em ABNT (padrão) ou formato livre, desenvolvido com o framework **Agno** e modelo **Claude** da Anthropic.

## 📋 Sobre o Projeto

Textos acadêmicos longos representam um desafio específico para tradução automática: exigem fidelidade terminológica, coerência ao longo de centenas de páginas e formatação compatível com normas acadêmicas. Este projeto resolve esses problemas com uma pipeline multi-agente que divide o documento em chunks inteligentes, traduz preservando contexto entre partes e revisa o resultado para garantir fluência em português brasileiro.

## ✨ Funcionalidades

- **Tradução de documentos longos** (PDFs e DOCXs de até ~500 páginas)
- **Chunking inteligente** com preservação de parágrafos e overlap de contexto — evita traduções incoerentes entre partes
- **Limpeza automática do texto extraído** — corrige hifenização, quebras espúrias de linha e caracteres inválidos comuns em PDFs
- **Detecção automática de metadados** (idioma, título, área do documento)
- **Pipeline dual-agente**: agente de tradução + agente de revisão independentes
- **Formatação ABNT completa**: margens, fonte Times New Roman 12pt, espaçamento 1,5, recuo, sumário e página de título automáticos
- **Suporte a formato livre** (sem formatação acadêmica)
- **Retry automático** em caso de falhas na API (até 3 tentativas por chunk)
- **Interface CLI** com parâmetros configuráveis

## 🛠️ Tecnologias

| Categoria | Tecnologia |
|-----------|-----------|
| Framework de Agentes | [Agno](https://github.com/agno-agi/agno) |
| Modelo de IA | Claude (Anthropic) via `agno.models.anthropic` |
| Extração de PDF | PyMuPDF (`fitz`) |
| Geração de DOCX | `python-docx` |
| Progresso visual | `tqdm` |
| Gerenciamento de env | `python-dotenv` |
| Gerenciamento de pacotes | `uv` |

## 🏗️ Arquitetura da Pipeline

```
PDF / DOCX
    │
    ▼
Extração de texto (PyMuPDF / python-docx)
    │
    ▼
Limpeza do texto (hifenização, quebras espúrias, chars inválidos)
    │
    ▼
Chunking inteligente (respeita parágrafos, overlap como contexto)
    │
    ▼
Detecção de metadados (idioma, título, área) via Agente IA
    │
    ▼
┌─────────────────────────────────┐
│   Agente Tradutor (Claude)      │  ← traduz cada chunk com contexto anterior
│   Agente Revisor (Claude)       │  ← revisa fluência e coerência
└─────────────────────────────────┘
    │
    ▼
Geração do DOCX (ABNT ou livre)
```

## ▶️ Como Executar

**Pré-requisitos:** Python 3.12+, `uv` instalado, chave de API da Anthropic.

```bash
# 1. Clone o repositório
git clone https://github.com/juan-borinelli/agente_traducao.git
cd agente_traducao

# 2. Instale as dependências com uv
uv sync

# 3. Configure a chave de API
cp .env.example .env
# Edite .env e adicione: ANTHROPIC_API_KEY=sua_chave_aqui
```

**Uso básico:**

```bash
# Traduzir PDF com formatação ABNT (padrão)
python agent.py --input artigo.pdf --output traducao.docx

# Traduzir com formato livre (sem ABNT)
python agent.py --input artigo.pdf --output traducao.docx --format livre

# Forçar revisão completa (recomendado para documentos críticos)
python agent.py --input artigo.pdf --output traducao.docx --review

# Ajustar tamanho dos chunks (padrão: 3000 caracteres)
python agent.py --input artigo.pdf --output traducao.docx --chunk-size 2000
```

**Parâmetros disponíveis:**

| Parâmetro | Descrição | Padrão |
|-----------|-----------|--------|
| `--input` / `-i` | Arquivo de entrada (PDF ou DOCX) | obrigatório |
| `--output` / `-o` | Arquivo de saída (.docx) | obrigatório |
| `--format` / `-f` | `abnt` ou `livre` | `abnt` |
| `--chunk-size` | Tamanho dos chunks em caracteres | `3000` |
| `--review` | Força revisão completa pelo agente revisor | desativado |

## 📁 Estrutura do Projeto

```
agente_traducao/
├── agent.py          # Pipeline completa: extração, chunking, tradução, formatação
├── pyproject.toml    # Dependências gerenciadas com uv
├── uv.lock           # Lockfile de dependências
├── .gitignore        # Exclui .env e arquivos sensíveis
└── .python-version   # Versão do Python utilizada
```

## ⚙️ Variáveis de Ambiente

Crie um arquivo `.env` na raiz do projeto com:

```env
ANTHROPIC_API_KEY=sua_chave_aqui
```

> ⚠️ Nunca suba o arquivo `.env` para o repositório. Ele já está no `.gitignore`.

## 📚 Conceitos Aplicados

- Arquitetura multi-agente com responsabilidades separadas (tradução × revisão)
- Chunking com overlap de contexto para coerência entre partes longas
- Limpeza de texto extraído de PDF (hifenização, quebras de linha, caracteres de controle)
- Detecção automática de metadados via prompt estruturado com saída JSON
- Formatação programática de documentos Word com normas ABNT
- Retry exponencial para resiliência em chamadas de API
- CLI com `argparse` para uso via terminal
- Gerenciamento seguro de credenciais com `python-dotenv`

---

Desenvolvido por [Juan Borinelli](https://github.com/juan-borinelli)