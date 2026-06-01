# MG Contécnica — Relatório de Pendências

Painel web gerado automaticamente a partir dos relatórios Excel exportados do sistema.  
Cada vez que um novo `.xlsx` é adicionado à pasta `data/`, o GitHub Actions gera um novo `index.html` e publica no GitHub Pages.

---

## Como funciona

```
Você sobe um novo Excel em data/
        ↓
GitHub Actions detecta a mudança
        ↓
Script Python lê o Excel, compara com a base da semana
        ↓
Gera index.html atualizado
        ↓
GitHub Pages publica → link fica atualizado
```

**Lógica de comparação:**
- **Base** = primeiro arquivo da semana (segunda-feira)
- **Atual** = arquivo mais recente
- A comparação mostra: tarefas em andamento, baixadas e adicionadas desde a base

---

## Configuração inicial (uma única vez)

### 1. Criar o repositório no GitHub

1. Acesse [github.com](https://github.com) e clique em **New repository**
2. Dê um nome (ex: `relatorio-pendencias`)
3. Deixe como **Private** se os dados são internos ⚠️
4. Clique em **Create repository**

### 2. Subir esses arquivos

No terminal, dentro desta pasta:

```bash
git init
git add .
git commit -m "setup inicial"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/relatorio-pendencias.git
git push -u origin main
```

### 3. Ativar o GitHub Pages

1. No repositório, vá em **Settings → Pages**
2. Em **Source**, selecione **GitHub Actions**
3. Salve

### 4. Ativar permissões da Action

1. Vá em **Settings → Actions → General**
2. Em **Workflow permissions**, selecione **Read and write permissions**
3. Salve

Pronto! O link do painel será:
```
https://SEU_USUARIO.github.io/relatorio-pendencias/
```

---

## Como atualizar (uso diário)

### Opção A — Pelo GitHub diretamente (sem instalar nada)

1. Acesse o repositório no GitHub
2. Clique na pasta `data/`
3. Clique em **Add file → Upload files**
4. Arraste o novo `.xlsx`
5. Clique em **Commit changes**

O relatório atualiza sozinho em ~1 minuto.

### Opção B — Pelo terminal (Git instalado)

```bash
# Copie o novo Excel para a pasta data/
cp ~/Downloads/Todas_-_Relatorio_*.xlsx data/

# Suba para o GitHub
git add data/
git commit -m "relatório $(date +%d/%m/%Y)"
git push
```

---

## Nomenclatura dos arquivos

O script detecta a data/hora automaticamente pelo nome do arquivo.  
Formatos suportados:

| Exemplo de nome | Interpretado como |
|---|---|
| `Todas_-_Relatorio_-_01-06-26_1009.xlsx` | 01/06/2026 10:09 |
| `Pendencias_2026-06-01_1430.xlsx` | 01/06/2026 14:30 |
| `Relatorio_01-06-2026.xlsx` | 01/06/2026 00:00 |

Se o nome não tiver data, usa a data de modificação do arquivo.

---

## Estrutura do repositório

```
├── data/                          ← coloque os Excel aqui
│   └── Todas_-_Relatorio_*.xlsx
├── scripts/
│   ├── gerar_relatorio.py         ← gerador principal do HTML
│   └── encontrar_e_gerar.py       ← orquestrador (escolhe base vs atual)
├── .github/
│   └── workflows/
│       └── gerar_relatorio.yml    ← Action que roda tudo automaticamente
├── index.html                     ← relatório gerado (não editar manualmente)
└── README.md
```

---

## Executar localmente (opcional)

```bash
pip install pandas openpyxl

# Visualização única (um arquivo)
python scripts/gerar_relatorio.py data/arquivo.xlsx

# Comparação (base vs atual)
python scripts/gerar_relatorio.py data/base.xlsx data/atual.xlsx

# Saída em arquivo específico
python scripts/gerar_relatorio.py data/base.xlsx data/atual.xlsx -o relatorio.html
```
