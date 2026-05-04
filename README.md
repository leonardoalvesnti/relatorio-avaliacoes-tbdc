# Relatório de Avaliações TBDC - Latitude Genética

Workflow n8n para geração automática do relatório de competições e avaliações de cultivares da plataforma TBDC.

## 📊 O que faz
- Puxa dados de **competições**, **fazendas**, **variedades** e **avaliações** da API TBDC
- Consolida tudo em um relatório Excel com **29 colunas**
- Inclui mapeamento de **territórios** (197 fazendas)
- Envia automaticamente por email (segundas e quintas às 08:30)

## 📁 Estrutura
```
├── Evaluation TBDC v2.json    # Workflow n8n (importar no n8n)
├── formatar_excel.py          # Script Python para formatar Excel com identidade visual
├── assets/
│   └── logo-latitude.png      # Logo Latitude Genética (usado no email)
└── docs/
    └── walkthrough.md         # Documentação do projeto
```

## 🎨 Identidade Visual
- **Verde Latitude**: `#00BB31`
- **Cinza escuro**: `#464646`

## 🔧 Como usar

### 1. Importar workflow no n8n
1. Abra o n8n
2. Vá em **Workflows > Import from File**
3. Selecione `Evaluation TBDC v2.json`
4. Configure as credentials (TBDC Basic Auth + Bearer Token)

### 2. Formatar Excel localmente (opcional)
```bash
python formatar_excel.py relatorio-avaliacoes-27-04-2026.xlsx
```

## 📋 Colunas do Relatório
ID Campo | Tipo | Status | ID Fazenda | Nome Fazenda | Proprietário | Safra | Período | Cultura | Rep. Técnico | Protocolo | Sist. Plantio | Irrigado | ID Tratamento | ID Cultivar | Cultivar | Produtividade | GM/RMG | Área Ha | Pop. Final | Cidade | Estado | Tipo Cultivar | Território | Genética | Altura (cm) | PMG (g) | ID Avaliação | ID Campo Avaliado
