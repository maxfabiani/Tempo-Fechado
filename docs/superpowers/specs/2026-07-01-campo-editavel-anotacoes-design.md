# Campo de Anotação Editável nas 6 Primeiras Guias

**Data:** 2026-07-01
**Projeto:** Tempo Fechado v8.21.38
**Status:** Aprovado

## 1. Objetivo

Permitir que o operador insira e edite anotações textuais livremente em qualquer
registro das 6 primeiras guias da sidebar (Espelho de Ponto, Inconsistências,
Sem Marcação, Ponto Aberto, Violações de Jornada, Violações Inter Jornada).

Cada anotação fica vinculada ao registro específico e persiste entre recargas e
reimportações do Excel de origem.

## 2. Armazenamento

**Arquivo:** `%USERPROFILE%/ponto_pdfs/config/anotacoes.db` (SQLite)

**Schema:**

```sql
CREATE TABLE IF NOT EXISTS anotacoes (
    id              TEXT PRIMARY KEY,
    nome            TEXT NOT NULL,
    data_ref        TEXT NOT NULL DEFAULT '',
    data_anterior   TEXT DEFAULT '',
    data_retorno    TEXT DEFAULT '',
    tipo_violacao   TEXT DEFAULT '',
    anotacao        TEXT NOT NULL DEFAULT '',
    usuario_criacao TEXT NOT NULL DEFAULT '',
    criado_em       TEXT NOT NULL DEFAULT '',
    atualizado_em   TEXT NOT NULL DEFAULT ''
);
```

## 3. IDs Determinísticos

Cada linha recebe um campo `_anotacao_id` na resposta JSON da API. O ID é um
hash MD5 (primeiros 16 caracteres hex) dos campos que tornam o registro único:

| Guia | Composição do ID |
|------|-----------------|
| 1-4 (consolidado, inconsistências, sem marcação, ponto aberto) | `md5(nome \| "\|" \| data)` |
| 5 - Violações 12x36 | `md5(nome \| "\|" \| data_referencia \| "\|12x36")` |
| 5 - Excesso semanal | `md5(nome \| "\|" \| periodo \| "\|excesso_semanal")` |
| 6 - Inter Jornada | `md5(nome \| "\|" \| data_anterior \| "\|" \| data_retorno \| "\|interjornada")` |

## 4. API

### `GET /api/anotacoes`

- Query: `?ids=id1,id2,id3` (opcional; sem ids retorna todas)
- Resposta: `{ "anotacoes": { "id1": "texto", "id2": "texto" } }`
- Usada pelo frontend ao carregar cada guia para buscar anotações das linhas visíveis

### `POST /api/anotacoes`

- Body:
```json
{
  "id": "a1b2c3d4e5f6...",
  "nome": "Fulano",
  "data_ref": "15/03/2024",
  "data_anterior": "",
  "data_retorno": "",
  "tipo_violacao": "",
  "anotacao": "Horas extras em sobreaviso"
}
```
- Executa `INSERT OR REPLACE` na tabela
- Resposta: `{ "ok": true }`

## 5. Frontend

### Ícone de balão

- Balão de diálogo na última célula de cada linha da tabela
- `💬` (contorno) para registros sem anotação
- `💬` (azul/preenchido) para registros com anotação
- Pequeno, não polui visualmente a linha

### Editor inline

1. Usuário clica no balão → uma linha extra (`<tr>`) é inserida abaixo da linha atual
2. A linha extra tem `colspan="N"` ocupando toda a largura da tabela
3. Conteúdo: textarea + botões Salvar / Cancelar
4. Ctrl+Enter salva, Escape cancela
5. Após salvar com sucesso (`POST /api/anotacoes`):
   - Editor é removido
   - Balão fica preenchido (azul)
6. Clicar fora do editor sem salvar também o recolhe

### Carregamento inicial

1. Ao abrir qualquer das 6 guias, o frontend coleta `_anotacao_id` de todas as
   linhas visíveis no `#tbody`
2. Faz `GET /api/anotacoes?ids=id1,id2,...`
3. Atualiza cada balão (preenchido/vazado) conforme o retorno

## 6. Arquivos modificados

| Arquivo | Mudanças |
|---------|----------|
| `robo_ponto_web.py` | 2 novas rotas (`GET/POST /api/anotacoes`) + injeção de `_anotacao_id` nos 5 endpoints das 6 guias |
| `static/app.js` | Ícone de balão, editor inline, chamadas fetch para API |
| `static/style.css` | Estilos do balão, linha de editor, textarea, botões |

## 7. Não escopo

- Exportar/importar anotações
- Histórico de versões da anotação
- Anotações em outras guias fora das 6 primeiras
- Notificações ou alertas baseados em anotações
