# Campo Editável (Anotações) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an inline editable annotation field (balloon icon + text editor) to records in the first 6 sidebar tabs.

**Architecture:** A SQLite database (`anotacoes.db`) stores annotations keyed by a deterministic hash ID per row. New API endpoints serve/save annotations. The frontend renders a balloon icon per row, loads existing annotations on tab open, and shows an inline editor on click.

**Tech Stack:** Python 3.9+, Flask, SQLite3, Vanilla JS, CSS3

## Global Constraints

- SQLite file at `%USERPROFILE%/ponto_pdfs/config/anotacoes.db`
- Annotation ID = first 16 hex chars of `hashlib.md5(composite_key.encode()).hexdigest()`
- No new npm/pip dependencies
- Follow existing code style (no comments, tab/space style as-is)

---

### Task 1: Backend — SQLite database + GET/POST /api/anotacoes

**Files:**
- Modify: `robo_ponto_web.py` (add ~80 lines after helper functions area)

**Interfaces:**
- Produces: `_anotacoes_db_path()` → `Path`, `_init_anotacoes_db()` → `sqlite3.Connection`, `_gerar_anotacao_id(nome, data_ref, tipo_violacao, extra)` → `str`

- [ ] **Step 1: Add helper functions before the first API route** (insert after `def _df_para_resposta_leve` around line 2284)

Insert these helpers:

```python
def _init_anotacoes_db() -> sqlite3.Connection:
    db_path = Path.home() / "ponto_pdfs" / "config" / "anotacoes.db"
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("""
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
        )
    """)
    conn.commit()
    return conn


def _gerar_anotacao_id(nome: str, data_ref: str, tipo_violacao: str = "", extra: str = "") -> str:
    chave = f"{nome}|{data_ref}|{tipo_violacao}|{extra}"
    return hashlib.md5(chave.encode("utf-8")).hexdigest()[:16]
```

- [ ] **Step 2: Add GET /api/anotacoes route** (insert after helper functions, before the first API route — e.g. before `@app.route("/api/consolidado")`)

```python
@app.route("/api/anotacoes")
@login_obrigatorio
def api_get_anotacoes():
    ids_param = request.args.get("ids", "").strip()
    conn = _init_anotacoes_db()
    try:
        if ids_param:
            ids_lista = [i.strip() for i in ids_param.split(",") if i.strip()]
            if not ids_lista:
                return jsonify({"anotacoes": {}})
            placeholders = ",".join("?" for _ in ids_lista)
            rows = conn.execute(f"SELECT id, anotacao FROM anotacoes WHERE id IN ({placeholders})", ids_lista).fetchall()
        else:
            rows = conn.execute("SELECT id, anotacao FROM anotacoes").fetchall()
        return jsonify({"anotacoes": {r["id"]: r["anotacao"] for r in rows}})
    finally:
        conn.close()
```

- [ ] **Step 3: Add POST /api/anotacoes route**

```python
@app.route("/api/anotacoes", methods=["POST"])
@login_obrigatorio
def api_post_anotacoes():
    dados = request.get_json(silent=True) or {}
    id_ = str(dados.get("id", "")).strip()
    anotacao = str(dados.get("anotacao", "")).strip()
    if not id_:
        return jsonify({"erro": "ID obrigatório"}), 400

    usuario = autenticar_requisicao_atual()[0].get("usuario", "")
    agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = _init_anotacoes_db()
    try:
        existente = conn.execute("SELECT * FROM anotacoes WHERE id = ?", (id_,)).fetchone()
        if existente:
            conn.execute("UPDATE anotacoes SET anotacao = ?, usuario_criacao = ?, atualizado_em = ? WHERE id = ?",
                         (anotacao, usuario, agora, id_))
        else:
            conn.execute(
                """INSERT INTO anotacoes (id, nome, data_ref, data_anterior, data_retorno, tipo_violacao,
                   anotacao, usuario_criacao, criado_em, atualizado_em)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (id_,
                 str(dados.get("nome", "")),
                 str(dados.get("data_ref", "")),
                 str(dados.get("data_anterior", "")),
                 str(dados.get("data_retorno", "")),
                 str(dados.get("tipo_violacao", "")),
                 anotacao, usuario, agora, agora),
            )
        conn.commit()
        return jsonify({"ok": True})
    finally:
        conn.close()
```

- [ ] **Step 4: Verify no syntax errors**

Run: `python -c "import py_compile; py_compile.compile('robo_ponto_web.py', doraise=True)"`
Expected: no output (success)

- [ ] **Step 5: Commit**

```bash
git add robo_ponto_web.py
git commit -m "feat: add anotacoes SQLite DB and GET/POST API routes"
```

---

### Task 2: Backend — Inject `_anotacao_id` into all 5 endpoint responses

**Files:**
- Modify: `robo_ponto_web.py` (inject `_anotacao_id` field into each row at 5 endpoints)

- [ ] **Step 1: Inject into `/api/consolidado`** (guias 1-2)

At the end of `api_consolidado()`, after computing `dados_resp`, inject the ID:

```python
dados_resp, meta_resp = _df_para_resposta_leve(df)
# Injetar _anotacao_id
for r in dados_resp:
    nome = str(r.get("nome", "") or "")
    data = str(r.get("data", "") or "")
    r["_anotacao_id"] = _gerar_anotacao_id(nome, data)
return jsonify({"erro": "", "dados": dados_resp, "kpis": kpis, **meta_resp})
```

- [ ] **Step 2: Inject into `/api/ponto-em-aberto`** (guia 3)

At the end of `api_ponto_em_aberto()`, same pattern:

```python
dados_resp, meta_resp = _df_para_resposta_leve(df)
for r in dados_resp:
    nome = str(r.get("nome", "") or "")
    data = str(r.get("data", "") or "")
    r["_anotacao_id"] = _gerar_anotacao_id(nome, data)
return jsonify({"erro": "", "dados": dados_resp, "kpis": _kpis_ponto_aberto(df), **meta_resp})
```

- [ ] **Step 3: Inject into `/api/ponto-aberto`** (guia 4)

At the end of `api_ponto_aberto()`, same injection:

```python
dados_resp, meta_resp = _df_para_resposta_leve(df)
for r in dados_resp:
    nome = str(r.get("nome", "") or "")
    data = str(r.get("data", "") or "")
    r["_anotacao_id"] = _gerar_anotacao_id(nome, data)
return jsonify({"erro": "", "dados": dados_resp, "kpis": { ... }, **meta_resp})
```

- [ ] **Step 4: Inject into `/api/violacoes-jornada`** (guia 5)

In `api_violacoes_jornada()`, the response data is in `resultados` list. After building it, inject IDs:

```python
for r in resultados:
    nome = str(r.get("nome", "") or "")
    data_ref = str(r.get("data_referencia", "") or "")
    tipo = str(r.get("violacao", "") or "").strip()
    if "12x36" in str(r.get("tipo_jornada", "")):
        r["_anotacao_id"] = _gerar_anotacao_id(nome, data_ref, "12x36")
    elif "semana" in str(r.get("violacao", "")).lower():
        periodo = str(r.get("periodo", "") or "")
        r["_anotacao_id"] = _gerar_anotacao_id(nome, periodo, "excesso_semanal")
    else:
        r["_anotacao_id"] = _gerar_anotacao_id(nome, data_ref, "violacao")

dados_resp = resultados[:LIMITE_LINHAS_RESPOSTA_TABELA]
```

- [ ] **Step 5: Inject into `/api/inter-jornada`** (guia 6)

In `api_inter_jornada()`, after building `resultados`, inject IDs:

```python
for r in resultados:
    nome = str(r.get("nome", "") or "")
    data_ant = str(r.get("data_anterior", "") or "")
    data_ret = str(r.get("data_retorno", "") or "")
    r["_anotacao_id"] = _gerar_anotacao_id(nome, f"{data_ant}|{data_ret}", "interjornada")

dados_resp = resultados[:LIMITE_LINHAS_RESPOSTA_TABELA]
```

- [ ] **Step 6: Verify syntax**

Run: `python -c "import py_compile; py_compile.compile('robo_ponto_web.py', doraise=True)"`
Expected: no output (success)

- [ ] **Step 7: Commit**

```bash
git add robo_ponto_web.py
git commit -m "feat: inject _anotacao_id into all 6 tab API responses"
```

---

### Task 3: Frontend — CSS for balloon and editor

**Files:**
- Modify: `static/style.css` (append styles at end)

- [ ] **Step 1: Add balloon and editor styles** at the end of `static/style.css`

```css
/* === Anotações (balão + editor inline) === */
.anotacao-balloon {
  cursor: pointer;
  font-size: 16px;
  user-select: none;
  transition: transform 0.15s;
  display: inline-block;
  line-height: 1;
}
.anotacao-balloon:hover {
  transform: scale(1.25);
}
.anotacao-balloon.has-text {
  color: #2563eb;
}
.anotacao-balloon.no-text {
  color: #9ca3af;
}

.anotacao-editor-row td {
  padding: 8px 12px !important;
  background: #f8fafc;
  border-bottom: 2px solid #2563eb;
}
.anotacao-editor-wrap {
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.anotacao-editor-wrap textarea {
  width: 100%;
  min-height: 60px;
  padding: 8px 10px;
  border: 1px solid #d1d5db;
  border-radius: 6px;
  font-family: inherit;
  font-size: 13px;
  resize: vertical;
  box-sizing: border-box;
}
.anotacao-editor-wrap textarea:focus {
  outline: none;
  border-color: #2563eb;
  box-shadow: 0 0 0 2px rgba(37,99,235,0.15);
}
.anotacao-editor-actions {
  display: flex;
  gap: 6px;
}
.anotacao-editor-actions .btn {
  font-size: 12px;
  padding: 4px 12px;
}
```

- [ ] **Step 2: Verify file is valid**

Run: No automated check — visual inspection is sufficient.

- [ ] **Step 3: Commit**

```bash
git add static/style.css
git commit -m "feat: add CSS styles for annotation balloon and inline editor"
```

---

### Task 4: Frontend — Balloon icon in render functions + annotation loading

**Files:**
- Modify: `static/app.js` (modify render functions + add annotation loading)

**Key insight:** All 6 tabs render data into `#tbody`. After any render function writes rows, we need to append the balloon column to each row's last cell, then load annotations.

Rather than modifying every individual render function, we'll:
1. Create a shared `_renderAnotacaoBalloon(anotacaoId, hasText)` that returns the balloon HTML
2. Create `_appendBaloesAposRender()` that runs after each tab's render, collects `_anotacao_id` from the JSON data, and appends balloon cells
3. Create `_carregarAnotacoes(ids)` that fetches annotations and updates balloon colors
4. Modify each `carregar*` function to call `_carregarAnotacoes` after data arrives

- [ ] **Step 1: Add helper functions** (add after the utility functions area, e.g. after function `renderizarCardsPorGuia()` around line 738)

```javascript
// v8.21.39 - Anotações do operador
function _htmlAnotacaoBalloon(anotacaoId, temAnotacao) {
  const cls = temAnotacao ? "has-text" : "no-text";
  return `<span class="anotacao-balloon ${cls}" data-anotacao-id="${anotacaoId}" onclick="abrirEditorAnotacao(this)">💬</span>`;
}

function _appendBaloesAposRender(ids) {
  const linhas = document.querySelectorAll("#tbody tr");
  const cabecalho = document.querySelector("#tabelaConsolidado thead tr");
  if (!cabecalho || !linhas.length) return;

  // Adiciona th de anotação se não existe
  if (!cabecalho.querySelector(".th-anotacao")) {
    const th = document.createElement("th");
    th.className = "th-anotacao";
    th.textContent = "";
    cabecalho.appendChild(th);
  }

  linhas.forEach((tr, i) => {
    if (tr.classList.contains("anotacao-editor-row")) return;
    const id = ids[i] || "";
    const temAnotacao = !!window._anotacoesCache?.[id];
    let td = tr.querySelector("td.td-anotacao");
    if (!td) {
      td = document.createElement("td");
      td.className = "td-anotacao";
      td.style.textAlign = "center";
      tr.appendChild(td);
    }
    td.innerHTML = _htmlAnotacaoBalloon(id, temAnotacao);
  });
}

async function _carregarAnotacoes(ids) {
  if (!ids || !ids.length) return;
  const val_ids = ids.filter(Boolean);
  if (!val_ids.length) return;
  try {
    const resp = await fetch(`/api/anotacoes?ids=${encodeURIComponent(val_ids.join(","))}`);
    const json = await resp.json();
    window._anotacoesCache = window._anotacoesCache || {};
    if (json.anotacoes) Object.assign(window._anotacoesCache, json.anotacoes);
    // Atualiza balões
    document.querySelectorAll(".anotacao-balloon").forEach(el => {
      const id = el.dataset.anotacaoId;
      if (window._anotacoesCache?.[id]) {
        el.classList.remove("no-text");
        el.classList.add("has-text");
      } else {
        el.classList.remove("has-text");
        el.classList.add("no-text");
      }
    });
  } catch (_) {}
}
```

- [ ] **Step 2: Add annotation loading call to `carregarConsolidado()`** (around line 1340-1370)

Find the block where `dadosConsolidadoAtual` is set and render happens. After `renderConsolidado()`, add:

```javascript
dadosConsolidadoAtual = json.dados || [];
renderConsolidado();
const ids = (json.dados || []).map(r => r._anotacao_id).filter(Boolean);
_appendBaloesAposRender(ids);
_carregarAnotacoes(ids);
```

- [ ] **Step 3: Add annotation loading to `carregarPontoEmAberto()`** (around line 1427-1431)

```javascript
dadosConsolidadoAtual = json.dados || [];
setCabecalhoSemMarcacaoCompacto();
renderSemMarcacaoCompacta(dadosConsolidadoAtual);
const ids = (json.dados || []).map(r => r._anotacao_id).filter(Boolean);
_appendBaloesAposRender(ids);
_carregarAnotacoes(ids);
```

- [ ] **Step 4: Add annotation loading to `carregarPontoAberto()`** (around line 1513-1515)

```javascript
dadosConsolidadoAtual = somenteImparesPontoAberto(json.dados || []);
setCabecalhoPontoAbertoCompacto();
renderPontoAbertoCompacto(dadosConsolidadoAtual);
const ids = (json.dados || []).map(r => r._anotacao_id).filter(Boolean);
_appendBaloesAposRender(ids);
_carregarAnotacoes(ids);
```

- [ ] **Step 5: Add annotation loading to `carregarViolacoesJornada()`** (find the render call, likely around line 2400-2420)

After the render call, add:

```javascript
const ids = (json.dados || []).map(r => r._anotacao_id).filter(Boolean);
_appendBaloesAposRender(ids);
_carregarAnotacoes(ids);
```

- [ ] **Step 6: Add annotation loading to `carregarInterJornada()`** (around line 1067-1070)

```javascript
dadosConsolidadoAtual = json.dados || [];
setCabecalhoInterJornada();
renderInterJornada(dadosConsolidadoAtual);
const ids = (json.dados || []).map(r => r._anotacao_id).filter(Boolean);
_appendBaloesAposRender(ids);
_carregarAnotacoes(ids);
```

- [ ] **Step 7: Commit**

```bash
git add static/app.js
git commit -m "feat: add balloon icon + annotation loading to all 6 tab renders"
```

---

### Task 5: Frontend — Inline editor on balloon click

**Files:**
- Modify: `static/app.js` (add `abrirEditorAnotacao` and `_salvarAnotacao` functions)

- [ ] **Step 1: Add `abrirEditorAnotacao` and `_salvarAnotacao` functions** (add near the helpers from Task 4)

```javascript
let _editorAberto = null;

function abrirEditorAnotacao(el) {
  if (_editorAberto) fecharEditorAnotacao();
  const id = el.dataset.anotacaoId;
  const tr = el.closest("tr");
  if (!tr) return;
  const textoAtual = (window._anotacoesCache?.[id] || "") + "";

  const editorTr = document.createElement("tr");
  editorTr.className = "anotacao-editor-row";
  editorTr.innerHTML = `
    <td colspan="99">
      <div class="anotacao-editor-wrap">
        <textarea placeholder="Digite sua anotação...">${escapeHtml(textoAtual)}</textarea>
        <div class="anotacao-editor-actions">
          <button class="btn" onclick="fecharEditorAnotacao()">Cancelar</button>
          <button class="btn" onclick="salvarEditorAnotacao()">💾 Salvar</button>
        </div>
      </div>
    </td>`;
  tr.parentNode.insertBefore(editorTr, tr.nextSibling);

  const ta = editorTr.querySelector("textarea");
  ta.focus();
  ta.setSelectionRange(ta.value.length, ta.value.length);

  _editorAberto = { id, tr, editorTr, el };
  ta.addEventListener("keydown", function (e) {
    if (e.key === "Escape") { fecharEditorAnotacao(); e.preventDefault(); }
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { salvarEditorAnotacao(); e.preventDefault(); }
  });
}

function fecharEditorAnotacao() {
  if (_editorAberto) {
    _editorAberto.editorTr.remove();
    _editorAberto = null;
  }
}

async function salvarEditorAnotacao() {
  if (!_editorAberto) return;
  const { id, el } = _editorAberto;
  const textarea = _editorAberto.editorTr.querySelector("textarea");
  const texto = textarea ? textarea.value.trim() : "";

  // Extrai dados para enviar ao backend
  const linha = el.closest("tr");
  const cells = linha ? linha.querySelectorAll("td") : [];
  const nome = cells.length > 1 ? (cells[1]?.textContent || "").trim() : "";
  const data = cells.length > 4 ? (cells[4]?.textContent || "").trim() : "";

  // Verifica tipo de violação (guia 5-6) pelo cabeçalho
  const ths = document.querySelectorAll("#tabelaConsolidado thead th");
  const temViolacao = Array.from(ths).some(th => th.textContent.includes("Violação"));
  const temInterJornada = Array.from(ths).some(th => th.textContent.includes("Intervalo"));

  const payload = { id, anotacao: texto, nome, data_ref: data };
  if (temViolacao) {
    const dataRef = cells.length > 5 ? (cells[5]?.textContent || "").trim() : "";
    const tipoCol = cells.length > 6 ? (cells[7]?.textContent || "").trim() : "";
    payload.data_ref = dataRef;
    payload.tipo_violacao = tipoCol;
  }
  if (temInterJornada) {
    const dataAnt = cells.length > 4 ? (cells[4]?.textContent || "").trim() : "";
    const dataRet = cells.length > 6 ? (cells[6]?.textContent || "").trim() : "";
    payload.data_anterior = dataAnt;
    payload.data_retorno = dataRet;
  }

  try {
    const resp = await fetch("/api/anotacoes", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const json = await resp.json();
    if (json.ok) {
      window._anotacoesCache = window._anotacoesCache || {};
      window._anotacoesCache[id] = texto;
      el.classList.remove("no-text");
      el.classList.add("has-text");
      fecharEditorAnotacao();
    } else {
      alert("Erro ao salvar anotação: " + (json.erro || "desconhecido"));
    }
  } catch (e) {
    alert("Erro de rede ao salvar anotação.");
  }
}
```

- [ ] **Step 2: Verify no syntax errors** (visual inspection)

- [ ] **Step 3: Commit**

```bash
git add static/app.js
git commit -m "feat: add inline annotation editor with save/cancel"
```

---

### Task 6: Self-Review & Final Verification

- [ ] **Step 1: Spec coverage check**

| Spec requirement | Covered in |
|-----------------|-----------|
| SQLite table `anotacoes` | Task 1 |
| Deterministic MD5 IDs | Task 1 (`_gerar_anotacao_id`) |
| IDs injected in 5 endpoints | Task 2 |
| `GET /api/anotacoes` | Task 1 |
| `POST /api/anotacoes` | Task 1 |
| Balloon icon per row | Task 4 (`_htmlAnotacaoBalloon`) |
| Balloon color (filled/outlined) | Task 4 (`has-text`/`no-text` classes) |
| Load annotations on tab open | Task 4 (`_carregarAnotacoes`) |
| Inline editor on click | Task 5 (`abrirEditorAnotacao`) |
| Ctrl+Enter save / Escape cancel | Task 5 |
| Editor row with colspan | Task 5 |

- [ ] **Step 2: No unused dependencies** — all code uses built-in Python modules and existing JS.

- [ ] **Step 3: Commit plan document**

```bash
git add docs/superpowers/plans/2026-07-01-campo-editavel-anotacoes-plan.md
git commit -m "docs: add implementation plan for annotation field"
```
