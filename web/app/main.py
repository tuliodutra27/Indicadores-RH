"""
Indicadores RH — AliseoSA
Aplicação web Flask para comparação de planilhas do RH.
"""

import os
import json
from pathlib import Path
from datetime import datetime
from functools import wraps

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    flash,
    send_file,
)
from werkzeug.utils import secure_filename
from io import BytesIO

from auth import autenticar_ad
from comparador import (
    processar_planilha, gerar_excel, carregar_planilha,
    carregar_estrutura_org, salvar_posicao_org, salvar_posicoes_org,
    salvar_relacoes_org, gerar_organograma_hibrido,
    validar_posicao_org, carregar_historico_org,
    restaurar_snapshot_org, importar_estrutura_org,
)

# ── App ────────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-troque-em-producao")
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

# ── Diretórios de dados ───────────────────────────────────────────────────
DATA_DIR       = Path(os.environ.get("DATA_DIR", "/data"))
SNAPSHOTS_DIR  = DATA_DIR / "snapshots"
RELATORIOS_DIR = DATA_DIR / "relatorios"
UPLOADS_DIR    = DATA_DIR / "uploads"
HISTORICO_FILE = DATA_DIR / "historico.json"

ALLOWED_EXT = {"xlsx", "xls"}


# ── Helpers ────────────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "usuario" not in session:
            flash("Faça login para continuar.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXT


def carregar_historico() -> list:
    if HISTORICO_FILE.exists():
        with open(HISTORICO_FILE, encoding="utf-8") as f:
            return json.load(f)
    return []


def salvar_historico(historico: list) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(HISTORICO_FILE, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, indent=2)


# ── Rotas ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    if "usuario" in session:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    if "usuario" in session:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        usuario = request.form.get("usuario", "").strip().lower()
        senha   = request.form.get("senha", "")

        if not usuario or not senha:
            flash("Informe usuário e senha.", "danger")
            return render_template("login.html")

        ok, resultado = autenticar_ad(usuario, senha)

        if ok:
            session["usuario"] = usuario
            session["nome"]    = resultado
            return redirect(url_for("dashboard"))
        else:
            flash(resultado, "danger")

    return render_template("login.html")


@app.route("/logout")
def logout():
    nome = session.get("nome", session.get("usuario", ""))
    session.clear()
    flash(f"Até logo, {nome}!", "info")
    return redirect(url_for("login"))


@app.route("/dashboard")
@login_required
def dashboard():
    historico = carregar_historico()
    historico = sorted(historico, key=lambda x: x.get("timestamp", ""), reverse=True)
    return render_template("dashboard.html", historico=historico)


@app.route("/upload", methods=["POST"])
@login_required
def upload():
    if "arquivo" not in request.files:
        flash("Nenhum arquivo enviado.", "danger")
        return redirect(url_for("dashboard"))

    file = request.files["arquivo"]
    if not file or file.filename == "":
        flash("Nenhum arquivo selecionado.", "danger")
        return redirect(url_for("dashboard"))

    if not allowed_file(file.filename):
        flash("Formato inválido. Envie um arquivo .xlsx ou .xls.", "danger")
        return redirect(url_for("dashboard"))

    # Salva o arquivo enviado
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = secure_filename(file.filename)
    saved    = UPLOADS_DIR / f"{ts}_{filename}"
    file.save(saved)

    try:
        resultado = processar_planilha(
            saved, SNAPSHOTS_DIR, RELATORIOS_DIR,
            usuario=session.get("usuario", "sistema"),
        )

        # Atualiza histórico (sem as listas detalhadas para manter o arquivo pequeno)
        historico = carregar_historico()
        historico.append({
            "id":           resultado["id"],
            "data":         resultado["data"],
            "arquivo":      resultado["arquivo"],
            "timestamp":    resultado["timestamp"],
            "usuario":      resultado["usuario"],
            "total":        resultado["total"],
            "total_clt":    resultado["total_clt"],
            "total_pj":     resultado["total_pj"],
            "n_adicoes":    resultado["n_adicoes"],
            "n_remocoes":   resultado["n_remocoes"],
            "n_alteracoes": resultado["n_alteracoes"],
            "is_baseline":  resultado["is_baseline"],
        })
        salvar_historico(historico)

        if resultado["is_baseline"]:
            flash(
                f"Baseline criada com sucesso! "
                f"{resultado['total']} colaboradores registrados "
                f"({resultado['total_clt']} CLT, {resultado['total_pj']} PJ). "
                "O próximo envio mostrará as diferenças.",
                "info",
            )
            return redirect(url_for("dashboard"))

        if "alerta_chave" in resultado:
            flash(resultado["alerta_chave"], "warning")

        return redirect(url_for("resultado", id=resultado["id"]))

    except Exception as exc:  # noqa: BLE001
        flash(f"Erro ao processar planilha: {exc}", "danger")
        return redirect(url_for("dashboard"))


@app.route("/resultado/<id>")
@login_required
def resultado(id):
    json_path = RELATORIOS_DIR / f"resultado_{id}.json"
    if not json_path.exists():
        flash("Resultado não encontrado.", "danger")
        return redirect(url_for("dashboard"))

    with open(json_path, encoding="utf-8") as f:
        dados = json.load(f)

    return render_template("resultado.html", dados=dados)


@app.route("/download/<id>")
@login_required
def download(id):
    json_path = RELATORIOS_DIR / f"resultado_{id}.json"
    if not json_path.exists():
        flash("Resultado não encontrado.", "danger")
        return redirect(url_for("dashboard"))

    with open(json_path, encoding="utf-8") as f:
        dados = json.load(f)

    xlsx_bytes = gerar_excel(dados)
    buf = BytesIO(xlsx_bytes)
    buf.seek(0)

    nome_arquivo = f"relatorio_rh_{dados['data']}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/colaboradores/<id>")
@login_required
def colaboradores(id):
    """Exibe todos os colaboradores do snapshot de uma comparação."""
    snapshots = list(SNAPSHOTS_DIR.glob(f"snapshot_{id}_*.xlsx"))
    if not snapshots:
        flash("Snapshot não encontrado para esta comparação.", "warning")
        return redirect(url_for("dashboard"))

    snapshot = snapshots[0]
    df = carregar_planilha(snapshot)

    # Metadados do histórico para exibir data/arquivo/usuário
    historico = carregar_historico()
    meta = next((h for h in historico if h["id"] == id), {})

    # Monta lista ordenada por nome
    colaboradores_list = []
    for chave, row in df.iterrows():
        tipo = "PJ" if chave.startswith("PJ::") else "CLT"
        colaboradores_list.append({
            "tipo":         tipo,
            "matricula":    str(row.get("MATRÍCULA", "")),
            "nome":         str(row.get("NOME", "")),
            "cargo":        str(row.get("CARGO", "")),
            "departamento": str(row.get("DEPARTAMENTO", "")),
            "gestor":       str(row.get("GESTOR", "")),
            "admissao":     str(row.get("ADMISSÃO", "")),
        })

    colaboradores_list.sort(key=lambda x: x["nome"])

    total_clt = sum(1 for c in colaboradores_list if c["tipo"] == "CLT")
    total_pj  = sum(1 for c in colaboradores_list if c["tipo"] == "PJ")

    return render_template(
        "colaboradores.html",
        colaboradores=colaboradores_list,
        meta=meta,
        total=len(colaboradores_list),
        total_clt=total_clt,
        total_pj=total_pj,
    )


@app.route("/organograma/<id>")
@login_required
def organograma(id):
    """Exibe o organograma hierárquico do efetivo com edição manual de posições."""
    snapshots = list(SNAPSHOTS_DIR.glob(f"snapshot_{id}_*.xlsx"))
    if not snapshots:
        flash("Snapshot não encontrado para esta comparação.", "warning")
        return redirect(url_for("dashboard"))

    df = carregar_planilha(snapshots[0])
    historico = carregar_historico()
    meta = next((h for h in historico if h["id"] == id), {})

    # Monta lista de colaboradores com a chave do snapshot
    colab_list = []
    for chave, row in df.iterrows():
        tipo = "PJ" if chave.startswith("PJ::") else "CLT"
        colab_list.append({
            "chave":        chave,
            "tipo":         tipo,
            "nome":         str(row.get("NOME", "")),
            "cargo":        str(row.get("CARGO", "")),
            "departamento": str(row.get("DEPARTAMENTO", "")),
        })

    # Carrega posições manuais salvas e gera organograma híbrido
    struct    = carregar_estrutura_org(DATA_DIR)
    resultado = gerar_organograma_hibrido(colab_list, struct["posicoes"], struct["relacoes"])

    # Lista completa para o dropdown de seleção de pai no modal
    todos_colab = sorted(
        [{"chave": c["chave"], "nome": c["nome"], "cargo": c["cargo"]}
         for c in colab_list],
        key=lambda x: x["nome"],
    )
    # Adiciona o nó raiz como opção válida de pai
    todos_colab.insert(0, {
        "chave": "__ROOT__",
        "nome":  "ALISEO SA",
        "cargo": "Conselho de Administração",
    })

    return render_template(
        "organograma.html",
        arvore         = resultado["nodes"],
        arestas_extras = resultado["arestas_extras"],
        todos_colab    = todos_colab,
        n_manual       = resultado["n_manual"],
        n_sugerido     = resultado["n_sugerido"],
        meta           = meta,
        total          = len(colab_list),
    )


# ── API: edição do organograma ─────────────────────────────────────────────

@app.route("/api/organograma/posicao", methods=["POST"])
@login_required
def api_salvar_posicao_org():
    """Salva a posição hierárquica de UM colaborador (AJAX)."""
    data = request.get_json(silent=True) or {}
    chave        = data.get("chave", "").strip()
    parent_chave = data.get("parentChave", "").strip()
    nome_pessoa  = data.get("nomePessoa", chave)
    if not chave or chave == "__ROOT__":
        return {"ok": False, "erro": "chave inválida"}, 400

    # Valida se não cria ciclo
    struct = carregar_estrutura_org(DATA_DIR)
    ok_ciclo, msg_ciclo = validar_posicao_org(struct["posicoes"], chave, parent_chave)
    if not ok_ciclo:
        return {"ok": False, "erro": msg_ciclo, "ciclo": True}, 409

    pai_label = parent_chave if parent_chave and parent_chave != "__ROOT__" else "Conselho"
    acao = f"Posição: {nome_pessoa} → {pai_label}"
    try:
        salvar_posicao_org(DATA_DIR, chave, parent_chave, session.get("usuario", ""), acao=acao)
        return {"ok": True}
    except Exception as exc:
        return {"ok": False, "erro": str(exc)}, 500


@app.route("/api/organograma/relacoes", methods=["POST"])
@login_required
def api_salvar_relacoes_org():
    """Salva os gestores adicionais (múltiplos) de UM colaborador (AJAX)."""
    data       = request.get_json(silent=True) or {}
    filho      = data.get("filho", "").strip()
    pais       = data.get("pais", [])
    nome_filho = data.get("nomePessoa", filho)
    if not filho or filho == "__ROOT__":
        return {"ok": False, "erro": "filho inválido"}, 400
    n = len(pais)
    acao = f"Gestores adicionais de {nome_filho}: {n} selecionado(s)"
    try:
        salvar_relacoes_org(DATA_DIR, filho, pais, session.get("usuario", ""), acao=acao)
        return {"ok": True}
    except Exception as exc:
        return {"ok": False, "erro": str(exc)}, 500


@app.route("/api/organograma/historico", methods=["GET"])
@login_required
def api_historico_org():
    """Retorna o log de alterações do organograma."""
    return {"historico": carregar_historico_org(DATA_DIR)}


@app.route("/api/organograma/restaurar", methods=["POST"])
@login_required
def api_restaurar_org():
    """Restaura o organograma para um snapshot específico."""
    data = request.get_json(silent=True) or {}
    ts   = data.get("ts", "").strip()
    if not ts:
        return {"ok": False, "erro": "timestamp inválido"}, 400
    ok, msg = restaurar_snapshot_org(DATA_DIR, ts, session.get("usuario", ""))
    if ok:
        return {"ok": True, "msg": msg}
    return {"ok": False, "erro": msg}, 400


@app.route("/api/organograma/importar", methods=["POST"])
@login_required
def api_importar_org():
    """Importa/substitui toda a estrutura do organograma (JSON completo)."""
    data     = request.get_json(silent=True) or {}
    posicoes = data.get("posicoes")
    relacoes = data.get("relacoes_adicionais", data.get("relacoes", []))
    if posicoes is None:
        return {"ok": False, "erro": "JSON inválido: campo 'posicoes' ausente."}, 400
    try:
        importar_estrutura_org(DATA_DIR, posicoes, relacoes, session.get("usuario", ""))
        return {"ok": True}
    except Exception as exc:
        return {"ok": False, "erro": str(exc)}, 500


@app.route("/api/organograma/aceitar-sugestoes", methods=["POST"])
@login_required
def api_aceitar_sugestoes_org():
    """Confirma em lote as posições sugeridas automaticamente (AJAX)."""
    data     = request.get_json(silent=True) or {}
    posicoes = data.get("posicoes", {})
    if not posicoes:
        return {"ok": False, "erro": "nenhuma posição enviada"}, 400
    try:
        salvar_posicoes_org(DATA_DIR, posicoes, session.get("usuario", ""))
        return {"ok": True}
    except Exception as exc:
        return {"ok": False, "erro": str(exc)}, 500


@app.route("/historico/<id>/excluir", methods=["POST"])
@login_required
def excluir_historico(id):
    """Remove uma entrada do histórico (mas mantém o snapshot)."""
    historico = carregar_historico()
    historico = [h for h in historico if h["id"] != id]
    salvar_historico(historico)
    flash("Entrada removida do histórico.", "success")
    return redirect(url_for("dashboard"))


if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5000)
