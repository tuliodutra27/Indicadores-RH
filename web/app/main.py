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
from comparador import processar_planilha, gerar_excel, carregar_planilha, gerar_organograma_data

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
    """Exibe o organograma hierárquico do efetivo a partir de um snapshot."""
    snapshots = list(SNAPSHOTS_DIR.glob(f"snapshot_{id}_*.xlsx"))
    if not snapshots:
        flash("Snapshot não encontrado para esta comparação.", "warning")
        return redirect(url_for("dashboard"))

    snapshot = snapshots[0]
    df = carregar_planilha(snapshot)

    historico = carregar_historico()
    meta = next((h for h in historico if h["id"] == id), {})

    colab_list = []
    for chave, row in df.iterrows():
        tipo = "PJ" if chave.startswith("PJ::") else "CLT"
        colab_list.append({
            "tipo":         tipo,
            "nome":         str(row.get("NOME", "")),
            "cargo":        str(row.get("CARGO", "")),
            "departamento": str(row.get("DEPARTAMENTO", "")),
            "gestor":       str(row.get("GESTOR", "")),
        })

    arvore = gerar_organograma_data(colab_list)

    return render_template(
        "organograma.html",
        arvore=arvore,
        meta=meta,
        total=len(colab_list),
    )


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
