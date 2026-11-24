import io
import json
import os
import re
import unicodedata
import uuid
import zipfile
from datetime import datetime
from docx import Document
from flask import Flask, flash, redirect, render_template, request, send_file, send_from_directory, session, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "chave-secreta-portfolio"

SENHA_ADM = "1234"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
DATA_FILE = os.path.join(BASE_DIR, "dados.json")

EXTENSOES_PERMITIDAS = {
    "pdf", "doc", "docx", "ppt", "pptx", "xls", "xlsx", "csv",
    "txt", "md", "py", "html", "css", "js", "ts", "jsx", "tsx",
    "json", "xml", "sql", "java", "c", "cpp", "cs", "php",
    "zip", "rar", "7z", "png", "jpg", "jpeg", "gif", "webp", "svg"
}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # Limite de 50 MB por envio

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def normalizar(texto):
    return unicodedata.normalize("NFC", str(texto or "")).strip().lower()


def extensao_permitida(nome_arquivo):
    return "." in nome_arquivo and nome_arquivo.rsplit(".", 1)[1].lower() in EXTENSOES_PERMITIDAS


def nome_arquivo_limpo(titulo, extensao):
    limpo = re.sub(r'[\\/*?:"<>|]', "", titulo).strip()
    return f"{limpo or 'atividade'}.{extensao}"


def nome_seguro_arquivo(nome_arquivo):
    sem_acento = unicodedata.normalize("NFKD", nome_arquivo).encode("ascii", "ignore").decode("ascii")
    return secure_filename(sem_acento) or "arquivo"


def normalizar_atividade(item):
    """Garante compatibilidade com atividades antigas (1 arquivo) e novas (vários arquivos)."""
    if "arquivos" not in item:
        item["arquivos"] = [
            {
                "id": "0",
                "original": item.get("arquivo_original", "arquivo"),
                "salvo": item.get("arquivo_salvo", ""),
            }
        ]
    return item


def carregar_atividades():
    if not os.path.exists(DATA_FILE):
        return []
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        try:
            dados = json.load(f)
            return [normalizar_atividade(a) for a in dados]
        except json.JSONDecodeError:
            return []


def salvar_atividades(atividades):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(atividades, f, ensure_ascii=False, indent=2)


def buscar_arquivo(atividade_id, arquivo_id):
    atividades = carregar_atividades()
    atividade = next((a for a in atividades if a["id"] == atividade_id), None)
    if not atividade:
        return None, None
    arquivo = next((arq for arq in atividade["arquivos"] if arq["id"] == arquivo_id), None)
    return atividade, arquivo


@app.route("/")
def index():
    atividades = carregar_atividades()
    busca = request.args.get("busca", "").strip()
    materia_filtro = request.args.get("materia", "").strip()

    busca_norm = normalizar(busca)
    materia_norm = normalizar(materia_filtro)

    materias_unicas = {}
    for a in atividades:
        mat = str(a.get("materia", "")).strip()
        if mat:
            chave = normalizar(mat)
            if chave not in materias_unicas:
                materias_unicas[chave] = mat
    materias_disponiveis = sorted(materias_unicas.values(), key=normalizar)

    filtradas = atividades
    if busca_norm:
        filtradas = [
            a for a in filtradas
            if busca_norm in normalizar(a.get("titulo"))
            or busca_norm in normalizar(a.get("materia"))
            or any(busca_norm in normalizar(arq.get("original")) for arq in a.get("arquivos", []))
        ]
    if materia_norm:
        filtradas = [
            a for a in filtradas
            if normalizar(a.get("materia")) == materia_norm
        ]

    return render_template(
        "index.html",
        atividades=filtradas,
        total=len(atividades),
        materias=materias_disponiveis,
        busca=busca,
        materia_filtro=materia_filtro,
        materia_norm=materia_norm,
        is_admin=session.get("is_admin", False),
    )


@app.route("/login-adm", methods=["POST"])
def login_adm():
    senha = request.form.get("senha", "").strip()
    if senha == SENHA_ADM:
        session["is_admin"] = True
        flash("Modo ADM ativado.", "sucesso")
    else:
        flash("Senha incorreta.", "erro")
    return redirect(url_for("index"))


@app.route("/logout-adm", methods=["POST"])
def logout_adm():
    session.pop("is_admin", None)
    flash("Você saiu do modo ADM.", "sucesso")
    return redirect(url_for("index"))


@app.route("/upload", methods=["POST"])
def upload():
    if not session.get("is_admin"):
        flash("Acesso restrito ao ADM.", "erro")
        return redirect(url_for("index"))

    titulo = request.form.get("titulo", "").strip()
    materia = request.form.get("materia", "").strip() or "Geral"
    modo = request.form.get("modo", "arquivo")
    data_atual = datetime.now().strftime("%d/%m/%Y às %H:%M")

    if not titulo:
        flash("Informe um título para a atividade.", "erro")
        return redirect(url_for("index"))

    atividade_id = uuid.uuid4().hex[:8]
    arquivos_salvos = []

    if modo == "texto":
        texto = request.form.get("texto_atividade", "").strip()
        if not texto:
            flash("Digite ou cole o texto da atividade.", "erro")
            return redirect(url_for("index"))

        arquivo_id = uuid.uuid4().hex[:6]
        arquivo_original = nome_arquivo_limpo(titulo, "docx")
        nome_salvo = f"{atividade_id}_{arquivo_id}_{nome_seguro_arquivo(arquivo_original)}"
        caminho_completo = os.path.join(app.config["UPLOAD_FOLDER"], nome_salvo)

        doc = Document()
        doc.add_heading(titulo, level=1)
        p_meta = doc.add_paragraph(f"Matéria: {materia}  |  Data: {data_atual}")
        if p_meta.runs:
            p_meta.runs[0].italic = True
        doc.add_paragraph("")

        for linha in texto.splitlines():
            doc.add_paragraph(linha)

        doc.save(caminho_completo)
        arquivos_salvos.append({
            "id": arquivo_id,
            "original": arquivo_original,
            "salvo": nome_salvo,
        })
        mensagem_sucesso = "Texto convertido em Word (.docx) e adicionado ao portfólio!"

    else:
        arquivos_enviados = [
            f for f in request.files.getlist("arquivos")
            if f and f.filename and f.filename.strip()
        ]

        if not arquivos_enviados:
            flash("Selecione pelo menos um arquivo para enviar.", "erro")
            return redirect(url_for("index"))

        for arquivo in arquivos_enviados:
            if not extensao_permitida(arquivo.filename):
                flash(f"Formato não permitido no arquivo: {arquivo.filename}", "erro")
                return redirect(url_for("index"))

        for arquivo in arquivos_enviados:
            arquivo_id = uuid.uuid4().hex[:6]
            nome_seguro = nome_seguro_arquivo(arquivo.filename)
            nome_salvo = f"{atividade_id}_{arquivo_id}_{nome_seguro}"
            caminho_completo = os.path.join(app.config["UPLOAD_FOLDER"], nome_salvo)
            arquivo.save(caminho_completo)
            arquivos_salvos.append({
                "id": arquivo_id,
                "original": arquivo.filename,
                "salvo": nome_salvo,
            })

        qtd = len(arquivos_salvos)
        if qtd == 1:
            mensagem_sucesso = "Atividade adicionada ao portfólio!"
        else:
            mensagem_sucesso = f"Atividade adicionada com {qtd} arquivos no mesmo registro!"

    nova_atividade = {
        "id": atividade_id,
        "titulo": titulo,
        "materia": materia,
        "arquivos": arquivos_salvos,
        "data": data_atual,
    }

    atividades = carregar_atividades()
    atividades.insert(0, nova_atividade)
    salvar_atividades(atividades)

    flash(mensagem_sucesso, "sucesso")
    return redirect(url_for("index"))


@app.route("/visualizar/<atividade_id>/<arquivo_id>")
def visualizar(atividade_id, arquivo_id):
    _, arquivo = buscar_arquivo(atividade_id, arquivo_id)
    if not arquivo:
        flash("Arquivo não encontrado.", "erro")
        return redirect(url_for("index"))

    return send_from_directory(
        app.config["UPLOAD_FOLDER"],
        arquivo["salvo"],
        as_attachment=False,
    )


@app.route("/baixar/<atividade_id>/<arquivo_id>")
def baixar(atividade_id, arquivo_id):
    _, arquivo = buscar_arquivo(atividade_id, arquivo_id)
    if not arquivo:
        flash("Arquivo não encontrado.", "erro")
        return redirect(url_for("index"))

    return send_from_directory(
        app.config["UPLOAD_FOLDER"],
        arquivo["salvo"],
        as_attachment=True,
        download_name=arquivo["original"],
    )


@app.route("/baixar-tudo/<atividade_id>")
def baixar_tudo(atividade_id):
    atividades = carregar_atividades()
    atividade = next((a for a in atividades if a["id"] == atividade_id), None)
    if not atividade:
        flash("Atividade não encontrada.", "erro")
        return redirect(url_for("index"))

    memoria = io.BytesIO()
    nomes_usados = {}

    with zipfile.ZipFile(memoria, "w", zipfile.ZIP_DEFLATED) as zf:
        for arq in atividade["arquivos"]:
            caminho = os.path.join(app.config["UPLOAD_FOLDER"], arq["salvo"])
            if os.path.exists(caminho):
                nome_interno = arq["original"]
                if nome_interno in nomes_usados:
                    nomes_usados[nome_interno] += 1
                    base, ext = os.path.splitext(nome_interno)
                    nome_interno = f"{base} ({nomes_usados[nome_interno]}){ext}"
                else:
                    nomes_usados[nome_interno] = 0
                zf.write(caminho, arcname=nome_interno)

    memoria.seek(0)
    nome_zip = nome_arquivo_limpo(atividade["titulo"], "zip")
    return send_file(
        memoria,
        mimetype="application/zip",
        as_attachment=True,
        download_name=nome_zip,
    )


@app.route("/excluir/<atividade_id>", methods=["POST"])
def excluir(atividade_id):
    if not session.get("is_admin"):
        flash("Acesso restrito ao ADM.", "erro")
        return redirect(url_for("index"))

    atividades = carregar_atividades()
    atividade = next((a for a in atividades if a["id"] == atividade_id), None)

    if atividade:
        for arq in atividade.get("arquivos", []):
            caminho = os.path.join(app.config["UPLOAD_FOLDER"], arq["salvo"])
            if os.path.exists(caminho):
                os.remove(caminho)
        atividades = [a for a in atividades if a["id"] != atividade_id]
        salvar_atividades(atividades)
        flash("Atividade e seus arquivos foram removidos.", "sucesso")

    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True)