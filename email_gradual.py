"""Campanha de e-mail gradual — notificação progressiva de novos livros."""

import logging
import os
import re
from datetime import datetime

import requests
from dotenv import load_dotenv
from tenacity import retry, stop_after_attempt, wait_exponential

import auth

load_dotenv()

logger = logging.getLogger("biblioteca_ldcm")

SP_SITE_ID = os.getenv("SP_SITE_ID")
SP_LIST_ID = os.getenv("SP_LIST_ID")

EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE", "rodrigo.moreira@ldcm.com.br")
EMAIL_DESTINATARIOS = os.getenv("EMAIL_DESTINATARIOS", "")
DATA_CORTE = os.getenv("DATA_CORTE", "2026-03-30")

GRAPH_URL = "https://graph.microsoft.com/v1.0"
LOTE_TAMANHO = 100

# Logo LDCM horizontal (branca, fundo preto) — usar com mix-blend-mode:screen
_logo_path = os.path.join(os.path.dirname(__file__), "logo_base64.txt")
if os.path.exists(_logo_path):
    with open(_logo_path, "r") as _f:
        LOGO_DATA_URI = _f.read().strip()
else:
    LOGO_DATA_URI = ""


# ── 1. Buscar todos os livros da campanha ──────────────────────────


def _headers(token: str, extra: dict | None = None) -> dict:
    h = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    if extra:
        h.update(extra)
    return h


def _limpar_autor(html: str) -> str:
    """Extrai texto limpo de HTML do campo Autor/autores do SharePoint.

    Converte <div>, <p>, <br> em separadores e remove todas as tags.
    Ex: '<div class="...">Orlando Gomes<p>Antunes Varela</p></div>'
        -> 'Orlando Gomes; Antunes Varela'
    """
    if not html:
        return ""
    # Substituir tags de bloco por separador
    texto = re.sub(r'</?(div|p|br)\b[^>]*/?>', '\n', html, flags=re.IGNORECASE)
    # Remover qualquer tag restante
    texto = re.sub(r'<[^>]+>', '', texto)
    # Decode HTML entities
    texto = texto.replace('&#58;', ':').replace('&amp;', '&')
    # Limpar linhas e juntar com "; "
    partes = [p.strip() for p in texto.split('\n') if p.strip()]
    return '; '.join(partes)


def buscar_livros_campanha(token: str) -> list[dict]:
    """Retorna todos os livros Created >= DATA_CORTE, ordenados por número."""
    url = f"{GRAPH_URL}/sites/{SP_SITE_ID}/lists/{SP_LIST_ID}/items"
    campos = (
        "Id,_x0023_,Title,Autor_x002f_autores,Autores_email,Editora,"
        "Anodepublica_x00e7__x00e3_o,Link,Created"
    )
    params = {
        "$expand": f"fields($select={campos})",
        "$top": "200",
    }

    all_items = []
    page_url = url
    page_params = params
    while page_url:
        resp = requests.get(
            page_url, headers=_headers(token), params=page_params, timeout=60
        )
        resp.raise_for_status()
        data = resp.json()
        all_items.extend(data.get("value", []))
        page_url = data.get("@odata.nextLink")
        page_params = None

    # Filtrar: Created >= DATA_CORTE
    corte = datetime.fromisoformat(f"{DATA_CORTE}T00:00:00Z")
    livros = []
    for item in all_items:
        fields = item.get("fields", {})
        created_str = fields.get("Created", "")
        if not created_str:
            continue
        try:
            created = datetime.fromisoformat(created_str.replace("Z", "+00:00"))
        except ValueError:
            continue
        if created < corte.replace(tzinfo=created.tzinfo):
            continue

        link_field = fields.get("Link", "")
        if isinstance(link_field, dict):
            link_url = link_field.get("Url", "")
        elif isinstance(link_field, str) and link_field:
            link_url = link_field.split(", ")[0].strip()
        else:
            link_url = ""

        ano_raw = fields.get("Anodepublica_x00e7__x00e3_o", "")
        ano = str(int(ano_raw)) if ano_raw else ""

        # Autores_email (formatado pelo Power Automate) tem prioridade;
        # fallback: limpar HTML do campo Autor/autores
        autor = fields.get("Autores_email", "") or ""
        if not autor.strip():
            autor = _limpar_autor(fields.get("Autor_x002f_autores", ""))

        livros.append({
            "NumeroLista": int(fields.get("_x0023_", 0)),
            "Title": fields.get("Title", ""),
            "Autor": autor,
            "Editora": fields.get("Editora", ""),
            "Ano": ano,
            "Link": link_url,
        })

    livros.sort(key=lambda x: x["NumeroLista"])
    return livros


# ── 2. Offset ──────────────────────────────────────────────────────


def ler_offset() -> int:
    path = os.path.join(os.path.dirname(__file__), "offset.txt")
    if os.path.exists(path):
        return int(open(path).read().strip())
    return 0


def salvar_offset(valor: int) -> None:
    path = os.path.join(os.path.dirname(__file__), "offset.txt")
    with open(path, "w") as f:
        f.write(str(valor))


# ── 3. Montar HTML do e-mail ──────────────────────────────────────


def _linha_livro(livro: dict) -> str:
    titulo = livro["Title"] or "Sem titulo"
    autor = livro["Autor"] or ""
    editora = livro["Editora"] or ""
    ano = livro["Ano"] or ""
    link = livro["Link"] or ""

    link_html = (
        f'<a href="{link}" style="color:#3D5549;text-decoration:none;'
        f'font-weight:bold">Acessar &#8594;</a>'
    ) if link else ""

    return (
        f'<tr style="border-bottom:1px solid #ADA688">\n'
        f'<td style="padding:10px 12px;color:#000000">{titulo}</td>\n'
        f'<td style="padding:10px 12px;color:#77787B">{autor}</td>\n'
        f'<td style="padding:10px 12px;color:#77787B">{editora}</td>\n'
        f'<td style="padding:10px 12px;color:#77787B">{ano}</td>\n'
        f'<td style="padding:10px 12px">{link_html}</td>\n'
        f'</tr>'
    )


def montar_html(lote: list[dict]) -> str:
    qtd = len(lote)
    linhas = "\n".join(_linha_livro(livro) for livro in lote)

    th_style = (
        'padding:10px 12px;text-align:left;font-weight:normal;'
        'letter-spacing:1px;text-transform:uppercase;font-size:11px'
    )

    return (
        '<table width="100%" cellpadding="0" cellspacing="0" style="font-family:Calibri,Arial,sans-serif;background:#f4f3ef;">'
        '<tr><td align="center" style="padding:32px 16px 0;">'
        '<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:2px;overflow:hidden;">'
        '<tr><td style="background:#3D5549;padding:36px 40px 32px;text-align:center;">'
        f'<img alt="LDCM Advogados" width="160" style="display:block;margin:0 auto;mix-blend-mode:screen;max-width:160px;" src="{LOGO_DATA_URI}">'
        '</td></tr>'
        '<tr><td style="padding:36px 40px;color:#2a2a2a;font-size:15px;line-height:1.7;">'
        f'<h2 style="color:#3D5549;border-bottom:2px solid #ADA688;padding-bottom:8px;margin-top:0;">Biblioteca LDCM &#8212; {qtd} novos t&#237;tulos catalogados</h2>'
        '<p style="color:#77787B;font-size:14px;">Novos t&#237;tulos foram catalogados e adicionados ao acervo. Confira abaixo:</p>'
        '<table style="border-collapse:collapse;width:100%;font-size:13px;">'
        '<tbody>'
        '<tr style="background-color:#3D5549;color:#ffffff;">'
        f'<th style="{th_style}">T&#237;tulo</th>'
        f'<th style="{th_style}">Autores</th>'
        f'<th style="{th_style}">Editora</th>'
        f'<th style="{th_style}">Ano</th>'
        f'<th style="{th_style}">Link</th>'
        '</tr>'
        f'{linhas}'
        '</tbody>'
        '</table>'
        '<p style="margin-top:24px;margin-bottom:0;color:#3D5549;font-weight:600;">Comit&#234; de TI e Biblioteca</p>'
        '<p style="margin-top:2px;color:#77787B;font-size:13px;">LDCM Advogados</p>'
        '</td></tr>'
        '<tr><td style="background:#3D5549;padding:20px 40px;text-align:center;">'
        '<span style="font-family:Arial,sans-serif;font-size:10px;color:#ADA688;letter-spacing:0.15em;">'
        'S&#195;O PAULO | RIO DE JANEIRO | LDCM.COM.BR'
        '</span>'
        '</td></tr>'
        '</table>'
        '</td></tr>'
        '</table>'
    )


# ── 4. Enviar e-mail via Graph API ─────────────────────────────────


def _parse_destinatarios() -> list[dict]:
    return [
        {"emailAddress": {"address": e.strip()}}
        for e in EMAIL_DESTINATARIOS.split(";")
        if e.strip()
    ]


@retry(stop=stop_after_attempt(3), wait=wait_exponential(min=2, max=30))
def enviar_email(token: str, html: str, qtd_livros: int) -> int:
    destinatarios = _parse_destinatarios()
    if not destinatarios:
        raise ValueError("EMAIL_DESTINATARIOS nao configurado")

    url = f"{GRAPH_URL}/users/{EMAIL_REMETENTE}/sendMail"
    body = {
        "message": {
            "subject": f"Biblioteca LDCM \u2014 {qtd_livros} novos t\u00edtulos catalogados",
            "body": {"contentType": "HTML", "content": html},
            "from": {"emailAddress": {"address": "TI.biblioteca@ldcm.com.br"}},
            "toRecipients": destinatarios,
        },
        "saveToSentItems": True,
    }

    resp = requests.post(
        url,
        headers=_headers(token, {"Content-Type": "application/json"}),
        json=body,
        timeout=60,
    )
    if not resp.ok:
        raise requests.HTTPError(
            f"Falha ao enviar e-mail: {resp.status_code} - {resp.text}", response=resp
        )
    return resp.status_code
