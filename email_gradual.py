"""Campanha de e-mail gradual — notificação progressiva de novos livros."""

import logging
import math
import os
from datetime import datetime

import requests
from dotenv import load_dotenv
from tenacity import retry, stop_after_attempt, wait_exponential

import auth

load_dotenv()

logger = logging.getLogger("biblioteca_ldcm")

SP_SITE_ID = os.getenv("SP_SITE_ID")
SP_LIST_ID = os.getenv("SP_LIST_ID")

EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE", "biblioteca@ldcm.com.br")
EMAIL_DESTINATARIOS = os.getenv("EMAIL_DESTINATARIOS", "")
DATA_CORTE = os.getenv("DATA_CORTE", "2026-03-30")

GRAPH_URL = "https://graph.microsoft.com/v1.0"
PERCENTUAL_LOTE = 0.05
TETO_LOTE = 100

# Logo LDCM horizontal (branca, fundo preto) — usar com mix-blend-mode:screen
# sobre fundo Verde Floresta para eliminar o preto.
_logo_path = os.path.join(os.path.dirname(__file__), "logo_base64.txt")
if os.path.exists(_logo_path):
    with open(_logo_path, "r") as _f:
        LOGO_DATA_URI = _f.read().strip()
else:
    LOGO_DATA_URI = ""


# ── 1. Buscar livros da campanha ───────────────────────────────────


def _headers(token: str, extra: dict | None = None) -> dict:
    h = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    if extra:
        h.update(extra)
    return h


def buscar_livros_campanha(token: str) -> tuple[list[dict], int]:
    """Retorna (livros_pendentes, total_campanha).

    Livros pendentes: Created >= DATA_CORTE e CampanhaGradual != true.
    Total campanha: todos os livros Created >= DATA_CORTE (para cálculo do lote).
    """
    url = f"{GRAPH_URL}/sites/{SP_SITE_ID}/lists/{SP_LIST_ID}/items"
    campos = (
        "Id,_x0023_,Title,Autor_x002f_autores,Editora,"
        "Anodepublica_x00e7__x00e3_o,Link,CampanhaGradual,Created"
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

    # Filtrar localmente: Created >= DATA_CORTE
    corte = datetime.fromisoformat(f"{DATA_CORTE}T00:00:00Z")
    campanha = []
    for item in all_items:
        fields = item.get("fields", {})
        created_str = fields.get("Created", "")
        if not created_str:
            continue
        try:
            created = datetime.fromisoformat(created_str.replace("Z", "+00:00"))
        except ValueError:
            continue
        if created >= corte.replace(tzinfo=created.tzinfo):
            campanha.append(item)

    total_campanha = len(campanha)

    pendentes = []
    for item in campanha:
        fields = item.get("fields", {})
        if fields.get("CampanhaGradual"):
            continue
        # Extrair URL do campo Link (pode ser string ou dict)
        link_field = fields.get("Link", "")
        if isinstance(link_field, dict):
            link_url = link_field.get("Url", "")
        elif isinstance(link_field, str) and link_field:
            link_url = link_field.split(", ")[0].strip()
        else:
            link_url = ""

        ano_raw = fields.get("Anodepublica_x00e7__x00e3_o", "")
        ano = str(int(ano_raw)) if ano_raw else ""

        pendentes.append({
            "Id": item["id"],
            "NumeroLista": int(fields.get("_x0023_", 0)),
            "Title": fields.get("Title", ""),
            "Autor": fields.get("Autor_x002f_autores", ""),
            "Editora": fields.get("Editora", ""),
            "Ano": ano,
            "Link": link_url,
        })

    # Ordenar: mais antigos primeiro (menor número = mais antigo)
    pendentes.sort(key=lambda x: x["NumeroLista"])

    return pendentes, total_campanha


# ── 2. Calcular tamanho do lote ────────────────────────────────────


def calcular_lote(total_campanha: int) -> int:
    """Retorna o tamanho do lote: min(ceil(total × 5%), 100)."""
    if total_campanha == 0:
        return 0
    return min(math.ceil(total_campanha * PERCENTUAL_LOTE), TETO_LOTE)


# ── 3. Montar HTML do e-mail ──────────────────────────────────────


def _linha_livro(livro: dict) -> str:
    """Gera uma linha <tr> da tabela no padrão do e-mail semanal."""
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


def montar_html(lote: list[dict], pendentes_restantes: int, total_campanha: int) -> str:
    """Gera HTML completo do e-mail no padrão visual do e-mail semanal LDCM."""
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
        # Header
        '<tr><td style="background:#3D5549;padding:36px 40px 32px;text-align:center;">'
        f'<img alt="LDCM Advogados" width="160" style="display:block;margin:0 auto;mix-blend-mode:screen;max-width:160px;" src="{LOGO_DATA_URI}">'
        '</td></tr>'
        # Body
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
        f'<p style="color:#77787B;font-size:12px;margin-top:16px;">'
        f'Restam <strong>{pendentes_restantes}</strong> t&#237;tulos a serem notificados '
        f'(de {total_campanha} adicionados desde {DATA_CORTE}).</p>'
        '<p style="margin-top:24px;margin-bottom:0;color:#3D5549;font-weight:600;">Comit&#234; de TI e Biblioteca</p>'
        '<p style="margin-top:2px;color:#77787B;font-size:13px;">LDCM Advogados</p>'
        '</td></tr>'
        # Footer
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
    """Converte EMAIL_DESTINATARIOS (separados por ;) em lista Graph API."""
    return [
        {"emailAddress": {"address": e.strip()}}
        for e in EMAIL_DESTINATARIOS.split(";")
        if e.strip()
    ]


@retry(stop=stop_after_attempt(3), wait=wait_exponential(min=2, max=30))
def enviar_email(token: str, html: str, qtd_livros: int) -> int:
    """Envia e-mail via POST /users/{remetente}/sendMail. Retorna status code."""
    destinatarios = _parse_destinatarios()
    if not destinatarios:
        raise ValueError("EMAIL_DESTINATARIOS não configurado no .env")

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
            f"Falha ao enviar e-mail: {resp.status_code} — {resp.text}", response=resp
        )
    return resp.status_code


# ── 5. Marcar itens como enviados ──────────────────────────────────


@retry(stop=stop_after_attempt(3), wait=wait_exponential(min=2, max=30))
def marcar_enviado(token: str, item_id: str) -> int:
    """Marca CampanhaGradual = true para um item."""
    url = (
        f"{GRAPH_URL}/sites/{SP_SITE_ID}"
        f"/lists/{SP_LIST_ID}/items/{item_id}/fields"
    )
    resp = requests.patch(
        url,
        headers=_headers(token, {"Content-Type": "application/json"}),
        json={"CampanhaGradual": True},
        timeout=60,
    )
    if not resp.ok:
        raise requests.HTTPError(
            f"Falha ao marcar item {item_id}: {resp.status_code} — {resp.text}",
            response=resp,
        )
    return resp.status_code


def marcar_lote(token: str, lote: list[dict]) -> tuple[int, int]:
    """Marca todos os itens do lote. Retorna (marcados, falhas)."""
    marcados = 0
    falhas = 0
    for livro in lote:
        try:
            marcar_enviado(token, livro["Id"])
            marcados += 1
        except Exception as e:
            falhas += 1
            logger.error("Falha ao marcar item %s (%s): %s", livro["Id"], livro["Title"], e)
    return marcados, falhas
