"""E-mail diario — livros da Aplicacao do SharePoint → inscritos SharePoint"""

import json
import logging
import os
import sys
from datetime import date

from dotenv import load_dotenv

import auth
import email_comum

load_dotenv()

SENT_LOG = os.path.join(os.path.dirname(__file__), "sent_today.json")

logger = logging.getLogger("biblioteca_ldcm")
logger.setLevel(logging.INFO)
fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
os.makedirs("logs", exist_ok=True)
fh = logging.FileHandler("logs/diario.log", encoding="utf-8")
fh.setFormatter(fmt)
logger.addHandler(fh)
sh = logging.StreamHandler(sys.stdout)
sh.setFormatter(fmt)
logger.addHandler(sh)

CRIADOR = "Aplicação do SharePoint"


def _ler_sent_log() -> dict:
    """Retorna {"date": "YYYY-MM-DD", "emails": [...], "offset": N}."""
    if os.path.exists(SENT_LOG):
        with open(SENT_LOG, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"date": "", "emails": [], "offset": 0}


def _salvar_sent_log(data: str, emails: list[str], offset: int):
    with open(SENT_LOG, "w", encoding="utf-8") as f:
        json.dump({"date": data, "emails": emails, "offset": offset}, f)


def _montar_email(lote):
    botao = (
        '<div style="margin-top:24px;text-align:center;">'
        f'<a href="{email_comum.URL_DESINSCREVER}" style="{email_comum.BTN_STYLE}background:transparent;color:#3D5549;border:1.5px solid #3D5549;">'
        'Quero deixar de receber este e-mail</a>'
        '</div>'
    )
    return email_comum.montar_html_tabela(
        lote,
        titulo_email=f"Biblioteca LDCM &#8212; {len(lote)} novos t&#237;tulos catalogados",
        texto_intro="Novos t&#237;tulos foram catalogados automaticamente e adicionados ao acervo. Confira abaixo:",
        botao_html=botao,
    )


def executar(dry_run: bool = True):
    logger.info("=" * 60)
    logger.info("E-mail diario -- livros automaticos")
    logger.info("=" * 60)

    token = auth.get_token()
    logger.info("Autenticacao OK")

    subscribers = email_comum.ler_subscribers(token)
    if not subscribers:
        logger.info("Nenhum inscrito. Nada a enviar.")
        return

    logger.info("Inscritos: %d (%s)", len(subscribers), ", ".join(subscribers))

    livros = email_comum.buscar_livros(token, criador=CRIADOR)
    logger.info("Livros automaticos desde %s: %d", email_comum.DATA_CORTE, len(livros))

    hoje = date.today().isoformat()
    log = _ler_sent_log()

    # Se o dia mudou, avança o offset e reseta o log
    if log["date"] != hoje:
        offset = email_comum.ler_offset()
        logger.info("Novo dia. Offset atual: %d", offset)
        ja_enviados = []
    else:
        offset = log["offset"]
        ja_enviados = [e.lower() for e in log["emails"]]
        logger.info("Mesmo dia. Offset do dia: %d, ja receberam: %d", offset, len(ja_enviados))

    if offset >= len(livros):
        logger.info("Todos os livros automaticos ja foram enviados.")
        return

    lote = livros[offset:offset + 100]
    novo_offset = offset + len(lote)
    logger.info("Lote: %d titulos (posicoes %d a %d de %d)", len(lote), offset + 1, novo_offset, len(livros))

    # Determinar quem precisa receber
    destinatarios = [e for e in subscribers if e.lower() not in ja_enviados]
    if not destinatarios:
        logger.info("Todos os inscritos ja receberam o lote de hoje.")
        return

    logger.info("Destinatarios desta execucao: %d (%s)", len(destinatarios), ", ".join(destinatarios))

    for i, livro in enumerate(lote, 1):
        logger.info("  [%d/%d] #%05d -- %s", i, len(lote), livro["NumeroLista"], livro["Title"])

    html = _montar_email(lote)
    logger.info("HTML gerado: %d caracteres", len(html))

    if dry_run:
        with open("logs/ultimo_diario_dry.html", "w", encoding="utf-8") as f:
            f.write(html)
        logger.info("DRY RUN -- e-mail NAO enviado.")
        return

    email_comum.enviar_email(
        token,
        destinatarios,
        f"Biblioteca LDCM \u2014 {len(lote)} novos t\u00edtulos catalogados",
        html,
    )
    logger.info("E-mail enviado para %d inscritos", len(destinatarios))

    # Atualizar log: acumular quem já recebeu hoje
    todos_enviados = list(set(ja_enviados + [e.lower() for e in destinatarios]))
    _salvar_sent_log(hoje, todos_enviados, novo_offset)

    # Persistir offset apenas na primeira execução do dia
    if log["date"] != hoje:
        email_comum.salvar_offset(novo_offset)
        logger.info("Offset: %d -> %d", offset, novo_offset)
    else:
        logger.info("Offset mantido (novos inscritos receberam o mesmo lote do dia)")


if __name__ == "__main__":
    executar(dry_run="--enviar" not in sys.argv)
