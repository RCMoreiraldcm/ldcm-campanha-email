"""E-mail semanal — livros incluidos pelo Gustavo Alo → todos@ldcm.com.br"""

import logging
import os
import sys

from dotenv import load_dotenv

import auth
import email_comum

load_dotenv()

logger = logging.getLogger("biblioteca_ldcm")
logger.setLevel(logging.INFO)
fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
os.makedirs("logs", exist_ok=True)
fh = logging.FileHandler("logs/semanal.log", encoding="utf-8")
fh.setFormatter(fmt)
logger.addHandler(fh)
sh = logging.StreamHandler(sys.stdout)
sh.setFormatter(fmt)
logger.addHandler(sh)

CRIADOR = "Gustavo Aló (GAA)"
DESTINATARIO = "todos@ldcm.com.br"


def executar(dry_run: bool = True):
    logger.info("=" * 60)
    logger.info("E-mail semanal -- livros do Gustavo Alo")
    logger.info("=" * 60)

    token = auth.get_token()
    logger.info("Autenticacao OK")

    livros = email_comum.buscar_livros(token, criador=CRIADOR)
    logger.info("Livros do Gustavo desde %s: %d", email_comum.DATA_CORTE, len(livros))

    # Ler offset semanal
    offset_path = os.path.join(os.path.dirname(__file__), "offset_semanal.txt")
    offset = 0
    if os.path.exists(offset_path):
        offset = int(open(offset_path).read().strip())

    if offset >= len(livros):
        logger.info("Todos os livros do Gustavo ja foram enviados.")
        return

    lote = livros[offset:offset + 100]
    novo_offset = offset + len(lote)
    logger.info("Lote: %d titulos (posicoes %d a %d)", len(lote), offset + 1, novo_offset)

    for i, livro in enumerate(lote, 1):
        logger.info("  [%d/%d] #%05d -- %s", i, len(lote), livro["NumeroLista"], livro["Title"])

    # Botao para inscrever no diario
    botao = (
        '<div style="margin-top:24px;padding:20px;background:#f4f3ef;border-radius:4px;">'
        '<p style="margin:0 0 8px;color:#2a2a2a;font-size:14px;line-height:1.6;">'
        'Al&#233;m destes, <strong>milhares de novos t&#237;tulos</strong> est&#227;o sendo adicionados automaticamente ao acervo '
        '(muitos de autores estrangeiros e alguns antigos). '
        'Se quiser, voc&#234; pode receber um e-mail di&#225;rio com essas inclus&#245;es:</p>'
        f'<a href="{email_comum.URL_INSCREVER}" style="{email_comum.BTN_STYLE}background:#3D5549;color:#ffffff;">Quero receber o e-mail di&#225;rio</a>'
        '</div>'
    )

    html = email_comum.montar_html_tabela(
        lote,
        titulo_email=f"Biblioteca LDCM &#8212; {len(lote)} novos t&#237;tulos",
        texto_intro="Novos t&#237;tulos foram inclu&#237;dos no acervo pela equipe. Confira abaixo:",
        botao_html=botao,
    )
    logger.info("HTML gerado: %d caracteres", len(html))

    if dry_run:
        with open("logs/ultimo_semanal_dry.html", "w", encoding="utf-8") as f:
            f.write(html)
        logger.info("DRY RUN -- e-mail NAO enviado.")
        return

    email_comum.enviar_email(
        token,
        [DESTINATARIO],
        f"Biblioteca LDCM \u2014 {len(lote)} novos t\u00edtulos",
        html,
    )
    logger.info("E-mail enviado para %s", DESTINATARIO)

    with open(offset_path, "w") as f:
        f.write(str(novo_offset))
    logger.info("Offset semanal: %d -> %d", offset, novo_offset)


if __name__ == "__main__":
    executar(dry_run="--enviar" not in sys.argv)
