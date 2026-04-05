"""E-mail diario — livros da Aplicacao do SharePoint → subscribers.txt"""

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
fh = logging.FileHandler("logs/diario.log", encoding="utf-8")
fh.setFormatter(fmt)
logger.addHandler(fh)
sh = logging.StreamHandler(sys.stdout)
sh.setFormatter(fmt)
logger.addHandler(sh)

CRIADOR = "Aplicação do SharePoint"


def executar(dry_run: bool = True):
    logger.info("=" * 60)
    logger.info("E-mail diario -- livros automaticos")
    logger.info("=" * 60)

    token = auth.get_token()
    logger.info("Autenticacao OK")

    # Verificar subscribers (lista SharePoint)
    subscribers = email_comum.ler_subscribers(token)
    if not subscribers:
        logger.info("Nenhum inscrito. Nada a enviar.")
        return

    logger.info("Inscritos: %d (%s)", len(subscribers), ", ".join(subscribers))

    livros = email_comum.buscar_livros(token, criador=CRIADOR)
    logger.info("Livros automaticos desde %s: %d", email_comum.DATA_CORTE, len(livros))

    offset = email_comum.ler_offset()
    logger.info("Offset atual: %d", offset)

    if offset >= len(livros):
        logger.info("Todos os livros automaticos ja foram enviados.")
        return

    lote = livros[offset:offset + 100]
    novo_offset = offset + len(lote)
    logger.info("Lote: %d titulos (posicoes %d a %d de %d)", len(lote), offset + 1, novo_offset, len(livros))

    for i, livro in enumerate(lote, 1):
        logger.info("  [%d/%d] #%05d -- %s", i, len(lote), livro["NumeroLista"], livro["Title"])

    # Botao para desinscrever
    botao = (
        '<div style="margin-top:24px;text-align:center;">'
        f'<a href="{email_comum.URL_DESINSCREVER}" style="{email_comum.BTN_STYLE}background:transparent;color:#3D5549;border:1.5px solid #3D5549;">'
        'Quero deixar de receber este e-mail</a>'
        '</div>'
    )

    html = email_comum.montar_html_tabela(
        lote,
        titulo_email=f"Biblioteca LDCM &#8212; {len(lote)} novos t&#237;tulos catalogados",
        texto_intro="Novos t&#237;tulos foram catalogados automaticamente e adicionados ao acervo. Confira abaixo:",
        botao_html=botao,
    )
    logger.info("HTML gerado: %d caracteres", len(html))

    if dry_run:
        with open("logs/ultimo_diario_dry.html", "w", encoding="utf-8") as f:
            f.write(html)
        logger.info("DRY RUN -- e-mail NAO enviado.")
        return

    email_comum.enviar_email(
        token,
        subscribers,
        f"Biblioteca LDCM \u2014 {len(lote)} novos t\u00edtulos catalogados",
        html,
    )
    logger.info("E-mail enviado para %d inscritos", len(subscribers))

    email_comum.salvar_offset(novo_offset)
    logger.info("Offset: %d -> %d", offset, novo_offset)


if __name__ == "__main__":
    executar(dry_run="--enviar" not in sys.argv)
