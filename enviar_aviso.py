"""E-mail de aviso + primeiro lote semanal (Gustavo Alo) — 06/04/2026.

Uso:
    python enviar_aviso.py                # dry run, so para mim
    python enviar_aviso.py --enviar       # envia so para mim
    python enviar_aviso.py --enviar --todos  # envia para todos@ldcm.com.br
"""

import os
import sys

from dotenv import load_dotenv

import auth
import email_comum

load_dotenv()

CRIADOR_GUSTAVO = "Gustavo Aló (GAA)"


def montar_html(lote):
    logo = email_comum.LOGO_DATA_URI
    linhas = "\n".join(email_comum._linha_livro(livro) for livro in lote)
    qtd = len(lote)

    return (
        '<table width="100%" cellpadding="0" cellspacing="0" style="font-family:Calibri,Arial,sans-serif;background:#f4f3ef;">'
        '<tr><td align="center" style="padding:32px 16px 0;">'
        '<table width="600" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:2px;overflow:hidden;">'
        '<tr><td style="background:#3D5549;padding:36px 40px 32px;text-align:center;">'
        f'<img alt="LDCM Advogados" width="160" style="display:block;margin:0 auto;mix-blend-mode:screen;max-width:160px;" src="{logo}">'
        '</td></tr>'
        '<tr><td style="padding:36px 40px;color:#2a2a2a;font-size:15px;line-height:1.7;">'
        '<h2 style="color:#3D5549;border-bottom:2px solid #ADA688;padding-bottom:8px;margin-top:0;">Biblioteca LDCM &#8212; Aviso importante</h2>'
        '<p>Prezados,</p>'
        '<p>A biblioteca est&#225; passando por um processo de aperfei&#231;oamento que, dentre outras melhorias, incluir&#225; um aumento substancial do nosso acervo.</p>'
        '<p>A partir de agora, <strong>toda segunda-feira</strong> voc&#234;s receber&#227;o um e-mail com os novos t&#237;tulos inclu&#237;dos pela equipe na semana.</p>'
        f'<p>Abaixo est&#227;o os primeiros <strong>{qtd} t&#237;tulos</strong> adicionados desde 30.03:</p>'
        '<table style="border-collapse:collapse;width:100%;font-size:13px;">'
        '<tbody>'
        '<tr style="background-color:#3D5549;color:#ffffff;">'
        f'<th style="{email_comum.TH_STYLE}">T&#237;tulo</th>'
        f'<th style="{email_comum.TH_STYLE}">Autores</th>'
        f'<th style="{email_comum.TH_STYLE}">Editora</th>'
        f'<th style="{email_comum.TH_STYLE}">Ano</th>'
        f'<th style="{email_comum.TH_STYLE}">Link</th>'
        '</tr>'
        f'{linhas}'
        '</tbody>'
        '</table>'
        '<div style="margin-top:28px;padding:20px;background:#f4f3ef;border-radius:4px;">'
        '<p style="margin:0 0 8px;color:#2a2a2a;font-size:14px;line-height:1.6;font-weight:600;">E-mail di&#225;rio (opcional)</p>'
        '<p style="margin:0 0 12px;color:#2a2a2a;font-size:14px;line-height:1.6;">'
        'Al&#233;m das inclus&#245;es da equipe, <strong>milhares de novos t&#237;tulos</strong> est&#227;o sendo catalogados '
        'automaticamente no acervo. Muitos s&#227;o de autores estrangeiros e alguns s&#227;o antigos. '
        'Se quiser, voc&#234; pode receber um <strong>e-mail di&#225;rio</strong> com essas inclus&#245;es enquanto a campanha durar.</p>'
        f'<a href="{email_comum.MAILTO_INSCREVER}" style="{email_comum.BTN_STYLE}background:#3D5549;color:#ffffff;">'
        'Quero receber o e-mail di&#225;rio</a>'
        '</div>'
        '<p>Obrigado.</p>'
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


def enviar(dry_run=True, destinatario="rodrigo.moreira@ldcm.com.br"):
    token = auth.get_token()

    livros = email_comum.buscar_livros(token, criador=CRIADOR_GUSTAVO)
    print(f"Livros do Gustavo desde {email_comum.DATA_CORTE}: {len(livros)}")

    lote = livros[:100]
    html = montar_html(lote)

    if dry_run:
        os.makedirs("logs", exist_ok=True)
        with open("logs/aviso_dry.html", "w", encoding="utf-8") as f:
            f.write(html)
        print(f"DRY RUN -- HTML salvo em logs/aviso_dry.html ({len(html)} chars)")
        return

    email_comum.enviar_email(
        token,
        [destinatario],
        "Biblioteca LDCM \u2014 Aviso importante sobre atualiza\u00e7\u00f5es do acervo",
        html,
    )
    print(f"E-mail enviado para {destinatario}")

    # Salvar offset semanal
    offset_path = os.path.join(os.path.dirname(__file__), "offset_semanal.txt")
    with open(offset_path, "w") as f:
        f.write(str(len(lote)))
    print(f"Offset semanal salvo: {len(lote)}")


if __name__ == "__main__":
    is_enviar = "--enviar" in sys.argv
    dest = "todos@ldcm.com.br" if "--todos" in sys.argv else "rodrigo.moreira@ldcm.com.br"
    enviar(dry_run=not is_enviar, destinatario=dest)
