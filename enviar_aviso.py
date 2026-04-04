"""Envia o e-mail de aviso da campanha gradual para todos@ldcm.com.br.

Uso:
    python enviar_aviso.py              # envia imediatamente
    python enviar_aviso.py --dry-run    # mostra HTML sem enviar
"""

import os
import sys

import requests
from dotenv import load_dotenv

load_dotenv()

import auth

EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE", "rodrigo.moreira@ldcm.com.br")
GRAPH_URL = "https://graph.microsoft.com/v1.0"

# Logo base64
_logo_path = os.path.join(os.path.dirname(__file__), "logo_base64.txt")
with open(_logo_path, "r") as _f:
    LOGO_DATA_URI = _f.read().strip()


def montar_html() -> str:
    return (
        '<table width="100%" cellpadding="0" cellspacing="0" style="font-family: Calibri, Arial, sans-serif; background: #f4f3ef;">'
        "<tr>"
        '<td align="center" style="padding: 32px 16px 0;">'
        '<table width="600" cellpadding="0" cellspacing="0" style="background:#fff; border-radius:2px; overflow:hidden;">'
        "<tr>"
        '<td style="background:#3D5549; padding:36px 40px 32px; text-align:center;">'
        f'<img alt="LDCM Advogados" width="160" style="display:block;margin:0 auto;mix-blend-mode:screen;max-width:160px;" src="{LOGO_DATA_URI}">'
        "</td>"
        "</tr>"
        "<tr>"
        '<td style="padding:36px 40px; color:#2a2a2a; font-size:15px; line-height:1.7;">'
        '<h2 style="color:#3D5549; border-bottom:2px solid #ADA688; padding-bottom:8px; margin-top:0;">Biblioteca LDCM &#8212; Aviso importante</h2>'
        "<p>Prezados,</p>"
        "<p>A biblioteca est&#225; passando por um processo de aperfei&#231;oamento que, dentre outras melhorias, incluir&#225; um aumento substancial do nosso acervo.</p>"
        "<p>Para que n&#227;o recebam e-mails com uma lista impratic&#225;vel de inclus&#245;es, <strong>a partir de hoje voc&#234;s receber&#227;o, todo dia, um e-mail com a lista de parte dos t&#237;tulos que foram inclu&#237;dos</strong> a partir do dia 30.03.</p>"
        "<p>Voc&#234;s receber&#227;o este e-mail at&#233; que todas as inclus&#245;es tenham sido informadas, o que pode demorar. A partir da&#237;, voltar&#227;o os avisos semanais.</p>"
        "<p>O LDCM Advogados tem consci&#234;ncia da import&#226;ncia da biblioteca para o nosso fluxo de trabalho, n&#227;o apenas para pesquisas para a atua&#231;&#227;o em casos, mas tamb&#233;m para auxiliar nos estudos e trabalhos acad&#234;micos da equipe.</p>"
        "<p>O primeiro e-mail ser&#225; enviado na sequ&#234;ncia.</p>"
        "<p>Obrigado.</p>"
        '<p style="margin-top:24px; margin-bottom:0; color:#3D5549; font-weight:600;">Comit&#234; de TI e Biblioteca</p>'
        '<p style="margin-top:2px; color:#77787B; font-size:13px;">LDCM Advogados</p>'
        "</td>"
        "</tr>"
        "<tr>"
        '<td style="background:#3D5549; padding:20px 40px; text-align:center;">'
        '<span style="font-family:Arial,sans-serif; font-size:10px; color:#ADA688; letter-spacing:0.15em;">'
        "S&#195;O PAULO | RIO DE JANEIRO | LDCM.COM.BR"
        "</span>"
        "</td>"
        "</tr>"
        "</table>"
        "</td>"
        "</tr>"
        "</table>"
    )


def enviar():
    token = auth.get_token()
    html = montar_html()

    url = f"{GRAPH_URL}/users/{EMAIL_REMETENTE}/sendMail"
    body = {
        "message": {
            "subject": "Biblioteca LDCM \u2014 Aviso importante sobre atualiza\u00e7\u00f5es do acervo",
            "body": {"contentType": "HTML", "content": html},
            "from": {"emailAddress": {"address": "TI.biblioteca@ldcm.com.br"}},
            "toRecipients": [
                {"emailAddress": {"address": "todos@ldcm.com.br"}}
            ],
        },
        "saveToSentItems": True,
    }

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    resp = requests.post(url, headers=headers, json=body, timeout=60)
    if not resp.ok:
        print(f"ERRO: {resp.status_code} - {resp.text[:500]}")
        sys.exit(1)
    print("E-mail de aviso enviado com sucesso para todos@ldcm.com.br!")


if __name__ == "__main__":
    if "--dry-run" in sys.argv:
        print(montar_html())
    else:
        enviar()
