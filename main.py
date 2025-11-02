import pandas as pd
import webbrowser
import time
import re
import pyautogui
from urllib.parse import quote_plus
from datetime import datetime

# === CONFIGURAÇÕES ===
PLANILHA = "data/Disparo.xlsx"

MENSAGEM_PADRAO = (
    "Olá {saudacao}\n"
    "Meu nome é Paula Castro, sou assistente comercial do escritório Laécio Aguiar Advogados Associados.\n\n"
    "Identificamos que foi ajuizada uma ação trabalhista contra a empresa {empresa}, inscrita no CNPJ {cnpj}.\n"
    "A audiência já possui data marcada para {data_audiencia}.\n\n"
    "O processo foi movido por {reclamante}, e, se desejar, posso lhe encaminhar os detalhes completos da ação, "
    "como os pedidos, valores e documentos, sem qualquer custo ou compromisso.\n\n"
    "É importante agir com rapidez, pois em muitos casos a notificação judicial demora a ser recebida pela empresa, "
    "o que pode gerar prejuízos se a audiência ocorrer sem defesa constituída.\n\n"
    "Gostaria que eu lhe enviasse as informações completas sobre essa ação?"
)

TEMPO_ABRIR_CONVERSA = 12
DELAY_ENTRE_ENVIO = 8
# ======================


def ler_planilha(caminho: str) -> pd.DataFrame:
    """Lê e normaliza as colunas da planilha de contatos."""
    df = pd.read_excel(caminho)

    # Normaliza os nomes das colunas
    df.columns = (
        df.columns.str.upper()
        .str.normalize("NFKD")
        .str.encode("ascii", errors="ignore")
        .str.decode("utf-8")
        .str.strip()
    )

    mapeamento = {
        "EMPRESA": ["EMPRESA"],
        "SOCIO": ["SOCIO", "SOCIO"],
        "RECLAMANTE": ["RECLAMANTE"],
        "TELEFONE": ["TELEFONE", "FONE", "CELULAR"],
        "CNPJ": ["CNPJ"],
        "DATA DE AUDIENCIA": ["DATA DE AUDIENCIA", "DATA AUDIENCIA", "DATA AUDIENCIA"]
    }

    colunas_encontradas = {
        chave: next((opcao for opcao in opcoes if opcao in df.columns), None)
        for chave, opcoes in mapeamento.items()
    }

    faltando = [k for k, v in colunas_encontradas.items() if v is None]
    if faltando:
        raise ValueError(f"Colunas ausentes na planilha: {faltando}")

    df.rename(columns={v: k for k, v in colunas_encontradas.items() if v}, inplace=True)
    return df


def enviar_mensagens(df: pd.DataFrame) -> None:
    """Envia mensagens do WhatsApp Web baseadas nas informações da planilha."""
    total = len(df)
    print(f"🔹 Iniciando envio para {total} registros...\n")

    for i, linha in df.iterrows():
        empresa = str(linha["EMPRESA"]).strip()
        cnpj = str(linha["CNPJ"]).strip()
        reclamante = str(linha["RECLAMANTE"]).strip()
        socio = str(linha.get("SOCIO", "")).strip()
        data_audiencia = formatar_data(linha["DATA DE AUDIENCIA"])

        telefones = re.split(r"[;/]+", str(linha["TELEFONE"]).strip())
        telefones = [t.strip() for t in telefones if t.strip()]

        for telefone in telefones:
            saudacao = f"{socio}, tudo bem?" if socio else "tudo bem?"
            mensagem = MENSAGEM_PADRAO.format(
                saudacao=saudacao,
                empresa=empresa,
                cnpj=cnpj,
                reclamante=reclamante,
                socio=socio,
                data_audiencia=data_audiencia,
            )

            url = f"https://web.whatsapp.com/send?phone={telefone}&text={quote_plus(mensagem)}"
            webbrowser.open(url)
            print(f"[{i+1}/{total}] Enviando mensagem para {socio or 'Contato sem nome'} ({telefone})...")

            time.sleep(TEMPO_ABRIR_CONVERSA)
            pyautogui.press('enter')
            print(f"✅ Mensagem enviada para {telefone}")

            time.sleep(DELAY_ENTRE_ENVIO)
            pyautogui.hotkey('ctrl', 'w')

    print("\n✅ Todos os envios foram concluídos com sucesso!")


def formatar_data(data) -> str:
    """Formata data para DD-MM-YYYY."""
    if pd.isna(data):
        return ""
    try:
        return pd.to_datetime(data).strftime("%d-%m-%Y")
    except Exception:
        return str(data)


if __name__ == "__main__":
    df_contatos = ler_planilha(PLANILHA)
    enviar_mensagens(df_contatos)
