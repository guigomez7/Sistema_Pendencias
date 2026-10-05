# -*- coding: utf-8 -*-
"""
Sistema de Envio Automático de Pendências
==========================================

Sistema 100% local (sem site, sem banco de dados, sem hospedagem) que:
 1. Localiza a pasta "sistema_pendencias" na Área de Trabalho do Windows;
 2. Lê os arquivos Excel da pasta "pendencias" (padrão: "Pendências [Empresa] [Data].xlsx");
 3. Consulta o cadastro de empresas/e-mails em "empresas/Empresas.xlsx";
 4. Monta e envia um e-mail por empresa, com o Excel correspondente anexado;
 5. Move os arquivos processados para "enviados" (organizados por data) ou "erros";
 6. Mantém um histórico de envios em "logs" para evitar duplicidade.

Uso:
    python enviar_pendencias.py            -> pergunta se é modo teste ou envio real
    python enviar_pendencias.py --teste    -> força modo teste (não envia e-mails)
    python enviar_pendencias.py --enviar   -> força envio real (ainda pede confirmação)

Este arquivo normalmente é executado através do "enviar_pendencias.bat".
"""

import os
import re
import sys
import json
import ssl
import shutil
import smtplib
import time
import getpass
import hashlib
import unicodedata
from pathlib import Path
from datetime import datetime
from email.message import EmailMessage

try:
    import openpyxl
except ImportError:
    print("ERRO: a biblioteca 'openpyxl' não está instalada.")
    print("Rode novamente o arquivo enviar_pendencias.bat para instalar as dependências,")
    print("ou execute manualmente: pip install -r requirements.txt")
    input("\nPressione ENTER para sair...")
    sys.exit(1)


# ---------------------------------------------------------------------------
# Configuração de caminhos
# ---------------------------------------------------------------------------

PASTA_SCRIPT = Path(__file__).resolve().parent
ARQUIVO_CONFIG = PASTA_SCRIPT / "config.json"
ARQUIVO_CREDENCIAIS = PASTA_SCRIPT / "credenciais.json"

NOME_PASTA_SISTEMA = "sistema_pendencias"
REGEX_ARQUIVO_PENDENCIA = re.compile(
    r"^pend[eê]ncias\s+(.+?)\s+(\d{6,8})\.xlsx$", re.IGNORECASE
)
REGEX_EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# ---------------------------------------------------------------------------
# Utilitários gerais
# ---------------------------------------------------------------------------

def normalizar_texto(texto: str) -> str:
    """Remove acentos, espaços extras e deixa tudo em minúsculo, para comparação."""
    texto = texto.strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ASCII", "ignore").decode("ASCII")
    texto = re.sub(r"\s+", " ", texto)
    return texto.lower().strip()


def periodo_do_dia() -> str:
    """Retorna 'manha' ou 'tarde' com base na hora atual do computador."""
    hora_atual = datetime.now().hour
    return "manha" if hora_atual < 12 else "tarde"


def formatar_data_exibicao(data_str: str) -> str:
    """Converte DDMMYYYY em DD/MM/YYYY para exibição/assunto do e-mail."""
    try:
        data = datetime.strptime(data_str, "%d%m%Y")
        return data.strftime("%d/%m/%Y")
    except ValueError:
        return data_str


def pausar_saida(codigo: int = 0):
    input("\nPressione ENTER para sair...")
    sys.exit(codigo)


# ---------------------------------------------------------------------------
# Localização automática da pasta base na Área de Trabalho
# ---------------------------------------------------------------------------

def localizar_pasta_base(config: dict) -> Path:
    """
    Procura a pasta 'sistema_pendencias' na Área de Trabalho do usuário do Windows.
    Considera variações comuns: Desktop / Área de Trabalho, com ou sem OneDrive.
    Pode ser sobrescrito via config.json -> "pasta_base_personalizada".
    """
    caminho_personalizado = config.get("pasta_base_personalizada", "").strip()
    if caminho_personalizado:
        pasta = Path(caminho_personalizado)
        if pasta.exists():
            return pasta
        print(f"AVISO: 'pasta_base_personalizada' definida no config.json não foi encontrada: {pasta}")
        print("Tentando localizar automaticamente na Área de Trabalho...\n")

    home = Path.home()
    candidatos = [
        home / "Desktop" / NOME_PASTA_SISTEMA,
        home / "OneDrive" / "Desktop" / NOME_PASTA_SISTEMA,
        home / "Área de Trabalho" / NOME_PASTA_SISTEMA,
        home / "OneDrive" / "Área de Trabalho" / NOME_PASTA_SISTEMA,
    ]

    # Também tenta descobrir Desktops "OneDrive - <Empresa>" comuns em ambientes corporativos
    if home.exists():
        for item in home.iterdir():
            if item.is_dir() and item.name.lower().startswith("onedrive"):
                candidatos.append(item / "Desktop" / NOME_PASTA_SISTEMA)
                candidatos.append(item / "Área de Trabalho" / NOME_PASTA_SISTEMA)

    for candidato in candidatos:
        if candidato.exists() and candidato.is_dir():
            return candidato

    print("ERRO: não foi possível localizar a pasta 'sistema_pendencias' na Área de Trabalho.")
    print("Caminhos verificados:")
    for c in candidatos:
        print(f"  - {c}")
    print("\nCrie a pasta 'sistema_pendencias' na Área de Trabalho (com as subpastas")
    print("'pendencias' e 'empresas') ou informe o caminho manualmente em config.json")
    print("no campo 'pasta_base_personalizada'.")
    pausar_saida(1)


def preparar_pastas(pasta_base: Path) -> dict:
    """Garante que todas as subpastas necessárias existem e retorna seus caminhos."""
    pastas = {
        "pendencias": pasta_base / "pendencias",
        "empresas": pasta_base / "empresas",
        "enviados": pasta_base / "enviados",
        "erros": pasta_base / "erros",
        "logs": pasta_base / "logs",
    }
    faltando = [nome for nome, caminho in pastas.items()
                if nome in ("pendencias", "empresas") and not caminho.exists()]
    if faltando:
        print("ERRO: as seguintes pastas obrigatórias não foram encontradas dentro de "
              f"'{pasta_base}':")
        for nome in faltando:
            print(f"  - {nome}")
        pausar_saida(1)

    for nome in ("enviados", "erros", "logs"):
        pastas[nome].mkdir(parents=True, exist_ok=True)

    return pastas


# ---------------------------------------------------------------------------
# Configuração e credenciais
# ---------------------------------------------------------------------------

def carregar_config() -> dict:
    if not ARQUIVO_CONFIG.exists():
        print(f"ERRO: arquivo de configuração não encontrado: {ARQUIVO_CONFIG}")
        pausar_saida(1)
    try:
        with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as erro:
        print(f"ERRO: config.json está com formatação inválida ({erro}).")
        pausar_saida(1)


def carregar_credenciais() -> dict:
    """
    Carrega a senha do e-mail a partir de um arquivo separado (credenciais.json),
    nunca do código-fonte. Se o arquivo não existir ou a senha estiver vazia,
    pede a senha ao usuário (sem exibir na tela) e salva para as próximas execuções.
    """
    dados = {"senha": ""}
    if ARQUIVO_CREDENCIAIS.exists():
        try:
            with open(ARQUIVO_CREDENCIAIS, "r", encoding="utf-8") as f:
                dados = json.load(f)
        except json.JSONDecodeError:
            print("AVISO: credenciais.json inválido, será recriado.")
            dados = {"senha": ""}

    if not dados.get("senha"):
        print("Nenhuma senha de e-mail configurada ainda.")
        senha = getpass.getpass("Digite a senha do e-mail de envio (não será exibida): ").strip()
        dados["senha"] = senha
        with open(ARQUIVO_CREDENCIAIS, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
        print(f"Senha salva em '{ARQUIVO_CREDENCIAIS.name}' (mantenha este arquivo protegido "
              "e fora de qualquer repositório/nuvem pública).\n")

    return dados


# ---------------------------------------------------------------------------
# Cadastro de empresas
# ---------------------------------------------------------------------------

def carregar_cadastro_empresas(pasta_empresas: Path) -> dict:
    """
    Lê o(s) arquivo(s) .xlsx da pasta 'empresas'. Espera colunas:
        A = Nome da empresa | B = E-mail
    Retorna dicionário: nome_normalizado -> {"nome_original": ..., "email": ...}
    """
    arquivos_excel = sorted(
        [p for p in pasta_empresas.glob("*.xlsx") if not p.name.startswith("~$")]
    )
    if not arquivos_excel:
        print(f"ERRO: nenhum arquivo Excel de cadastro encontrado em '{pasta_empresas}'.")
        pausar_saida(1)

    cadastro = {}
    for arquivo in arquivos_excel:
        try:
            wb = openpyxl.load_workbook(arquivo, data_only=True)
            ws = wb.active
        except Exception as erro:
            print(f"AVISO: não foi possível abrir o cadastro '{arquivo.name}': {erro}")
            continue

        for linha in ws.iter_rows(min_row=1, values_only=True):
            if not linha or linha[0] is None:
                continue
            nome_bruto = str(linha[0]).strip()
            email_bruto = str(linha[1]).strip() if len(linha) > 1 and linha[1] else ""

            if nome_bruto.lower() in ("empresa", "nome", "nome da empresa"):
                continue  # pula cabeçalho

            chave = normalizar_texto(nome_bruto)
            if chave:
                cadastro[chave] = {"nome_original": nome_bruto, "email": email_bruto}

    return cadastro


# ---------------------------------------------------------------------------
# Interpretação dos arquivos de pendências
# ---------------------------------------------------------------------------

def interpretar_nome_arquivo(nome_arquivo: str):
    """
    Extrai (empresa, data_str) do nome do arquivo no padrão:
        'Pendências [Empresa] [DDMMAAAA].xlsx'
    Retorna (empresa, data_str, erro). 'erro' é None se tudo estiver correto.
    """
    m = REGEX_ARQUIVO_PENDENCIA.match(nome_arquivo.strip())
    if not m:
        return None, None, "Nome de arquivo fora do padrão 'Pendências [Empresa] [Data].xlsx'"

    empresa_bruta = m.group(1).strip()
    data_str = m.group(2).strip()

    if len(data_str) != 8:
        return empresa_bruta, data_str, f"Data inválida no nome do arquivo ('{data_str}' deveria ter 8 dígitos, DDMMAAAA)"

    try:
        datetime.strptime(data_str, "%d%m%Y")
    except ValueError:
        return empresa_bruta, data_str, f"Data inválida no nome do arquivo ('{data_str}' não é uma data real)"

    return empresa_bruta, data_str, None


def validar_email(email: str) -> bool:
    return bool(email) and bool(REGEX_EMAIL.match(email))


# ---------------------------------------------------------------------------
# Histórico de envios (evita duplicidade)
# ---------------------------------------------------------------------------

def caminho_historico(pasta_logs: Path) -> Path:
    return pasta_logs / "historico_envios.json"


def carregar_historico(pasta_logs: Path) -> dict:
    caminho = caminho_historico(pasta_logs)
    if caminho.exists():
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            print("AVISO: histórico de envios corrompido, um novo será iniciado.")
    return {}


def salvar_historico(pasta_logs: Path, historico: dict):
    caminho = caminho_historico(pasta_logs)
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(historico, f, ensure_ascii=False, indent=2)


def chave_historico(empresa_normalizada: str, data_str: str) -> str:
    return f"{empresa_normalizada}_{data_str}"


def registrar_log_execucao(pasta_logs: Path, linhas: list):
    """Grava um arquivo de log detalhado desta execução (texto simples)."""
    nome_arquivo = f"log_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt"
    caminho = pasta_logs / nome_arquivo
    with open(caminho, "w", encoding="utf-8") as f:
        f.write("\n".join(linhas))
    return caminho


# ---------------------------------------------------------------------------
# Processamento / classificação dos arquivos
# ---------------------------------------------------------------------------

def processar_pendencias(pasta_pendencias: Path, cadastro: dict, historico: dict) -> list:
    """
    Retorna uma lista de dicionários, um por arquivo, com todas as informações
    necessárias para exibir o resumo e, depois, enviar o e-mail.
    """
    resultado = []
    arquivos = sorted(
        [p for p in pasta_pendencias.glob("*.xlsx") if not p.name.startswith("~$")]
    )

    if not arquivos:
        return resultado

    for arquivo in arquivos:
        item = {
            "arquivo": arquivo,
            "nome_arquivo": arquivo.name,
            "empresa": None,
            "email": None,
            "data_str": None,
            "data_exibicao": None,
            "status": "ok",       # ok | erro | duplicado
            "motivo_erro": None,
        }

        empresa_bruta, data_str, erro_nome = interpretar_nome_arquivo(arquivo.name)
        if erro_nome:
            item["status"] = "erro"
            item["motivo_erro"] = erro_nome
            item["empresa"] = empresa_bruta
            resultado.append(item)
            continue

        item["data_str"] = data_str
        item["data_exibicao"] = formatar_data_exibicao(data_str)

        chave_empresa = normalizar_texto(empresa_bruta)
        cadastro_empresa = cadastro.get(chave_empresa)

        if not cadastro_empresa:
            item["status"] = "erro"
            item["empresa"] = empresa_bruta
            item["motivo_erro"] = "Empresa não cadastrada em Empresas.xlsx"
            resultado.append(item)
            continue

        nome_oficial = cadastro_empresa["nome_original"]
        email = cadastro_empresa["email"]
        item["empresa"] = nome_oficial

        if not validar_email(email):
            item["status"] = "erro"
            item["motivo_erro"] = f"E-mail ausente ou inválido no cadastro ('{email}')"
            resultado.append(item)
            continue

        item["email"] = email

        # Verifica duplicidade (mesma empresa + mesma data já enviada com sucesso)
        chave_hist = chave_historico(chave_empresa, data_str)
        if chave_hist in historico and historico[chave_hist].get("status") == "enviado":
            item["status"] = "duplicado"
            item["motivo_erro"] = (
                f"Já enviado anteriormente em {historico[chave_hist].get('enviado_em', '?')}"
            )
            resultado.append(item)
            continue

        resultado.append(item)

    return resultado


# ---------------------------------------------------------------------------
# Exibição do resumo e confirmação do usuário
# ---------------------------------------------------------------------------

def exibir_resumo(itens: list, modo_teste: bool):
    total = len(itens)
    ok = [i for i in itens if i["status"] == "ok"]
    erros = [i for i in itens if i["status"] == "erro"]
    duplicados = [i for i in itens if i["status"] == "duplicado"]

    print("=" * 78)
    print(f"RESUMO DO PROCESSAMENTO  —  {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    if modo_teste:
        print(">>> MODO TESTE ATIVADO: nenhum e-mail real será enviado <<<")
    print("=" * 78)
    print(f"Total de arquivos encontrados: {total}")
    print(f"  Prontos para envio.........: {len(ok)}")
    print(f"  Com erro (não serão enviados): {len(erros)}")
    print(f"  Duplicados (já enviados antes): {len(duplicados)}")
    print("-" * 78)

    if ok:
        print("\nARQUIVOS PRONTOS PARA ENVIO:")
        for i, item in enumerate(ok, 1):
            print(f"  [{i}] {item['nome_arquivo']}")
            print(f"       Empresa: {item['empresa']}   |   E-mail: {item['email']}   |   Data: {item['data_exibicao']}")

    if duplicados:
        print("\nARQUIVOS IGNORADOS POR DUPLICIDADE:")
        for i, item in enumerate(duplicados, 1):
            print(f"  [{i}] {item['nome_arquivo']} -> {item['motivo_erro']}")

    if erros:
        print("\nARQUIVOS COM ERRO (não serão enviados):")
        for i, item in enumerate(erros, 1):
            print(f"  [{i}] {item['nome_arquivo']} -> {item['motivo_erro']}")

    print("=" * 78 + "\n")
    return ok, erros, duplicados


def perguntar_sim_nao(pergunta: str) -> bool:
    while True:
        resposta = input(f"{pergunta} (S/N): ").strip().lower()
        if resposta in ("s", "sim"):
            return True
        if resposta in ("n", "nao", "não"):
            return False
        print("Resposta inválida. Digite S ou N.")


# ---------------------------------------------------------------------------
# Montagem e envio de e-mail
# ---------------------------------------------------------------------------

def montar_corpo_email() -> str:
    if periodo_do_dia() == "manha":
        return "Bom dia,\n\nSeguem as pendências atualizadas abaixo."
    return "Boa tarde,\n\nSeguem as pendências atualizadas abaixo."


def montar_email(item: dict, config: dict) -> EmailMessage:
    smtp_cfg = config["smtp"]
    assunto = f"Pendências {item['empresa']} - {item['data_exibicao']}"

    msg = EmailMessage()
    msg["From"] = f"{smtp_cfg.get('remetente_nome', '')} <{smtp_cfg['remetente_email']}>"
    msg["To"] = item["email"]
    msg["Subject"] = assunto
    msg.set_content(montar_corpo_email())

    with open(item["arquivo"], "rb") as f:
        dados_anexo = f.read()

    # EmailMessage.add_attachment cuida sozinho da codificação correta do nome
    # do arquivo (inclusive com acentos, ex: "Pendências ...xlsx") e usa o
    # tipo MIME correto do Excel, para que o anexo apareça como planilha
    # (e não como um arquivo genérico) na caixa de entrada do destinatário.
    msg.add_attachment(
        dados_anexo,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=item["nome_arquivo"],
    )

    return msg


def abrir_conexao_smtp(config: dict, senha: str) -> smtplib.SMTP:
    """
    Abre e autentica uma única conexão SMTP, reaproveitada para enviar todos
    os e-mails da execução (em vez de abrir uma conexão nova por e-mail).
    """
    smtp_cfg = config["smtp"]
    host = smtp_cfg["host"]
    porta = int(smtp_cfg["porta"])
    remetente = smtp_cfg["remetente_email"]
    usar_tls = smtp_cfg.get("usar_tls", True)

    contexto = ssl.create_default_context()

    if porta == 465:
        servidor = smtplib.SMTP_SSL(host, porta, context=contexto, timeout=30)
    else:
        servidor = smtplib.SMTP(host, porta, timeout=30)
        servidor.ehlo()
        if usar_tls:
            servidor.starttls(context=contexto)
            servidor.ehlo()

    servidor.login(remetente, senha)
    return servidor


def enviar_email_smtp(servidor: smtplib.SMTP, msg: EmailMessage):
    servidor.send_message(msg)


# ---------------------------------------------------------------------------
# Movimentação de arquivos processados
# ---------------------------------------------------------------------------

def mover_para_enviados(item: dict, pasta_enviados: Path):
    subpasta = pasta_enviados / item["data_exibicao"].replace("/", "-")
    subpasta.mkdir(parents=True, exist_ok=True)
    destino = subpasta / item["nome_arquivo"]
    destino = caminho_sem_conflito(destino)
    shutil.move(str(item["arquivo"]), str(destino))


def mover_para_erros(item: dict, pasta_erros: Path):
    destino = pasta_erros / item["nome_arquivo"]
    destino = caminho_sem_conflito(destino)
    shutil.move(str(item["arquivo"]), str(destino))


def caminho_sem_conflito(destino: Path) -> Path:
    """Se já existir um arquivo com o mesmo nome no destino, adiciona um sufixo."""
    if not destino.exists():
        return destino
    contador = 1
    while True:
        novo = destino.with_name(f"{destino.stem}_{contador}{destino.suffix}")
        if not novo.exists():
            return novo
        contador += 1


# ---------------------------------------------------------------------------
# Fluxo principal
# ---------------------------------------------------------------------------

def main():
    print("\nSistema de Envio Automático de Pendências\n" + "-" * 42)

    config = carregar_config()
    pasta_base = localizar_pasta_base(config)
    print(f"Pasta do sistema localizada em: {pasta_base}\n")

    pastas = preparar_pastas(pasta_base)

    cadastro = carregar_cadastro_empresas(pastas["empresas"])
    print(f"Cadastro de empresas carregado: {len(cadastro)} empresa(s).")

    historico = carregar_historico(pastas["logs"])

    itens = processar_pendencias(pastas["pendencias"], cadastro, historico)
    if not itens:
        print("\nNenhum arquivo de pendência encontrado na pasta 'pendencias'.")
        pausar_saida(0)

    # Determina modo teste: argumento de linha de comando tem prioridade,
    # depois pergunta ao usuário, usando o config.json como sugestão padrão.
    modo_teste = None
    if "--teste" in sys.argv:
        modo_teste = True
    elif "--enviar" in sys.argv:
        modo_teste = False

    if modo_teste is None:
        sugestao = config.get("modo_teste_padrao", True)
        texto_sugestao = "S" if sugestao else "N"
        resposta = input(
            f"Executar em MODO TESTE (não envia e-mails reais)? [{texto_sugestao}] (S/N): "
        ).strip().lower()
        if resposta == "":
            modo_teste = sugestao
        else:
            modo_teste = resposta in ("s", "sim")

    ok, erros, duplicados = exibir_resumo(itens, modo_teste)

    for item in erros + duplicados:
        mover_para_erros(item, pastas["erros"])

    if not ok:
        print("Nenhum arquivo apto para envio. Encerrando.")
        pausar_saida(0)

    if not perguntar_sim_nao("Confirma o envio dos e-mails listados acima?"):
        print("Operação cancelada pelo usuário. Nenhum e-mail foi enviado.")
        pausar_saida(0)

    credenciais = carregar_credenciais() if not modo_teste else {"senha": ""}
    senha = credenciais.get("senha", "")

    atraso_entre_envios = float(config.get("atraso_entre_envios_segundos", 2))
    max_tentativas = int(config.get("max_tentativas_envio", 3))

    print()
    linhas_log = [f"Execução em {datetime.now().strftime('%d/%m/%Y %H:%M:%S')} "
                  f"({'MODO TESTE' if modo_teste else 'ENVIO REAL'})"]

    enviados_com_sucesso = 0
    falhas_envio = 0

    # Abre UMA única conexão SMTP e a reaproveita para todos os e-mails desta
    # execução. Abrir uma conexão/login novo para cada e-mail (um por vez, em
    # sequência rápida) é o que costuma disparar erros temporários do tipo
    # "Try again later" ou "Temporary authentication failure" em servidores
    # de hospedagem compartilhada, que limitam quantas conexões/logins são
    # aceitos em um curto intervalo de tempo.
    servidor = None
    if not modo_teste:
        try:
            servidor = abrir_conexao_smtp(config, senha)
        except Exception as erro_conexao:
            print(f"ERRO: não foi possível conectar/autenticar no servidor SMTP: {erro_conexao}")
            pausar_saida(1)

    try:
        for i, item in enumerate(ok, 1):
            prefixo = f"[{i}/{len(ok)}] {item['empresa']} <{item['email']}>"

            if modo_teste:
                print(f"{prefixo} ... SIMULADO (modo teste, nada foi enviado)")
                linhas_log.append(f"{item['nome_arquivo']} -> SIMULADO (modo teste)")
                continue

            sucesso = False
            ultimo_erro = None

            for tentativa in range(1, max_tentativas + 1):
                try:
                    msg = montar_email(item, config)
                    enviar_email_smtp(servidor, msg)
                    sucesso = True
                    break
                except (smtplib.SMTPServerDisconnected,
                        smtplib.SMTPResponseException,
                        smtplib.SMTPException,
                        ConnectionError, OSError) as erro:
                    ultimo_erro = erro
                    print(f"{prefixo} ... tentativa {tentativa}/{max_tentativas} falhou "
                          f"({erro}). Aguardando e tentando novamente...")
                    try:
                        servidor.quit()
                    except Exception:
                        pass
                    time.sleep(atraso_entre_envios * 3)
                    try:
                        servidor = abrir_conexao_smtp(config, senha)
                    except Exception as erro_reconexao:
                        ultimo_erro = erro_reconexao
                        time.sleep(atraso_entre_envios * 3)

            try:
                if sucesso:
                    chave_hist = chave_historico(normalizar_texto(item["empresa"]), item["data_str"])
                    historico[chave_hist] = {
                        "status": "enviado",
                        "empresa": item["empresa"],
                        "email": item["email"],
                        "arquivo": item["nome_arquivo"],
                        "enviado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    }
                    salvar_historico(pastas["logs"], historico)

                    mover_para_enviados(item, pastas["enviados"])

                    print(f"{prefixo} ... ENVIADO")
                    linhas_log.append(f"{item['nome_arquivo']} -> ENVIADO para {item['email']}")
                    enviados_com_sucesso += 1
                else:
                    raise ultimo_erro if ultimo_erro else RuntimeError("Falha desconhecida no envio")

            except Exception as erro:
                print(f"{prefixo} ... ERRO: {erro}")
                linhas_log.append(f"{item['nome_arquivo']} -> ERRO: {erro}")
                item["status"] = "erro"
                item["motivo_erro"] = str(erro)
                try:
                    mover_para_erros(item, pastas["erros"])
                except Exception as erro_mover:
                    linhas_log.append(f"  (falha ao mover arquivo para 'erros': {erro_mover})")
                falhas_envio += 1

            if not modo_teste:
                time.sleep(atraso_entre_envios)
    finally:
        if servidor is not None:
            try:
                servidor.quit()
            except Exception:
                pass

    caminho_log = registrar_log_execucao(pastas["logs"], linhas_log)

    print("\n" + "=" * 78)
    print("PROCESSAMENTO CONCLUÍDO")
    if modo_teste:
        print(f"Modo teste: {len(ok)} arquivo(s) simulados, nenhum e-mail real enviado.")
    else:
        print(f"E-mails enviados com sucesso: {enviados_com_sucesso}")
        print(f"Falhas no envio..............: {falhas_envio}")
    print(f"Arquivos com erro/duplicados movidos: {len(erros) + len(duplicados)}")
    print(f"Log detalhado salvo em: {caminho_log}")
    print("=" * 78)

    pausar_saida(0)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperação interrompida pelo usuário.")
        sys.exit(1)
    except Exception as erro_geral:
        print(f"\nERRO INESPERADO: {erro_geral}")
        pausar_saida(1)
