#!/usr/bin/env python3
"""
Consulta a situação cadastral de CNPJs na Receita Federal, pelo terminal.

Exemplos:
  python3 consulta_cnpj_terminal.py base.txt
  python3 consulta_cnpj_terminal.py base.csv -o resultado.csv -r relatorio.txt
  python3 consulta_cnpj_terminal.py 33.000.167/0001-01 11.222.333/0001-81
  python3 consulta_cnpj_terminal.py base.txt --filtro ATIVA --intervalo 2
  python3 consulta_cnpj_terminal.py base.txt --sem-tabela -r relatorio.txt

Opções:
  entradas            arquivo .txt/.csv com CNPJs, e/ou CNPJs digitados direto
  -i, --intervalo     segundos entre consultas (padrão 1.5)
  -o, --csv           salva a tabela completa em CSV (CNPJ sem máscara, para abrir no Excel)
  -r, --relatorio     salva um relatório .txt com o resumo e a situação cadastral
  -f, --filtro        mostra na tabela só as linhas que contêm este texto
  --sem-tabela        não imprime a tabela no fim (só o resumo)

Ctrl+C interrompe a consulta e salva o que já foi consultado.
Usa só a biblioteca padrão do Python (não precisa de tkinter nem de outro arquivo).
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request
from datetime import datetime

BRASILAPI_URL = "https://brasilapi.com.br/api/cnpj/v1/{}"
RECEITAWS_URL = "https://www.receitaws.com.br/v1/cnpj/{}"
HEADERS = {"User-Agent": "consulta-cnpj-terminal/1.0", "Accept": "application/json"}

CAMPOS = [
    "cnpj", "status", "situacao", "data_situacao", "razao_social",
    "nome_fantasia", "municipio", "uf", "fonte", "erro",
]
PENDENTE, CONSULTADO, ERRO = "Pendente", "Consultado", "Erro"


# ---------- validação e formatação ----------

def limpar(texto: str) -> str:
    """Remove pontos, barras e traços. Aceita o formato alfanumérico (a partir de 2026)."""
    return re.sub(r"[^A-Za-z0-9]", "", texto).upper()


def cnpj_valido(cnpj: str) -> bool:
    c = limpar(cnpj)
    if len(c) != 14 or not c[:12].isalnum() or not c[12:].isdigit():
        return False
    if len(set(c)) == 1:
        return False
    valores = [ord(ch) - 48 for ch in c]  # dígitos valem 0-9; letras valem 17-42
    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    pesos2 = [6] + pesos1

    def digito(n, pesos):
        resto = 11 - (sum(v * p for v, p in zip(valores[:n], pesos)) % 11)
        return 0 if resto > 9 else resto

    return digito(12, pesos1) == valores[12] and digito(13, pesos2) == valores[13]


def mascara(cnpj: str) -> str:
    """Formata 00000000000000 como 00.000.000/0000-00. Se não tiver 14 caracteres, devolve como está."""
    c = limpar(cnpj)
    if len(c) != 14:
        return cnpj
    return f"{c[0:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}"


def motivo_invalido(bruto: str) -> str:
    """Explica em poucas palavras por que a entrada não é um CNPJ válido."""
    c = limpar(bruto)
    if not c:
        return "linha vazia"
    if c.isdigit() and len(c) == 11:
        return "tem 11 dígitos: parece CPF (este programa consulta só CNPJ)"
    if c.isdigit() and 12 <= len(c) <= 13:
        return f"tem {len(c)} dígitos; mesmo com zeros à esquerda não confere (CNPJ tem 14)"
    if c.isdigit() and len(c) < 14:
        return f"tem {len(c)} dígitos (CNPJ tem 14)"
    if len(c) != 14:
        return f"tem {len(c)} caracteres (CNPJ tem 14)"
    if len(set(c)) == 1:
        return "todos os dígitos são iguais"
    return "dígitos verificadores não conferem"


def candidato(token: str) -> tuple[str, bool] | None:
    """
    Decide se um pedaço de texto da base parece um CNPJ/CPF.
    Devolve (cnpj_normalizado, zeros_completados) ou None se não parece número de documento.
    Completa com zeros à esquerda quando o Excel removeu os zeros (12 ou 13 dígitos).
    """
    c = limpar(token)
    if c.isdigit() and 11 <= len(c) <= 14:
        if 12 <= len(c) <= 13:
            completado = c.zfill(14)
            return (completado, True) if cnpj_valido(completado) else (c, False)
        return c, False
    if len(c) == 14 and c.isalnum():
        return c, False
    return None


# ---------- consulta ----------

def requisitar(url: str, timeout: int = 15):
    req = urllib.request.Request(url, headers=HEADERS)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            corpo = json.loads(e.read().decode("utf-8"))
        except Exception:
            corpo = {"message": str(e.reason)}
        return e.code, corpo


def consultar_brasilapi(cnpj: str) -> dict:
    status, d = requisitar(BRASILAPI_URL.format(cnpj))
    if status != 200:
        raise RuntimeError(f"HTTP {status}: {d.get('message', d)}")
    return {
        "fonte": "BrasilAPI",
        "situacao": d.get("descricao_situacao_cadastral") or "",
        "data_situacao": d.get("data_situacao_cadastral") or "",
        "razao_social": d.get("razao_social") or "",
        "nome_fantasia": d.get("nome_fantasia") or "",
        "municipio": d.get("municipio") or "",
        "uf": d.get("uf") or "",
    }


def consultar_receitaws(cnpj: str) -> dict:
    status, d = requisitar(RECEITAWS_URL.format(cnpj))
    if status != 200 or d.get("status") == "ERROR":
        raise RuntimeError(d.get("message", f"HTTP {status}"))
    return {
        "fonte": "ReceitaWS",
        "situacao": d.get("situacao") or "",
        "data_situacao": d.get("data_situacao") or "",
        "razao_social": d.get("nome") or "",
        "nome_fantasia": d.get("fantasia") or "",
        "municipio": d.get("municipio") or "",
        "uf": d.get("uf") or "",
    }


def consultar(cnpj_bruto: str) -> dict:
    """Consulta um CNPJ. Tenta a BrasilAPI e, se falhar, a ReceitaWS."""
    cnpj = limpar(cnpj_bruto)
    linha = item_pendente(cnpj)
    if not cnpj_valido(cnpj):
        linha.update(status=ERRO, erro=motivo_invalido(cnpj))
        return linha
    erros = []
    for consulta in (consultar_brasilapi, consultar_receitaws):
        try:
            linha.update(consulta(cnpj))
            linha["status"] = CONSULTADO
            return linha
        except Exception as e:
            erros.append(f"{consulta.__name__}: {e}")
    linha.update(status=ERRO, erro=" | ".join(erros))
    return linha


# ---------- entrada ----------

def item_pendente(cnpj: str) -> dict:
    item = {campo: "" for campo in CAMPOS}
    item["cnpj"] = cnpj
    item["status"] = PENDENTE
    return item


def escolher_coluna(cabecalho: list, pedido: str | None) -> int | None:
    """Acha a coluna com os documentos: pelo nome informado ou por um cabeçalho com 'cnpj' ou 'cpf'."""
    nomes = [str(c or "").strip().lower() for c in cabecalho]
    if pedido:
        alvo = pedido.strip().lower()
        for i, nome in enumerate(nomes):
            if nome == alvo or alvo in nome:
                return i
        disponiveis = ", ".join(n for n in nomes if n)
        raise SystemExit(f"coluna '{pedido}' não encontrada. Colunas da planilha: {disponiveis}")
    for i, nome in enumerate(nomes):
        if "cnpj" in nome or "cpf" in nome:
            return i
    return None


def tokens_do_arquivo(caminho: str, coluna: str | None) -> list[str]:
    """Lê um .txt, .csv ou .xlsx e devolve os textos que podem conter CNPJ."""
    ext = os.path.splitext(caminho)[1].lower()
    if ext == ".xlsx":
        try:
            import openpyxl
        except ImportError:
            raise SystemExit("para ler .xlsx instale o openpyxl:  pip3 install openpyxl   (ou salve a planilha como CSV)")
        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
        tabelas = [list(ws.iter_rows(values_only=True)) for ws in wb.worksheets]
    elif ext == ".csv":
        with open(caminho, encoding="utf-8-sig", newline="") as f:
            amostra = f.read(4096)
            f.seek(0)
            try:
                separador = csv.Sniffer().sniff(amostra, delimiters=",;\t|").delimiter
            except csv.Error:
                separador = ";" if amostra.count(";") > amostra.count(",") else ","
            tabelas = [list(csv.reader(f, delimiter=separador))]
    else:
        with open(caminho, encoding="utf-8-sig") as f:
            return re.split(r"[;,\t|\s]+", f.read())

    saida: list[str] = []
    for linhas in tabelas:
        if not linhas:
            continue
        idx = escolher_coluna(list(linhas[0]), coluna)
        corpo = linhas[1:] if idx is not None else linhas
        for linha in corpo:
            celulas = [linha[idx]] if idx is not None and idx < len(linha) else (
                [] if idx is not None else list(linha)
            )
            for v in celulas:
                if v is None or isinstance(v, datetime):  # datas não são documentos
                    continue
                if isinstance(v, float) and v.is_integer():
                    v = int(v)
                saida.append(str(v))
    return saida


def montar_itens(entradas: list[str], coluna: str | None = None) -> list[dict]:
    """
    Monta a lista a partir de arquivos e/ou CNPJs digitados.
    Linhas que não são CNPJ válido entram com status Erro e o motivo, para você ver o que falhou.
    """
    itens, vistos = [], set()

    def adicionar(valor: str, nota: str = ""):
        if valor in vistos:
            return
        vistos.add(valor)
        item = item_pendente(valor)
        if not cnpj_valido(valor):
            item.update(status=ERRO, erro=motivo_invalido(valor))
        elif nota:
            item["erro"] = nota  # observação (ex.: zeros completados), a consulta ainda é feita
        itens.append(item)

    for entrada in entradas:
        if os.path.isfile(entrada):
            for token in tokens_do_arquivo(entrada, coluna):
                cand = candidato(token)
                if cand is None:
                    continue
                valor, completado = cand
                if valor in vistos:
                    continue
                if completado:
                    adicionar(valor, "zeros à esquerda completados (Excel pode ter removido)")
                else:
                    adicionar(valor)
        else:
            c = limpar(entrada)
            if c:
                adicionar(c)
    return itens


# ---------- resumo, tabela e arquivos ----------

def resumo(itens: list[dict]) -> dict:
    consultados = [i for i in itens if i["status"] == CONSULTADO]
    ativas = sum(1 for i in consultados if i["situacao"].upper() == "ATIVA")
    return {
        "total": len(itens),
        "pendentes": sum(1 for i in itens if i["status"] == PENDENTE),
        "consultados": len(consultados),
        "ativas": ativas,
        "nao_ativas": len(consultados) - ativas,
        "erros": sum(1 for i in itens if i["status"] == ERRO),
    }


def trunc(texto: str, n: int) -> str:
    return texto if len(texto) <= n else texto[: n - 1] + "…"


def imprimir_tabela(itens: list[dict], filtro: str | None) -> None:
    termo = (filtro or "").lower()
    linhas = [i for i in itens if not termo or termo in " ".join(str(v) for v in i.values()).lower()]
    cabecalho = f"{'CNPJ':<18}  {'SITUAÇÃO':<12}  {'DATA':<10}  {'UF':<2}  {'RAZÃO SOCIAL':<36}  OBSERVAÇÃO"
    print("\n" + cabecalho)
    print("-" * len(cabecalho))
    for i in linhas:
        if i["status"] == ERRO:
            obs = "ERRO: " + trunc(i["erro"], 60)
        elif i["status"] == CONSULTADO and i["erro"]:
            obs = trunc(i["erro"], 60)
        elif i["status"] == CONSULTADO:
            obs = i["fonte"]
        else:
            obs = i["status"]
        print(
            f"{mascara(i['cnpj']):<18}  {trunc(i['situacao'] or '-', 12):<12}  {i['data_situacao'] or '-':<10}  "
            f"{i['uf'] or '-':<2}  {trunc(i['razao_social'] or '-', 36):<36}  {obs}"
        )
    print(f"\n{len(linhas)} de {len(itens)} linhas exibidas")


def salvar_csv(itens: list[dict], caminho: str) -> None:
    with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS, extrasaction="ignore")
        w.writeheader()
        w.writerows(itens)


def gerar_relatorio_texto(itens: list[dict], base_nome: str, agora: datetime) -> str:
    r = resumo(itens)
    linhas = [
        "RELATÓRIO DE CONSULTA DE SITUAÇÃO CADASTRAL (RECEITA FEDERAL) - CNPJ",
        "=" * 72,
        f"Gerado em: {agora:%d/%m/%Y %H:%M}",
        f"Base: {base_nome}",
        "",
        "RESUMO",
        f"  Total na lista:           {r['total']}",
        f"  Consultados com sucesso:  {r['consultados']}",
        f"  Com erro:                 {r['erros']}",
        f"  Pendentes:                {r['pendentes']}",
        "",
        "SITUAÇÃO CADASTRAL",
    ]
    grupos: dict[str, list[dict]] = {}
    for item in itens:
        if item["status"] == CONSULTADO:
            grupos.setdefault(item["situacao"] or "(sem situação)", []).append(item)
    if not grupos:
        linhas.append("  Nenhuma consulta concluída.")
    for situacao in sorted(grupos):
        linhas.append(f"  {situacao}: {len(grupos[situacao])}")
    linhas.append("")
    for situacao in sorted(grupos):
        linhas.append(f"DETALHES - {situacao} ({len(grupos[situacao])})")
        for item in grupos[situacao]:
            linhas.append(
                f"  {mascara(item['cnpj'])} | {item['data_situacao'] or '-'} | "
                f"{item['uf'] or '-'} | {item['razao_social'] or '-'}"
            )
        linhas.append("")
    erros = [i for i in itens if i["status"] == ERRO]
    if erros:
        linhas.append(f"ERROS ({len(erros)})")
        for item in erros:
            linhas.append(f"  {item['cnpj']} | {item['erro']}")
        linhas.append("")
    pendentes = [i for i in itens if i["status"] == PENDENTE]
    if pendentes:
        linhas.append(f"PENDENTES ({len(pendentes)})")
        for item in pendentes:
            linhas.append(f"  {mascara(item['cnpj'])}")
        linhas.append("")
    return "\n".join(linhas) + "\n"


# ---------- programa principal ----------

def main() -> int:
    p = argparse.ArgumentParser(
        description="Consulta situação cadastral de CNPJs na Receita Federal (terminal).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("entradas", nargs="+", help="arquivo .txt/.csv com CNPJs e/ou CNPJs digitados")
    p.add_argument("-i", "--intervalo", type=float, default=1.5, help="segundos entre consultas (padrão 1.5)")
    p.add_argument("-o", "--csv", metavar="ARQUIVO", help="salva a tabela completa em CSV")
    p.add_argument("-r", "--relatorio", metavar="ARQUIVO", help="salva um relatório .txt")
    p.add_argument("-f", "--filtro", help="mostra na tabela só as linhas com este texto")
    p.add_argument("-c", "--coluna", metavar="NOME",
                   help="nome da coluna com os documentos (planilhas e CSV). "
                        "Padrão: coluna com 'cnpj' ou 'cpf' no cabeçalho")
    p.add_argument("--sem-tabela", action="store_true", help="não imprime a tabela no fim")
    args = p.parse_args()

    itens = montar_itens(args.entradas, args.coluna)
    if not itens:
        p.error("nenhum CNPJ encontrado nas entradas")

    base_nome = ", ".join(os.path.basename(e) for e in args.entradas if os.path.isfile(e)) or "(linha de comando)"
    pendentes = [idx for idx, it in enumerate(itens) if it["status"] == PENDENTE]
    erros_entrada = sum(1 for it in itens if it["status"] == ERRO)
    print(
        f"{len(itens)} linhas encontradas: {len(pendentes)} para consultar, "
        f"{erros_entrada} inválidas (aparecem com o motivo na tabela).\n"
        f"Intervalo de {args.intervalo}s entre consultas. Ctrl+C para parar.\n"
    )

    try:
        for n, idx in enumerate(pendentes, 1):
            cnpj = itens[idx]["cnpj"]
            obs_anterior = itens[idx]["erro"]
            r = consultar(cnpj)
            if obs_anterior and r["status"] == CONSULTADO:
                r["erro"] = obs_anterior  # mantém a observação de zeros completados
            itens[idx] = r
            prefixo = f"[{n}/{len(pendentes)}] {mascara(cnpj)}"
            if r["status"] == CONSULTADO:
                print(f"{prefixo} -> {r['situacao'] or '-'} | {trunc(r['razao_social'] or '-', 50)} ({r['fonte']})")
            else:
                print(f"{prefixo} -> ERRO: {trunc(r['erro'], 100)}")
            if n < len(pendentes) and args.intervalo:
                time.sleep(args.intervalo)
    except KeyboardInterrupt:
        print("\nConsulta interrompida. Salvando o que já foi consultado...")

    r = resumo(itens)
    print(
        f"\nResumo: total {r['total']} | consultados {r['consultados']} | ativas {r['ativas']} | "
        f"não ativas {r['nao_ativas']} | erros {r['erros']} | pendentes {r['pendentes']}"
    )

    if not args.sem_tabela:
        imprimir_tabela(itens, args.filtro)
    if args.csv:
        salvar_csv(itens, args.csv)
        print(f"CSV salvo em: {args.csv}")
    if args.relatorio:
        with open(args.relatorio, "w", encoding="utf-8-sig") as f:
            f.write(gerar_relatorio_texto(itens, base_nome, datetime.now()))
        print(f"Relatório salvo em: {args.relatorio}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
