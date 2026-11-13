#!/usr/bin/env python3
"""
Servidor do Painel de Vendas.

- Entrega index.html, style.css e script.js no navegador.
- Grava cada venda na planilha vendas.xlsx (aba "Vendas"), na mesma estrutura da sua tabela.

Como usar:
    pip install openpyxl
    python servidor.py
O navegador abre em http://localhost:8000
"""
import datetime
import json
import mimetypes
import os
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

PASTA = os.path.dirname(os.path.abspath(__file__))
PLANILHA = os.path.join(PASTA, "vendas.xlsx")
PORTA = int(os.environ.get("PORT", "8000"))

# ---- Tabelas de preços e frete (fonte única: o painel lê daqui) ----
PRECOS = {
    "Pen drive 32GB": 45,
    'Monitor OLED 33"': 1790,
    'SmartTV TV 50" 4K': 2950,
    'TV 48" 4K': 2840,
    'TV 60" 4K': 3470,
    "Smartphone Android": 3200,
    "Smartphone iOS": 8300,
}
FRETE = {"Centro-Oeste": 15, "Nordeste": 20, "Norte": 25, "Sul": 15, "Sudeste": 10}  # por unidade
REGIOES = ["Sudeste", "Sul", "Centro-Oeste", "Nordeste", "Norte"]
ESTADOS = {
    "Norte": {"AC": "Acre", "AP": "Amapá", "AM": "Amazonas", "PA": "Pará", "RO": "Rondônia", "RR": "Roraima", "TO": "Tocantins"},
    "Nordeste": {"AL": "Alagoas", "BA": "Bahia", "CE": "Ceará", "MA": "Maranhão", "PB": "Paraíba", "PE": "Pernambuco", "PI": "Piauí", "RN": "Rio Grande do Norte", "SE": "Sergipe"},
    "Centro-Oeste": {"DF": "Distrito Federal", "GO": "Goiás", "MT": "Mato Grosso", "MS": "Mato Grosso do Sul"},
    "Sudeste": {"ES": "Espírito Santo", "MG": "Minas Gerais", "RJ": "Rio de Janeiro", "SP": "São Paulo"},
    "Sul": {"PR": "Paraná", "RS": "Rio Grande do Sul", "SC": "Santa Catarina"},
}
UF_REGIAO = {uf: reg for reg, ufs in ESTADOS.items() for uf in ufs}

CABECALHO = ["Pedido", "Região", "Estado", "Produtos", "Data", "Quantidade", "Valor Unitário", "Frete", "Total"]
MOEDA = '"R$" #,##0.00'
LOCK = threading.Lock()


class ErroPlanilha(Exception):
    def __init__(self, mensagem, status=500):
        super().__init__(mensagem)
        self.mensagem = mensagem
        self.status = status


# ------------------------------------------------------------------ planilha
_borda = Side(style="thin", color="000000")
BORDA = Border(left=_borda, right=_borda, top=_borda, bottom=_borda)
VERDE = PatternFill("solid", fgColor="C6E0B4")
CENTRO = Alignment(horizontal="center", vertical="center")


def _cabecalho(celula, texto):
    celula.value = texto
    celula.fill = VERDE
    celula.font = Font(bold=True)
    celula.alignment = CENTRO
    celula.border = BORDA


def escrever_linha(ws, r, v):
    """Escreve uma venda na linha r. A coluna Total é fórmula: Quantidade x (Valor + Frete)."""
    valores = [
        v["pedido"], "Região " + v["regiao"], v["estado"], v["produto"], v["data"],
        v["quantidade"], v["valorUnit"], v["frete"], f"=F{r}*(G{r}+H{r})",
    ]
    for c, valor in enumerate(valores, 1):
        cel = ws.cell(r, c, valor)
        cel.border = BORDA
        cel.alignment = CENTRO
    ws.cell(r, 5).number_format = "dd/mm/yyyy"
    for c in (7, 8, 9):
        ws.cell(r, c).number_format = MOEDA


def criar_planilha(caminho, vendas=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "Vendas"
    for c, texto in enumerate(CABECALHO, 1):
        _cabecalho(ws.cell(1, c), texto)
    for r, v in enumerate(vendas, 2):
        escrever_linha(ws, r, v)

    _cabecalho(ws["K1"], "Tabela de Preços")
    _cabecalho(ws["L1"], None)
    ws.merge_cells("K1:L1")
    for i, (nome, preco) in enumerate(PRECOS.items(), 2):
        ws.cell(i, 11, nome).border = BORDA
        ws.cell(i, 12, preco).border = BORDA
        ws.cell(i, 12).number_format = MOEDA

    _cabecalho(ws["K12"], "Valores de Frete")
    _cabecalho(ws["L12"], None)
    ws.merge_cells("K12:L12")
    for i, reg in enumerate(["Centro-Oeste", "Nordeste", "Norte", "Sul", "Sudeste"], 13):
        ws.cell(i, 11, "Região " + reg).border = BORDA
        ws.cell(i, 12, FRETE[reg]).border = BORDA
        ws.cell(i, 12).number_format = MOEDA

    _cabecalho(ws["K19"], "Soma")
    ws["L19"] = "=SUM(I2:I100000)"
    ws["L19"].number_format = MOEDA
    ws["L19"].border = BORDA

    for col, largura in zip("ABCDEFGHIJKL", [10, 18, 9, 22, 13, 12, 16, 10, 16, 4, 22, 14]):
        ws.column_dimensions[col].width = largura
    ws.freeze_panes = "A2"
    wb.save(caminho)


def garantir_planilha():
    if not os.path.exists(PLANILHA):
        criar_planilha(PLANILHA)


def abrir(somente_leitura=False):
    garantir_planilha()
    try:
        wb = load_workbook(PLANILHA, data_only=somente_leitura)
    except PermissionError:
        raise ErroPlanilha("Não consegui abrir vendas.xlsx. Feche-a no Excel e tente de novo.", 423)
    except Exception:
        raise ErroPlanilha("vendas.xlsx está corrompida ou não é uma planilha válida.", 500)
    ws = wb["Vendas"] if "Vendas" in wb.sheetnames else wb.active
    return wb, ws


def salvar(wb):
    try:
        wb.save(PLANILHA)
    except PermissionError:
        raise ErroPlanilha("A planilha vendas.xlsx está aberta no Excel. Feche-a e tente de novo.", 423)


def _numero(x):
    x = float(x)
    return int(x) if x == int(x) else round(x, 2)


def _data(valor):
    if isinstance(valor, datetime.datetime):
        return valor.date().isoformat()
    if isinstance(valor, datetime.date):
        return valor.isoformat()
    texto = str(valor or "").strip()
    for formato in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(texto, formato).date().isoformat()
        except ValueError:
            pass
    return ""


def _linhas(ws):
    r = 2
    while ws.cell(r, 1).value not in (None, ""):
        yield r
        r += 1


def ler_vendas():
    with LOCK:
        wb, ws = abrir(somente_leitura=True)
        vendas = []
        for r in _linhas(ws):
            try:
                q, vu, fr = (_numero(ws.cell(r, c).value) for c in (6, 7, 8))
                vendas.append({
                    "pedido": int(ws.cell(r, 1).value),
                    "regiao": str(ws.cell(r, 2).value).replace("Região", "").strip(),
                    "estado": str(ws.cell(r, 3).value).strip(),
                    "produto": str(ws.cell(r, 4).value).strip(),
                    "data": _data(ws.cell(r, 5).value),
                    "quantidade": int(q), "valorUnit": vu, "frete": fr,
                    "total": _numero(q * (vu + fr)),
                })
            except (TypeError, ValueError):
                continue  # linha incompleta: ignora
        wb.close()
        return vendas


def registrar(dados):
    estado = str(dados.get("estado", "")).upper()
    produto = dados.get("produto")
    if estado not in UF_REGIAO:
        raise ErroPlanilha("Estado inválido.", 400)
    if produto not in PRECOS:
        raise ErroPlanilha("Produto inválido.", 400)
    try:
        quantidade = int(dados.get("quantidade"))
        data = datetime.date.fromisoformat(str(dados.get("data")))
    except (TypeError, ValueError):
        raise ErroPlanilha("Data ou quantidade inválida.", 400)
    if not 1 <= quantidade <= 100000:
        raise ErroPlanilha("A quantidade deve ser 1 ou mais.", 400)

    regiao = UF_REGIAO[estado]
    with LOCK:
        wb, ws = abrir()
        ultimo, r = 1000, 2
        while ws.cell(r, 1).value not in (None, ""):
            try:
                ultimo = max(ultimo, int(ws.cell(r, 1).value))
            except (TypeError, ValueError):
                pass
            r += 1
        venda = {
            "pedido": ultimo + 1, "data": data, "regiao": regiao, "estado": estado, "produto": produto,
            "quantidade": quantidade, "valorUnit": PRECOS[produto], "frete": FRETE[regiao],
        }
        escrever_linha(ws, r, venda)
        salvar(wb)
    venda["data"] = data.isoformat()
    venda["total"] = quantidade * (PRECOS[produto] + FRETE[regiao])
    return venda


def excluir(pedido):
    with LOCK:
        wb, ws = abrir()
        linhas = list(_linhas(ws))
        alvo = next((r for r in linhas if str(ws.cell(r, 1).value) == str(pedido)), None)
        if alvo is None:
            raise ErroPlanilha("Venda não encontrada.", 404)
        ultima = linhas[-1]
        for r in range(alvo, ultima):          # sobe as linhas de baixo (só colunas A..H)
            for c in range(1, 9):
                ws.cell(r, c).value = ws.cell(r + 1, c).value
        for c in range(1, 10):                  # limpa a última linha
            cel = ws.cell(ultima, c)
            cel.value = None
            cel.border = Border()
            cel.number_format = "General"
        for r in _linhas(ws):                   # refaz as fórmulas do Total
            ws.cell(r, 9).value = f"=F{r}*(G{r}+H{r})"
        salvar(wb)


# --------------------------------------------------------------------- web
ESTATICOS = {"/": "index.html", "/index.html": "index.html", "/style.css": "style.css", "/script.js": "script.js"}


class Handler(BaseHTTPRequestHandler):
    def _json(self, status, obj):
        corpo = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(corpo)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(corpo)

    def _tratar(self, funcao):
        try:
            funcao()
        except ErroPlanilha as e:
            self._json(e.status, {"erro": e.mensagem})
        except Exception as e:  # erro inesperado: mostra no terminal e avisa o painel
            print("Erro:", repr(e))
            self._json(500, {"erro": "Erro inesperado no servidor. Veja o terminal."})

    def do_GET(self):
        def go():
            if self.path == "/api/config":
                return self._json(200, {"precos": PRECOS, "frete": FRETE, "regioes": REGIOES, "estados": ESTADOS})
            if self.path == "/api/vendas":
                return self._json(200, ler_vendas())
            arquivo = ESTATICOS.get(self.path.split("?")[0])
            if not arquivo:
                return self._json(404, {"erro": "Não encontrado."})
            with open(os.path.join(PASTA, arquivo), "rb") as f:
                corpo = f.read()
            tipo = mimetypes.guess_type(arquivo)[0] or "text/plain"
            self.send_response(200)
            self.send_header("Content-Type", tipo + "; charset=utf-8")
            self.send_header("Content-Length", str(len(corpo)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(corpo)
        self._tratar(go)

    def do_POST(self):
        def go():
            if self.path != "/api/vendas":
                return self._json(404, {"erro": "Não encontrado."})
            tamanho = int(self.headers.get("Content-Length") or 0)
            try:
                dados = json.loads(self.rfile.read(tamanho) or b"{}")
            except ValueError:
                raise ErroPlanilha("Dados inválidos.", 400)
            self._json(201, registrar(dados))
        self._tratar(go)

    def do_DELETE(self):
        def go():
            if not self.path.startswith("/api/vendas/"):
                return self._json(404, {"erro": "Não encontrado."})
            excluir(self.path.rsplit("/", 1)[1])
            self._json(200, {"ok": True})
        self._tratar(go)

    def log_message(self, formato, *args):
        pass  # terminal limpo


if __name__ == "__main__":
    garantir_planilha()
    servidor = ThreadingHTTPServer(("127.0.0.1", PORTA), Handler)
    url = f"http://localhost:{PORTA}"
    print(f"Painel de Vendas rodando em {url}")
    print("Planilha:", PLANILHA)
    print("Para parar: Ctrl+C")
    threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor encerrado.")
