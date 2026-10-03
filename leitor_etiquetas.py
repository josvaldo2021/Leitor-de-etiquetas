"""Leitor de etiquetas de rolos de PVB.

Lê fotos de etiquetas com OCR local (RapidOCR), extrai os campos e gera uma
planilha Excel com uma linha por foto.

Uso:
    python leitor_etiquetas.py [pasta_das_fotos] [-o saida.xlsx]
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from dataclasses import dataclass
from datetime import date
from pathlib import Path

# Espessuras comerciais de PVB (mm). Fora disso o valor provavelmente foi mal lido.
ESPESSURAS_PADRAO = {0.38, 0.76, 1.14, 1.52, 2.28}
CONFIANCA_MINIMA = 0.90
EXTENSOES = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}


# --------------------------------------------------------------------------- #
# Geometria das caixas de texto do OCR
# --------------------------------------------------------------------------- #
@dataclass
class Caixa:
    texto: str
    conf: float
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2

    @property
    def h(self) -> float:
        return self.y1 - self.y0

    @property
    def norm(self) -> str:
        """Texto em maiúsculas, sem espaços, com pontuação chinesa trocada."""
        return self.texto.upper().replace(" ", "").replace("：", ":")


def caixas_de_ocr(resultado) -> list[Caixa]:
    """Converte (boxes, txts, scores) do RapidOCR em objetos Caixa."""
    caixas = []
    for box, txt, conf in resultado:
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        caixas.append(Caixa(txt.strip(), float(conf), min(xs), min(ys), max(xs), max(ys)))
    return caixas


# --------------------------------------------------------------------------- #
# Definição dos campos
# --------------------------------------------------------------------------- #
# Cada campo tem rótulos-âncora (regex sobre Caixa.norm) e uma regex que o valor
# precisa casar. O valor é procurado (1) colado no próprio rótulo, ex.
# "生产日期:2026/3/30", ou (2) na caixa mais próxima à direita na mesma linha.
# Rótulos chineses vêm primeiro: na etiqueta Qilu o valor fica alinhado com o
# chinês, e o inglês (menor, embaixo) fica a meio caminho da linha seguinte.
NUM = r"\d{1,4}(?:\.\d{1,3})?"
INTEIRO = r"\d{1,5}"

CAMPOS: dict[str, tuple[list[str], str]] = {
    "rolo": ([r"^货号:?", r"^ROLLNO\.?:?"], r"[A-Z]{1,3}\d[A-Z0-9.\-]{6,}"),
    "cor": ([r"^颜色:?", r"^COLOR(?!BAND)"], r"[A-Z][A-Z ]{2,}"),
    "codigo": ([r"^CODE"], r"[A-Z0-9+]{1,6}"),
    "grau": ([r"^用途等级:?", r"^GRADE"], r"[A-Z][A-Z0-9+]{0,3}"),
    "espessura_mm": ([r"^厚度:?", r"^THICKNESS"], r"\d\.\d{1,2}"),
    "largura_mm": ([r"^宽度:?", r"^WIDTH"], INTEIRO),
    "comprimento_m": ([r"^长度:?", r"^LENGTH"], INTEIRO),
    "area_m2": ([r"^面积:?", r"^SQUARE"], NUM),
    "peso_liquido_kg": ([r"^NETW\w{4,5}$"], INTEIRO),
    "peso_bruto_kg": ([r"^GROSS\w{5,6}$"], INTEIRO),
    "data_fabricacao": ([r"^生产日期:?", r"^MFD:?", r"^DATE"], r"20\d{2}[./\-]\d{1,2}[./\-]\d{1,2}"),
    "validade": ([r"^保质期:?", r"^EXP:?", r"^VALIDITY:?"], r"\d{1,2}(?:YEARS?|MONTHS?|个月)"),
    "inspetor": ([r"^检验员:?", r"^INSPECTOR"], r".{2,}"),
}

CAMPOS_INTEIROS = {"largura_mm", "comprimento_m", "peso_liquido_kg", "peso_bruto_kg"}
CAMPOS_NUMERICOS = CAMPOS_INTEIROS | {"espessura_mm", "area_m2"}


def limpar_valor(campo: str, texto: str) -> str:
    t = texto.replace("：", ":").strip(" :")
    if campo == "inspetor":
        return t
    t = t.upper()
    if campo in CAMPOS_NUMERICOS:
        t = t.replace(",", ".").replace(" ", "")
        t = t.translate(str.maketrans("OIL", "011"))
        if campo in CAMPOS_INTEIROS:
            # carimbo falhado: "3210" às vezes sai "321.0"
            t = t.replace(".", "")
    elif campo == "rolo":
        # descarta texto chinês colado no final, ex. "DT260612-41A10005用途等级"
        t = re.match(r"[A-Z0-9.\-]*", t.replace(" ", ""))[0]
    elif campo in ("validade", "data_fabricacao", "codigo", "grau"):
        t = t.replace(" ", "")
    return t


def ler_campo(caixas: list[Caixa], campo: str) -> tuple[str | None, float]:
    ancoras, padrao_valor = CAMPOS[campo]
    regex_valor = re.compile(rf"^(?:{padrao_valor})$")
    for padrao_ancora in ancoras:
        regex_ancora = re.compile(padrao_ancora)
        for ancora in caixas:
            m = regex_ancora.search(ancora.norm)
            if not m:
                continue
            # (1) valor colado no rótulo (preserva espaços quando dá, ex. nomes)
            original = ancora.texto.upper().replace("：", ":")
            m_orig = regex_ancora.search(original)
            resto = limpar_valor(campo, ancora.texto[m_orig.end():] if m_orig else ancora.norm[m.end():])
            if resto and regex_valor.match(resto):
                return resto, ancora.conf
            # (2) caixa à direita na mesma linha; desalinhamento vertical pesa
            # mais que distância horizontal, para não pegar a linha vizinha.
            candidatas = sorted(
                (c for c in caixas
                 if c is not ancora
                 and c.x0 > ancora.x0 + ancora.h * 0.5
                 and abs(c.cy - ancora.cy) < 0.8 * max(ancora.h, c.h)),
                key=lambda c: (c.x0 - ancora.x1) + 4 * abs(c.cy - ancora.cy),
            )
            for c in candidatas:
                valor = limpar_valor(campo, c.texto)
                if regex_valor.match(valor):
                    return valor, c.conf
    return None, 0.0


# --------------------------------------------------------------------------- #
# Normalização e validação
# --------------------------------------------------------------------------- #
def detectar_modelo(caixas: list[Caixa]) -> tuple[str, str]:
    tudo = " ".join(c.norm for c in caixas)
    if "QILU" in tudo or "齐鲁" in tudo:
        return "Qilu", "Shandong Qilu Ethylene Chemicals"
    if "INTERLAYER" in tudo or "SHADEWIDTH" in tudo:
        return "PVB Interlayer", ""
    return "Desconhecido", ""


def para_data(texto: str | None) -> date | None:
    if not texto:
        return None
    m = re.match(r"(20\d{2})[./\-](\d{1,2})[./\-](\d{1,2})", texto)
    if not m:
        return None
    try:
        return date(int(m[1]), int(m[2]), int(m[3]))
    except ValueError:
        return None


def validade_em_meses(texto: str | None) -> int | None:
    if not texto:
        return None
    m = re.match(r"(\d+)(YEAR|MONTH|个月)", texto)
    if not m:
        return None
    return int(m[1]) * (12 if m[2] == "YEAR" else 1)


def somar_meses(d: date, meses: int) -> date:
    ano, mes = divmod(d.month - 1 + meses, 12)
    ano += d.year
    dia = min(d.day, [31, 29 if ano % 4 == 0 else 28, 31, 30, 31, 30,
                      31, 31, 30, 31, 30, 31][mes])
    return date(ano, mes + 1, dia)


def normalizar_rolo(rolo: str | None) -> str | None:
    if not rolo:
        return None
    # Modelo "PVB Interlayer": B1.2026.0622.0039 -- reconstrói os pontos
    digitos = re.sub(r"[^0-9]", "", rolo[2:])
    if re.match(r"^B\d", rolo) and len(digitos) == 12:
        return f"{rolo[:2]}.{digitos[:4]}.{digitos[4:8]}.{digitos[8:]}"
    # Modelo Qilu: DT260612-41A10002 / PC30-9260026 -- não tem pontos
    return rolo.replace(".", "")


# --------------------------------------------------------------------------- #
# Nº do rolo da Qilu: segunda leitura e conferência com a data
# --------------------------------------------------------------------------- #
# Estrutura medida em 03/10/2026 nas etiquetas da Qilu: 2 letras + data de
# fabricação AAMMDD + "-" + 2 dígitos + "A" + sequencial (ex. JD260504-32A13113).
# Os NÚMEROS saem certos em toda leitura; as 2 LETRAS do começo vêm desbotadas na
# impressão e o OCR as troca com segurança alta (JD -> "JO" 0,95; JT -> "11" 0,97):
# confiança do OCR não serve de alarme. O que serve é ler de novo e comparar.
RE_ROLO_QILU = re.compile(r"^([A-Z0-9]{0,3}?)(\d{6})[-.]?(\d{2})A(\d{4,6})$")


def decidir_rolo_qilu(leituras: list[str], fabricacao: date | None) -> tuple[str | None, list[str]]:
    """Decide o nº do rolo a partir de várias leituras da mesma faixa.

    Devolve (rolo, avisos). rolo None = nenhuma leitura tem a estrutura Qilu
    (ex. "PC30-9260026"): quem chama mantém a primeira leitura. Letras que não dá
    para afirmar saem como "??" -- nunca chutadas -- e geram aviso, o que impede a
    entrada automática no estoque até alguém conferir a foto."""
    partes = []
    for bruto in leituras:
        t = re.sub(r"[\s]", "", (bruto or "").upper()).strip(".-:")
        m = RE_ROLO_QILU.match(t)
        if m:
            partes.append((m.group(1), f"{m.group(2)}-{m.group(3)}A{m.group(4)}", m.group(2)))
    if not partes:
        return None, []
    corpo, _ = Counter(p[1] for p in partes).most_common(1)[0]
    do_corpo = [p for p in partes if p[1] == corpo]
    validos = [p[0] for p in do_corpo if re.fullmatch(r"[A-Z]{2}", p[0])]
    lidos = ", ".join(dict.fromkeys(p[0] or "-" for p in do_corpo))
    avisos = []
    prefixo, n = Counter(validos).most_common(1)[0] if validos else ("", 0)
    if validos and n == len(do_corpo):
        pass                                    # todas as leituras concordam
    elif n >= 2:
        # há discordância, mas DUAS leituras afirmam as mesmas letras
        avisos.append(f"letras do nº incertas (leituras: {lidos}) -- confira a foto")
    else:
        # uma leitura só não basta para afirmar letras ("1T" x "IT": o certo era "JT")
        prefixo = "??"
        avisos.append(f"letras do nº ilegíveis (leituras: {lidos}) -- confira a foto")
    data_no_numero = do_corpo[0][2]
    if fabricacao and data_no_numero != fabricacao.strftime("%y%m%d"):
        avisos.append(f"nº do rolo ({data_no_numero}) não confere com a data de fabricação")
    return prefixo + corpo, avisos


def ancora_do_rolo(caixas: list[Caixa]) -> Caixa | None:
    for padrao in CAMPOS["rolo"][0]:
        for c in caixas:
            if re.search(padrao, c.norm):
                return c
    return None


def leituras_brutas_do_rolo(caixas: list[Caixa]) -> list[str]:
    """Os textos da 1ª leitura na faixa do nº do rolo, SEM o filtro de formato
    (o filtro descarta "11260615-32A14055", que a decisão aproveita)."""
    anc = ancora_do_rolo(caixas)
    if not anc:
        return []
    out = []
    m = re.search(r"^(货号|ROLLNO\.?):?", anc.norm)
    if m and len(anc.norm) > m.end():
        out.append(anc.norm[m.end():])
    for c in caixas:
        if c is not anc and c.x0 > anc.x0 + anc.h * 0.5 and abs(c.cy - anc.cy) < 0.8 * max(anc.h, c.h):
            if len(re.findall(r"\d", c.texto)) >= 6:
                out.append(c.texto)
    return out


def precisa_reler(caixas: list[Caixa]) -> bool:
    """So relê quando a 1ª leitura NAO deu um nº Qilu com 2 letras: reler custa
    ~2 leituras de OCR, e quando a 1ª leitura e consistente (mesmo errada, como o
    "JO") a releitura repete o mesmo resultado."""
    if detectar_modelo(caixas)[0] != "Qilu":
        return False
    for t in leituras_brutas_do_rolo(caixas):
        m = RE_ROLO_QILU.match(re.sub(r"\s", "", t.upper()).strip(".-:"))
        if m and re.fullmatch(r"[A-Z]{2}", m.group(1)):
            return False
    return True


def extrair(caixas: list[Caixa], releituras: list[str] | None = None) -> dict:
    modelo, fabricante = detectar_modelo(caixas)
    reg: dict = {"modelo": modelo, "fabricante": fabricante}
    confs: dict[str, float] = {}
    for campo in CAMPOS:
        valor, conf = ler_campo(caixas, campo)
        reg[campo] = valor
        if valor is not None:
            confs[campo] = conf

    for campo in CAMPOS_NUMERICOS:
        if reg[campo] is not None:
            reg[campo] = (int if campo in CAMPOS_INTEIROS else float)(reg[campo])

    reg["rolo"] = normalizar_rolo(reg["rolo"])
    if reg["cor"]:
        reg["cor"] = re.sub(r"^PVB(?=\S)", "PVB ", reg["cor"])

    reg["data_fabricacao"] = para_data(reg["data_fabricacao"])

    avisos_rolo: list[str] = []
    if modelo == "Qilu":
        leituras = leituras_brutas_do_rolo(caixas) + list(releituras or [])
        decidido, avisos_rolo = decidir_rolo_qilu(leituras, reg["data_fabricacao"])
        if decidido:
            reg["rolo"] = decidido
            confs["rolo"] = 1.0       # a confiança agora é a concordância entre leituras (avisos_rolo)

    meses = validade_em_meses(reg["validade"])
    reg["validade_meses"] = meses
    reg["data_vencimento"] = (somar_meses(reg["data_fabricacao"], meses)
                              if reg["data_fabricacao"] and meses else None)

    if reg["area_m2"] is None and reg["largura_mm"] and reg["comprimento_m"]:
        reg["area_m2"] = round(reg["largura_mm"] / 1000 * reg["comprimento_m"], 1)
        reg["area_calculada"] = True

    reg["avisos"] = validar(reg, confs) + avisos_rolo
    return reg


def validar(reg: dict, confs: dict[str, float]) -> list[str]:
    avisos = []
    obrigatorios = ["rolo", "cor", "espessura_mm", "largura_mm", "comprimento_m", "data_fabricacao"]
    faltando = [c for c in obrigatorios if reg.get(c) in (None, "")]
    if faltando:
        avisos.append("não lido: " + ", ".join(faltando))

    baixa = [c for c, v in confs.items() if v < CONFIANCA_MINIMA]
    if baixa:
        avisos.append("baixa confiança: " + ", ".join(baixa))

    esp = reg.get("espessura_mm")
    if esp is not None and esp not in ESPESSURAS_PADRAO:
        avisos.append(f"espessura incomum ({esp})")
    if reg.get("largura_mm") is not None and not 500 <= reg["largura_mm"] <= 4000:
        avisos.append(f"largura fora do normal ({reg['largura_mm']})")

    larg, comp, area = reg.get("largura_mm"), reg.get("comprimento_m"), reg.get("area_m2")
    if larg and comp and area and not reg.get("area_calculada"):
        esperado = larg / 1000 * comp
        if abs(esperado - area) > max(2, esperado * 0.01):
            avisos.append(f"área {area} ≠ largura×comprimento {esperado:.0f}")

    liq, bruto = reg.get("peso_liquido_kg"), reg.get("peso_bruto_kg")
    if liq and bruto and liq >= bruto:
        avisos.append("peso líquido ≥ bruto")

    # A data costuma estar codificada no nº do rolo; se não bater, um dos dois foi mal lido.
    # (Qilu: conferido em decidir_rolo_qilu, pela estrutura do número.)
    rolo, fab = reg.get("rolo"), reg.get("data_fabricacao")
    if rolo and fab and reg.get("modelo") != "Qilu":
        digitos = re.sub(r"\D", "", rolo)
        codigos = (fab.strftime("%Y%m%d"), fab.strftime("%y%m%d"))
        if rolo.startswith("B") and not any(c in digitos for c in codigos):
            avisos.append("data de fabricação não confere com nº do rolo")
    return avisos


# --------------------------------------------------------------------------- #
# OCR
# --------------------------------------------------------------------------- #
class LeitorOCR:
    def __init__(self):
        import logging
        from rapidocr import RapidOCR
        self._motor = RapidOCR()
        logging.getLogger("RapidOCR").setLevel(logging.WARNING)

    def ler(self, caminho: Path) -> list[Caixa]:
        import cv2
        import numpy as np
        # cv2.imread não abre caminhos com acentos no Windows; imdecode abre.
        img = cv2.imdecode(np.fromfile(str(caminho), dtype=np.uint8), cv2.IMREAD_COLOR)
        if img is None:
            raise ValueError("não foi possível abrir a imagem")
        melhor: list[Caixa] = []
        self.imagem = img
        # Se a foto estiver deitada/invertida, quase nenhum rótulo é encontrado:
        # tenta as outras rotações e fica com a que achar mais campos.
        for rot in (None, cv2.ROTATE_90_CLOCKWISE, cv2.ROTATE_90_COUNTERCLOCKWISE, cv2.ROTATE_180):
            girada = img if rot is None else cv2.rotate(img, rot)
            r = self._motor(girada)
            caixas = caixas_de_ocr(zip(r.boxes, r.txts, r.scores)) if r.txts else []
            if pontuacao(caixas) > pontuacao(melhor):
                melhor = caixas
                self.imagem = girada        # a releitura recorta na MESMA rotação das caixas
            # 4 de 5 campos achados = a foto esta de pe (um campo desbotado, como o nº
            # do rolo, nao pode custar 3 leituras inteiras a mais em outras rotacoes)
            if pontuacao(melhor) >= 4:
                break
        return melhor

    def faixa_do_rolo(self, caixas: list[Caixa], alturas: float, margem: float = 0.6):
        """A faixa da imagem à direita do rótulo do nº do rolo, na rotação lida.
        `alturas`: largura em alturas do rótulo (13 cobre o nº da Qilu; o recorte
        da conferência usa mais, para caber o nº mais longo do "PVB Interlayer")."""
        anc = ancora_do_rolo(caixas)
        img = getattr(self, "imagem", None)
        if anc is None or img is None:
            return None
        y0, y1 = int(anc.y0 - anc.h * margem), int(anc.y1 + anc.h * margem)
        x0, x1 = int(anc.x1), int(min(img.shape[1], anc.x1 + alturas * anc.h))
        faixa = img[max(0, y0):max(0, y1), max(0, x0):max(0, x1)]
        return faixa if faixa.size else None

    def gravar_imagens(self, caixas: list[Caixa], pasta_img: Path, nome: str) -> tuple[str, str]:
        """Fatia 091 do estoque: a FOTO DE EXIBIÇÃO (JPEG, de pé, lado maior <= 1000 px)
        e o RECORTE do nº do rolo (JPEG, faixa ampliada 2x, sem binarizar: o operador
        precisa ver o traço desbotado como ele é). Sempre JPEG baseline: o UserForm do
        Excel não abre PNG (medido: erro 481). Devolve os caminhos, ou "" se não deu."""
        import cv2
        img = getattr(self, "imagem", None)
        if img is None:
            return "", ""
        try:
            pasta_img.mkdir(exist_ok=True)
            if sys.platform == "win32":
                import ctypes
                ctypes.windll.kernel32.SetFileAttributesW(str(pasta_img), 0x02)   # oculta
        except OSError:
            return "", ""
        params = [cv2.IMWRITE_JPEG_QUALITY, 85, cv2.IMWRITE_JPEG_PROGRESSIVE, 0]
        foto = pasta_img / f"{nome}.foto.jpg"
        escala = 1000 / max(img.shape[:2])
        menor = cv2.resize(img, None, fx=escala, fy=escala, interpolation=cv2.INTER_AREA) if escala < 1 else img
        caminho_foto = str(foto) if self._gravar_jpeg(menor, foto, params) else ""
        caminho_rec = ""
        # margem maior que a da releitura: no "PVB Interlayer" o nº tem letra maior
        # que o rótulo "ROLL NO." e a margem de 0,6 cortava o topo dos dígitos
        faixa = self.faixa_do_rolo(caixas, 18, margem=1.2)
        if faixa is not None:
            grande = cv2.resize(faixa, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
            rec = pasta_img / f"{nome}.numero.jpg"
            caminho_rec = str(rec) if self._gravar_jpeg(grande, rec, params) else ""
        return caminho_foto, caminho_rec

    @staticmethod
    def _gravar_jpeg(img, destino: Path, params) -> bool:
        import cv2
        try:
            ok, buf = cv2.imencode(".jpg", img, params)
            if ok:
                buf.tofile(str(destino))       # tofile aceita caminho com acento; imwrite não
            return bool(ok)
        except OSError:
            return False

    def reler_faixa_do_rolo(self, caixas: list[Caixa]) -> list[str]:
        """Segunda leitura SÓ da faixa do nº do rolo, em três versões da imagem
        (ampliada, contraste local, preto-e-branco). Medido em 03/10/2026: recupera
        letras desbotadas que a 1ª leitura troca ("11" -> "JT") e, quando não
        recupera, as leituras discordam -- que é o alarme."""
        import cv2
        faixa = self.faixa_do_rolo(caixas, 13)
        if faixa is None:
            return []
        cinza = cv2.resize(cv2.cvtColor(faixa, cv2.COLOR_BGR2GRAY), None, fx=2, fy=2,
                           interpolation=cv2.INTER_CUBIC)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8)).apply(cinza)
        otsu = cv2.threshold(clahe, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        out = []
        for versao in (clahe, otsu):
            r = self._motor(cv2.cvtColor(versao, cv2.COLOR_GRAY2BGR))
            for t in (r.txts or []):
                if len(re.findall(r"\d", t)) >= 6:
                    out.append(t)
        return out


def pontuacao(caixas: list[Caixa]) -> int:
    return sum(ler_campo(caixas, c)[0] is not None
               for c in ("rolo", "espessura_mm", "largura_mm", "comprimento_m", "data_fabricacao"))


# --------------------------------------------------------------------------- #
# Excel
# --------------------------------------------------------------------------- #
COLUNAS = [
    ("arquivo", "Arquivo", 34),
    ("modelo", "Modelo etiqueta", 15),
    ("fabricante", "Fabricante", 22),
    ("rolo", "Nº do rolo", 22),
    ("cor", "Cor", 16),
    ("espessura_mm", "Espessura (mm)", 10),
    ("largura_mm", "Largura (mm)", 10),
    ("comprimento_m", "Comprimento (m)", 12),
    ("area_m2", "Área (m²)", 10),
    ("peso_liquido_kg", "Peso líq. (kg)", 10),
    ("peso_bruto_kg", "Peso bruto (kg)", 10),
    ("codigo", "Code", 8),
    ("grau", "Grade", 8),
    ("data_fabricacao", "Fabricação", 12),
    ("validade_meses", "Validade (meses)", 10),
    ("data_vencimento", "Vencimento", 12),
    ("inspetor", "Inspetor", 16),
    ("avisos", "Avisos (conferir)", 60),
]


def salvar_excel(registros: list[dict], destino: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "Etiquetas PVB"
    ws.append([titulo for _, titulo, _ in COLUNAS])
    for i, (_, _, largura) in enumerate(COLUNAS, start=1):
        cel = ws.cell(row=1, column=i)
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor="305496")
        cel.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[cel.column_letter].width = largura

    amarelo = PatternFill("solid", fgColor="FFF2CC")
    for reg in registros:
        linha = []
        for chave, _, _ in COLUNAS:
            v = reg.get(chave)
            linha.append("; ".join(v) if isinstance(v, list) else v)
        ws.append(linha)
        r = ws.max_row
        for i, (chave, _, _) in enumerate(COLUNAS, start=1):
            if chave.startswith("data_"):
                ws.cell(row=r, column=i).number_format = "DD/MM/YYYY"
        if reg.get("avisos"):
            for i in range(1, len(COLUNAS) + 1):
                ws.cell(row=r, column=i).fill = amarelo

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(destino)


def chave_do_rolo(rolo: str | None) -> str:
    """Nº do rolo só com letras e dígitos: 'PC30 - 9260027' e 'PC30-9260027'
    são o mesmo rolo. É a chave que o estoque (VBA) usa para comparar."""
    return re.sub(r"[^0-9A-Z]", "", (rolo or "").upper())


# Formato fixo para outro programa ler (o estoque em VBA): UTF-8, separador ";",
# datas ISO e ponto decimal -- nada que dependa da configuração regional.
# Os nomes das colunas são o contrato; acrescentar coluna no fim não quebra quem lê.
COLUNAS_CSV = ["arquivo", "caminho", "modelo", "fabricante", "rolo", "rolo_chave", "cor",
               "espessura_mm", "largura_mm", "comprimento_m", "area_m2",
               "peso_liquido_kg", "peso_bruto_kg", "codigo", "grau", "data_fabricacao",
               "validade_meses", "data_vencimento", "inspetor", "avisos",
               # leitor 2026-10-03.4: imagens para a janela de conferência (fatia 091)
               "foto_exibicao", "recorte"]


def salvar_csv(registros: list[dict], destino: Path) -> None:
    def celula(v) -> str:
        if v is None:
            return ""
        if isinstance(v, list):
            v = " | ".join(v)
        elif isinstance(v, date):
            v = v.isoformat()
        # sem aspas no formato: o separador e quebras de linha não podem aparecer no valor
        return str(v).replace(";", ",").replace("\r", " ").replace("\n", " ")

    linhas = [";".join(COLUNAS_CSV)]
    for reg in registros:
        reg = {**reg, "rolo_chave": chave_do_rolo(reg.get("rolo"))}
        linhas.append(";".join(celula(reg.get(c)) for c in COLUNAS_CSV))
    destino.write_text("\n".join(linhas) + "\n", encoding="utf-8")


# --------------------------------------------------------------------------- #
# Cache de leitura (na pasta das fotos)
# --------------------------------------------------------------------------- #
# Fechar a tela de importação no estoque descarta a prévia; reabrir a mesma pasta
# relia tudo (~5 s por foto). O cache guarda o resultado de cada foto ao lado dela:
# foto igual (nome, tamanho, data) + mesma versão do leitor = resultado na hora.
# Versão diferente invalida tudo -- leitura feita com regras antigas não volta.
NOME_CACHE = ".leitura_etiquetas.json"
PASTA_IMAGENS = ".leitura_etiquetas"          # foto de exibição e recorte (JPEG), oculta
CAMPOS_IMAGEM = ("foto_exibicao", "recorte")
CAMPOS_DATA = ("data_fabricacao", "data_vencimento")


def versao_do_leitor() -> str:
    """No pacote: o versao.txt ao lado do .exe. Rodando pelo Python: uma impressão
    digital do próprio código (qualquer mudança nele invalida o cache)."""
    import hashlib
    if getattr(sys, "frozen", False):
        arq = Path(sys.executable).with_name("versao.txt")
        if arq.exists():
            return arq.read_text(encoding="ascii", errors="replace").strip()
    return "dev-" + hashlib.sha1(Path(__file__).read_bytes()).hexdigest()[:12]


class CacheDeLeitura:
    def __init__(self, pasta: Path, versao: str, ligado: bool = True):
        import json
        self.caminho = pasta / NOME_CACHE
        self.pasta_img = pasta / PASTA_IMAGENS
        self.versao = versao
        self.fotos: dict = {}
        self.usados = 0
        if not ligado:
            return
        try:
            dados = json.loads(self.caminho.read_text(encoding="utf-8"))
            if dados.get("versao") == versao:
                self.fotos = dados.get("fotos", {})
        except (OSError, ValueError):
            pass                       # sem cache, cache corrompido ou de outro formato: lê tudo

    @staticmethod
    def _assinatura(foto: Path) -> list:
        st = foto.stat()
        return [st.st_size, st.st_mtime_ns]

    def buscar(self, foto: Path) -> dict | None:
        item = self.fotos.get(foto.name)
        if not item or item.get("assinatura") != self._assinatura(foto):
            return None
        reg = dict(item["reg"])
        for c in CAMPOS_DATA:
            if reg.get(c):
                reg[c] = date.fromisoformat(reg[c])
        reg["avisos"] = list(reg.get("avisos") or [])
        # as imagens são parte do resultado: faltando uma, a foto é lida de novo
        for c in CAMPOS_IMAGEM:
            if reg.get(c):
                absoluto = self.pasta_img / reg[c]
                if not absoluto.exists():
                    return None
                reg[c] = str(absoluto)
        self.usados += 1
        return reg

    def guardar(self, foto: Path, reg: dict) -> None:
        if any(a.startswith("erro:") for a in reg.get("avisos", [])):
            return                     # falha não fica guardada: tenta de novo na próxima vez
        limpo = {k: (v.isoformat() if isinstance(v, date) else v) for k, v in reg.items()
                 if k not in ("arquivo", "caminho")}
        limpo["avisos"] = list(reg.get("avisos") or [])
        for c in CAMPOS_IMAGEM:                  # guardado pelo nome: a pasta pode mudar de letra/caminho
            if limpo.get(c):
                limpo[c] = Path(limpo[c]).name
        self.fotos[foto.name] = {"assinatura": self._assinatura(foto), "reg": limpo}

    def salvar(self, presentes: list[Path]) -> None:
        """Grava só as fotos que ainda estão na pasta. Pasta sem permissão de
        gravação (rede) não é erro: o leitor segue, só sem cache."""
        import json, os
        nomes = {f.name for f in presentes}
        dados = {"versao": self.versao, "fotos": {k: v for k, v in self.fotos.items() if k in nomes}}
        tmp = self.caminho.with_name(NOME_CACHE + ".tmp")
        try:
            tmp.write_text(json.dumps(dados, ensure_ascii=False), encoding="utf-8")
            os.replace(tmp, self.caminho)
            if sys.platform == "win32":
                import ctypes
                ctypes.windll.kernel32.SetFileAttributesW(str(self.caminho), 0x02)   # oculto
        except OSError as e:
            print(f"(cache de leitura não gravado: {e})", file=sys.stderr)


# --------------------------------------------------------------------------- #
# Programa principal
# --------------------------------------------------------------------------- #
def marcar_duplicados(registros: list[dict]) -> None:
    vistos: dict[str, str] = {}
    for reg in registros:
        rolo = reg.get("rolo")
        if not rolo:
            continue
        if rolo in vistos:
            reg["avisos"].append(f"rolo repetido (também em {vistos[rolo]})")
        else:
            vistos[rolo] = reg["arquivo"]


def main() -> int:
    ap = argparse.ArgumentParser(description="Extrai dados de etiquetas de PVB para Excel ou CSV.")
    ap.add_argument("pasta", nargs="?", default="etiquetas", help="pasta com as fotos (padrão: etiquetas)")
    ap.add_argument("-o", "--saida", default="etiquetas_pvb.xlsx",
                    help="arquivo de saída: .xlsx (para ler) ou .csv (para o estoque importar)")
    ap.add_argument("--sem-cache", action="store_true",
                    help="lê todas as fotos de novo, ignorando o cache da pasta")
    args = ap.parse_args()

    pasta = Path(args.pasta)
    fotos = sorted(p for p in pasta.iterdir() if p.suffix.lower() in EXTENSOES) if pasta.is_dir() else []
    if not fotos:
        print(f"Nenhuma imagem encontrada em '{pasta}'.", file=sys.stderr)
        return 1

    cache = CacheDeLeitura(pasta, versao_do_leitor(), ligado=not args.sem_cache)
    print(f"{len(fotos)} fotos")
    leitor = None                       # o OCR só carrega se alguma foto precisar ser lida
    registros = []
    for n, foto in enumerate(fotos, start=1):
        reg = cache.buscar(foto)
        origem = "cache"
        if reg is None:
            origem = "lida"
            try:
                if leitor is None:
                    print("Carregando OCR...")
                    leitor = LeitorOCR()
                caixas = leitor.ler(foto)
                releituras = leitor.reler_faixa_do_rolo(caixas) if precisa_reler(caixas) else []
                reg = extrair(caixas, releituras)
                reg["foto_exibicao"], reg["recorte"] = leitor.gravar_imagens(
                    caixas, pasta / PASTA_IMAGENS, foto.name)
            except Exception as e:  # uma foto ruim não deve derrubar o lote
                reg = {"avisos": [f"erro: {e}"]}
            cache.guardar(foto, reg)
        reg["arquivo"] = foto.name
        reg["caminho"] = str(foto.resolve())
        registros.append(reg)
        status = "OK" if not reg["avisos"] else "CONFERIR"
        print(f"[{n}/{len(fotos)}] {status:8} {reg.get('rolo') or '-':22} {foto.name}  ({origem})")
    cache.salvar(fotos)
    if cache.usados:
        print(f"{cache.usados} de {len(fotos)} fotos vieram do cache (já lidas antes, sem mudança).")

    marcar_duplicados(registros)
    saida = Path(args.saida)
    try:
        if saida.suffix.lower() == ".csv":
            salvar_csv(registros, saida)
        else:
            salvar_excel(registros, saida)
    except PermissionError:
        print(f"Não foi possível gravar '{saida}' (está aberto no Excel?).", file=sys.stderr)
        return 1

    conferir = sum(1 for r in registros if r["avisos"])
    print(f"\nArquivo salvo em: {saida.resolve()}")
    print(f"{len(registros) - conferir} OK, {conferir} para conferir (linhas em amarelo).")
    return 0


def executar() -> int:
    """Código de saída é contrato com quem chama (o estoque em VBA):
    0 = arquivo gravado; 1 = pasta sem imagem ou arquivo não gravado; 2 = falha inesperada."""
    # O console do Windows (cp850) não tem todo caractere que uma etiqueta traz;
    # trocar por '?' é melhor que derrubar a leitura no meio por causa de um print.
    for fluxo in (sys.stdout, sys.stderr):
        try:
            fluxo.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    try:
        return main()
    except SystemExit:
        raise
    except Exception as e:
        print(f"Erro ao ler as etiquetas: {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(executar())
