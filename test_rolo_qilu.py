"""Testes da decisão do nº do rolo Qilu (sem imagem, sem OCR).

Rode com:  .venv\\Scripts\\python test_rolo_qilu.py
Os casos vêm das fotos de 03/10/2026 (leituras reais do OCR).
"""
from datetime import date

from leitor_etiquetas import decidir_rolo_qilu

CASOS = [
    # (leituras, data de fabricação, rolo esperado, trecho que DEVE aparecer nos avisos ou None)
    (["JD260504-32A13108"], date(2026, 5, 4), "JD260504-32A13108", None),
    (["JD26050432A13106"], date(2026, 5, 4), "JD260504-32A13106", None),             # sem traço
    (["11260615-32A14055", "1T260615-32A14055", "JT260615-32A14055", "JT260615-32A14055"],
     date(2026, 6, 15), "JT260615-32A14055", "incertas"),                            # recuperada, mas avisa
    (["11260615-32A14054"] * 4, date(2026, 6, 15), "??260615-32A14054", "ilegíveis"),
    (["1T260615-33A14057", "IT260615-33A14057"], date(2026, 6, 15), "??260615-33A14057", "ilegíveis"),
    (["260324-33A11770", ".260324-33A11770"], date(2026, 3, 24), "??260324-33A11770", "ilegíveis"),
    (["JD260504-32A13113"], date(2026, 5, 5), "JD260504-32A13113", "não confere com a data"),
    (["PC30-9260026"], date(2026, 3, 30), None, None),                               # outro formato: não mexe
    ([], None, None, None),
]


def main():
    falhas = 0
    for leituras, fab, esperado, aviso in CASOS:
        rolo, avisos = decidir_rolo_qilu(leituras, fab)
        ok = rolo == esperado and (aviso is None and not avisos or aviso is not None and any(aviso in a for a in avisos))
        falhas += not ok
        print("%s %-20s %-20s %s" % ("ok " if ok else "XX ", leituras[:1], rolo, avisos))
    print("\n%d falha(s)" % falhas)
    return 1 if falhas else 0


if __name__ == "__main__":
    raise SystemExit(main())
