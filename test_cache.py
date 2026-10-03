"""Teste do cache de leitura na pasta das fotos (roda o leitor de verdade).

Rode com:  .venv\\Scripts\\python test_cache.py
Usa 3 fotos de etiquetas\\ copiadas para uma pasta temporária.
"""
import json, os, shutil, subprocess, sys, tempfile, time
from pathlib import Path

AQUI = Path(__file__).parent
FOTOS = ["WhatsApp Image 2026-10-03 at 16.06.57.jpeg", "WhatsApp Image 2026-10-03 at 16.07.02.jpeg",
         "WhatsApp Image 2026-10-03 at 16.07.05.jpeg"]
EXTRA = "WhatsApp Image 2026-10-03 at 16.07.09.jpeg"


def rodar(pasta, *extra):
    csv = pasta.parent / "saida.csv"
    t = time.time()
    r = subprocess.run([sys.executable, str(AQUI / "leitor_etiquetas.py"), str(pasta), "-o", str(csv), *extra],
                       capture_output=True, text=True, encoding="utf-8", errors="replace",
                       env=dict(os.environ, PYTHONIOENCODING="utf-8"))
    linhas = [l for l in r.stdout.splitlines() if l.startswith("[")]
    origens = [l.rsplit("(", 1)[-1].rstrip(")") for l in linhas]
    conteudo = csv.read_text(encoding="utf-8") if csv.exists() else ""
    return r.returncode, origens, conteudo, time.time() - t, r.stderr


def main():
    falhas = []

    def checar(nome, cond, info=""):
        print("%s %s %s" % ("ok " if cond else "XX ", nome, info))
        if not cond:
            falhas.append(nome)

    base = Path(tempfile.mkdtemp(prefix="cache_etq_"))
    pasta = base / "fotos (com espaço)"
    pasta.mkdir()
    for f in FOTOS:
        shutil.copy2(AQUI / "etiquetas" / f, pasta)
    try:
        cod, o1, csv1, t1, _ = rodar(pasta)
        cache = pasta / ".leitura_etiquetas.json"
        checar("1a leitura le tudo", cod == 0 and o1 == ["lida"] * 3, o1)
        checar("cache criado e oculto", cache.exists() and bool(os.stat(cache).st_file_attributes & 2))

        # 091: imagens para a janela de conferência
        import cv2, numpy as np
        imgs = pasta / ".leitura_etiquetas"
        cab = csv1.splitlines()[0].split(";")
        linhas1 = [dict(zip(cab, l.split(";"))) for l in csv1.splitlines()[1:]]
        checar("colunas novas no FIM", cab[-2:] == ["foto_exibicao", "recorte"], cab[-3:])
        validas = 0
        for l in linhas1:
            for c in ("foto_exibicao", "recorte"):
                if l[c]:
                    im = cv2.imdecode(np.fromfile(l[c], np.uint8), cv2.IMREAD_COLOR)
                    validas += im is not None and l[c].lower().endswith(".jpg") and \
                        (c != "foto_exibicao" or max(im.shape[:2]) <= 1000)
        checar("foto de exibicao e recorte JPEG validos", validas == 6 and bool(os.stat(imgs).st_file_attributes & 2),
               "%d de 6" % validas)
        exe = AQUI / "dist" / "leitor_etiquetas" / "leitor_etiquetas.exe"
        ref = base / "ref.csv"
        if exe.exists():
            # numa COPIA: o pacote de outra versao regravaria o cache desta pasta com a versao dele
            copia = base / "copia_ref"
            shutil.copytree(pasta, copia, ignore=shutil.ignore_patterns(".leitura_etiquetas*"))
            subprocess.run([str(exe), str(copia), "-o", str(ref), "--sem-cache"], capture_output=True)
            # as 20 colunas antigas, menos `caminho` (a referencia roda em outra pasta)
            antigas = lambda t: [";".join(c for i, c in enumerate(l.split(";")[:20]) if i != 1)
                                 for l in t.splitlines()]
            vref = (AQUI / "dist" / "leitor_etiquetas" / "versao.txt").read_text().strip()
            checar("20 colunas antigas iguais as do pacote %s" % vref,
                   antigas(ref.read_text(encoding="utf-8")) == antigas(csv1))

        cod, o2, csv2, t2, _ = rodar(pasta)
        checar("2a leitura vem do cache", cod == 0 and o2 == ["cache"] * 3, "%s (%.1fs x %.1fs)" % (o2, t2, t1))
        checar("CSV identico ao da leitura completa", csv2 == csv1)

        shutil.copy2(AQUI / "etiquetas" / EXTRA, pasta)
        cod, o3, csv3, _, _ = rodar(pasta)
        checar("foto nova: so ela e lida", o3.count("lida") == 1 and o3.count("cache") == 3, o3)

        alvo = pasta / FOTOS[1]
        os.utime(alvo, (time.time(), time.time() + 5))       # foto "alterada"
        cod, o4, _, _, _ = rodar(pasta)
        checar("foto alterada e relida", o4.count("lida") == 1, o4)

        recorte = next(imgs.glob("*16.07.05*.numero.jpg"), None) or next(imgs.glob("*.numero.jpg"))
        recorte.unlink()
        cod, o5, _, _, _ = rodar(pasta)
        checar("imagem apagada: so aquela foto e relida", o5.count("lida") == 1 and recorte.exists(), o5)

        (pasta / EXTRA).unlink()
        rodar(pasta)
        nomes = set(json.loads(cache.read_text(encoding="utf-8"))["fotos"])
        checar("foto removida sai do cache", EXTRA not in nomes and len(nomes) == 3, sorted(nomes))

        cod, o6, csv6, _, _ = rodar(pasta, "--sem-cache")
        checar("--sem-cache le tudo", o6 == ["lida"] * 3 and csv6 == csv1, o6)

        dados = json.loads(cache.read_text(encoding="utf-8"))
        dados["versao"] = "versao-antiga"
        os.chmod(cache, 0o666)
        subprocess.run(["attrib", "-h", str(cache)], capture_output=True)
        cache.write_text(json.dumps(dados), encoding="utf-8")
        cod, o7, _, _, _ = rodar(pasta)
        checar("versao diferente do leitor invalida o cache", o7 == ["lida"] * 3, o7)

        cache.unlink()
        (pasta / ".leitura_etiquetas.json.tmp").mkdir()       # impede a gravacao do cache
        cod, o8, csv8, _, err = rodar(pasta)
        checar("pasta sem gravacao: le e segue sem cache", cod == 0 and csv8 == csv1 and "cache" in err,
               err.strip()[:80])
    finally:
        shutil.rmtree(base, ignore_errors=True)
    print("\n%d falha(s)" % len(falhas))
    return 1 if falhas else 0


if __name__ == "__main__":
    raise SystemExit(main())
