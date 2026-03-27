"""
conversor_pptx.py
============
Converte arquivos .pptx / .ppt com cifras musicais para .txt,
preservando o posicionamento das cifras acima das silabas e
removendo acentos do texto e dos nomes de arquivo.

Dependencias externas:
  - LibreOffice  (comando: libreoffice)
  - Poppler      (comando: pdftotext)

Instalacao no Ubuntu/Debian:
  sudo apt install libreoffice poppler-utils

Uso como modulo:
  from conversor import converter_arquivo, converter_pasta

Uso direto (linha de comando):
  python conversor_pptx.py arquivo.pptx
  python conversor_pptx.py pasta/
  python conversor_pptx.py pasta/ --saida resultados/
"""

import os
import re
import subprocess
import tempfile
import unicodedata
from pathlib import Path


# ---------------------------------------------------------------------------
# Utilitarios de texto
# ---------------------------------------------------------------------------

def _decode_escaped(s: str) -> str:
    """Converte sequencias #Uxxxx (geradas por alguns ZIPs) em unicode real."""
    return re.sub(
        r"#U([0-9a-fA-F]{4})",
        lambda m: chr(int(m.group(1), 16)),
        s,
    )


def remover_acentos(texto: str) -> str:
    """Remove todos os diacriticos de uma string."""
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def nome_sem_acento(nome_arquivo: str) -> str:
    """
    Dado o nome de um arquivo (pode conter sequencias #Uxxxx),
    retorna o nome decodificado e sem acentos.
    """
    return remover_acentos(_decode_escaped(nome_arquivo))


# ---------------------------------------------------------------------------
# Deteccao de linha de cifra
# ---------------------------------------------------------------------------

_CHORD_PAT = re.compile(
    r"\b[A-G][#b]?"
    r"(?:maj\d*|min|m(?:aj)?\d*|dim\d*|aug|sus[24]?|add\d+|\d+)*"
    r"(?:/[A-G][#b]?)?\b"
)
_LABEL_PREFIX = re.compile(
    r"^(?:Introd?\w*|Intro|Tom|Ton|Coda|Refr[aã]o)\s*[:\.]?\s*",
    flags=re.IGNORECASE,
)
_LEFTOVER = re.compile(r"[\s()\[\]x0-9+\-/#_.,;:!?|'\"]+")


def _e_linha_de_cifra(linha: str) -> bool:
    """
    Retorna True se a linha for composta predominantemente por cifras
    (acordes musicais), possivelmente com prefixo de label (ex: 'Introd.:').
    """
    s = linha.strip()
    if not s:
        return False
    s = _LABEL_PREFIX.sub("", s).strip()
    s2 = _CHORD_PAT.sub("", s)
    s2 = _LEFTOVER.sub("", s2)
    return len(s2) <= 2


# ---------------------------------------------------------------------------
# Limpeza de encoding corrompido gerado pelo pdftotext no Windows
# ---------------------------------------------------------------------------

def _limpar_encoding(texto: str) -> str:
    """
    Corrige sequencias corrompidas geradas pelo pdftotext quando
    o PDF tem codificacao Latin-1/Windows-1252 interpretada como UTF-8.
    Depois remove todos os acentos restantes.
    """
    # Tenta reinterpretar como latin-1 se houver muitos caracteres suspeitos
    suspeitos = sum(1 for c in texto if '\x80' <= c <= '\xff')
    if suspeitos > len(texto) * 0.05:
        try:
            texto = texto.encode('latin-1').decode('utf-8', errors='replace')
        except Exception:
            pass

    # Substitui sequencias conhecidas de encoding errado
    substituicoes = [
        ('Ã£', 'a'), ('Ã§', 'c'), ('Ã©', 'e'), ('Ã­', 'i'),
        ('Ã³', 'o'), ('Ãº', 'u'), ('Ã¡', 'a'), ('Ã ', 'a'),
        ('Ã¢', 'a'), ('Ãª', 'e'), ('Ã´', 'o'), ('Ã»', 'u'),
        ('Ã•', 'O'), ('Ã‡', 'C'), ('Ã‰', 'E'), ('Ã"', 'O'),
        ('Ã€', 'A'), ('Ã‚', 'A'), ('Ã"', 'O'), ('Ãœ', 'U'),
        ('A£', 'a'), ('A§', 'c'), ('A©', 'e'), ('A­', 'i'),
        ('A³', 'o'), ('Aº', 'u'), ('A¡', 'a'), ('A¼', 'u'),
        ('A‡', 'C'), ('A‰', 'E'), ('Aœ', 'U'),
    ]
    for errado, correto in substituicoes:
        texto = texto.replace(errado, correto)

    # Remove acentos unicode restantes
    return remover_acentos(texto)


# ---------------------------------------------------------------------------
# Processamento do texto extraido pelo pdftotext
# ---------------------------------------------------------------------------

def _processar_texto_pdf(raw: str) -> str:
    """
    Transforma o texto bruto do pdftotext -layout em cifra formatada:

    Regras aplicadas:
      1. Slides separados por \\f viram blocos separados por linha em branco.
      2. Linhas de cifra sao identificadas pela funcao _e_linha_de_cifra().
      3. Apos a primeira cifra encontrada:
           - linhas de letra consecutivas (sem cifra no meio) sao unidas.
           - linhas em branco entre cifra e letra sao removidas.
      4. Multiplas linhas em branco consecutivas sao colapsadas em uma.
      5. Todos os acentos sao removidos.
    """
    entradas = []          # lista de ('tipo', 'texto')
    achou_primeira_cifra = False
    primeiro_slide = True

    for slide in raw.split("\f"):
        slide_txt = slide.strip("\n")
        if not slide_txt.strip():
            continue

        for linha in slide_txt.split("\n"):
            linha = linha.rstrip().rstrip("_").rstrip()

            if not linha.strip():
                entradas.append(("blank", ""))
            elif primeiro_slide:
                entradas.append(("title", linha))
            elif _e_linha_de_cifra(linha):
                entradas.append(("chord", linha))
                achou_primeira_cifra = True
            else:
                entradas.append(("lyric", linha))

        entradas.append(("blank", ""))   # separador de slide
        primeiro_slide = False

    # --- Passo 1: juntar linhas de letra consecutivas ---
    merged = []
    i = 0
    while i < len(entradas):
        tipo, txt = entradas[i]
        if achou_primeira_cifra and tipo == "lyric":
            combined = txt
            j = i + 1
            while j < len(entradas) and entradas[j][0] == "lyric":
                combined = combined.rstrip() + " " + entradas[j][1].strip()
                j += 1
            merged.append(("lyric", combined))
            i = j
        else:
            merged.append((tipo, txt))
            i += 1

    # --- Passo 2: remover blanks entre cifra e letra ---
    cleaned = []
    i = 0
    while i < len(merged):
        tipo, txt = merged[i]
        if tipo == "chord":
            cleaned.append((tipo, txt))
            j = i + 1
            while j < len(merged) and merged[j][0] == "blank":
                j += 1
            i = j
        else:
            cleaned.append((tipo, txt))
            i += 1

    # --- Passo 3: colapsar blanks multiplos ---
    result = []
    prev_blank = False
    for tipo, txt in cleaned:
        is_blank = tipo == "blank"
        if is_blank and prev_blank:
            continue
        result.append("" if is_blank else txt)
        prev_blank = is_blank

    # --- Passo 4: remover acentos ---
    result = [remover_acentos(l) for l in result]

    return "\n".join(result).strip()


# ---------------------------------------------------------------------------
# Conversao de um unico arquivo
# ---------------------------------------------------------------------------

def converter_arquivo(
    caminho_entrada: str | Path,
    caminho_saida: str | Path | None = None,
    pasta_saida: str | Path | None = None,
) -> Path:
    """
    Converte um arquivo .pptx ou .ppt para .txt com cifras formatadas.

    Parametros
    ----------
    caminho_entrada : str | Path
        Caminho do arquivo .pptx ou .ppt de entrada.
    caminho_saida : str | Path, opcional
        Caminho completo do arquivo .txt de saida.
        Se omitido, usa `pasta_saida` ou o mesmo diretorio da entrada.
    pasta_saida : str | Path, opcional
        Diretorio onde salvar o .txt (nome gerado automaticamente).

    Retorno
    -------
    Path
        Caminho do arquivo .txt gerado.

    Excecoes
    --------
    FileNotFoundError
        Se o arquivo de entrada nao existir.
    RuntimeError
        Se a conversao PDF ou extracao de texto falhar.
    """
    entrada = Path(caminho_entrada)
    if not entrada.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {entrada}")

    # Nome de saida sem acentos
    nome_limpo = nome_sem_acento(entrada.stem) + ".txt"

    if caminho_saida:
        saida = Path(caminho_saida)
    elif pasta_saida:
        saida = Path(pasta_saida) / nome_limpo
    else:
        saida = entrada.parent / nome_limpo

    saida.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmpdir:
        # 1. Converter para PDF via LibreOffice
        # Tenta o comando direto primeiro; se falhar, usa o caminho completo do Windows
        import sys
        if sys.platform == "win32":
            soffice = r"C:\Program Files\LibreOffice\program\soffice.exe"
        else:
            soffice = "libreoffice"

        r = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", tmpdir, str(entrada)],
            capture_output=True,
        )
        pdf = Path(tmpdir) / (entrada.stem + ".pdf")
        if not pdf.exists():
            stderr = r.stderr.decode(errors="replace")
            raise RuntimeError(
                f"LibreOffice nao gerou o PDF para '{entrada.name}'.\n{stderr}"
            )

        # 2. Extrair texto com layout preservado
        r = subprocess.run(
            ["pdftotext", "-layout", "-enc", "UTF-8", str(pdf), "-"],
            capture_output=True,
        )
        if r.returncode != 0:
            raise RuntimeError(f"pdftotext falhou: {r.stderr}")

        # Decodifica forçando UTF-8, ignorando erros
        raw = r.stdout.decode("utf-8", errors="replace")

        # Limpa sequencias corrompidas geradas por encoding errado do pdftotext
        raw = _limpar_encoding(raw)

        texto = _processar_texto_pdf(raw)

    saida.write_text(texto, encoding="utf-8")
    return saida


# ---------------------------------------------------------------------------
# Conversao de uma pasta inteira
# ---------------------------------------------------------------------------

def converter_pasta(
    pasta_entrada: str | Path,
    pasta_saida: str | Path | None = None,
    recursivo: bool = False,
    ignorar_duplicatas: bool = True,
) -> dict[str, Path | Exception]:
    """
    Converte todos os arquivos .pptx e .ppt de uma pasta para .txt.

    Parametros
    ----------
    pasta_entrada : str | Path
        Diretorio com os arquivos de entrada.
    pasta_saida : str | Path, opcional
        Diretorio de saida. Se omitido, usa o mesmo da entrada.
    recursivo : bool
        Se True, busca arquivos em subpastas tambem.
    ignorar_duplicatas : bool
        Se True e um .pptx e .ppt tiverem o mesmo nome base,
        processa apenas o .pptx.

    Retorno
    -------
    dict
        Mapeamento { nome_original -> Path(saida) } para sucessos,
        ou { nome_original -> Exception } para falhas.
    """
    entrada = Path(pasta_entrada)
    saida_base = Path(pasta_saida) if pasta_saida else entrada

    glob = "**/*" if recursivo else "*"
    pptx = sorted(entrada.glob(f"{glob}.pptx"))
    ppt  = sorted(entrada.glob(f"{glob}.ppt"))

    # Filtrar .ppt cujo nome base (sem acento) ja tem .pptx equivalente
    if ignorar_duplicatas:
        bases_pptx = {nome_sem_acento(f.stem) for f in pptx}
        ppt = [f for f in ppt if nome_sem_acento(f.stem) not in bases_pptx]

    arquivos = pptx + ppt
    resultados: dict[str, Path | Exception] = {}

    for arq in arquivos:
        try:
            destino = converter_arquivo(arq, pasta_saida=saida_base)
            resultados[arq.name] = destino
            print(f"  OK   {destino.name}")
        except Exception as e:
            resultados[arq.name] = e
            print(f"  ERRO {arq.name}: {e}")

    return resultados


# ---------------------------------------------------------------------------
# Interface de linha de comando
# ---------------------------------------------------------------------------

def _cli():
    import argparse

    parser = argparse.ArgumentParser(
        description="Converte .pptx/.ppt com cifras musicais para .txt",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Exemplos:
  python conversor_pptx.py musica.pptx
  python conversor_pptx.py musica.pptx --saida resultado.txt
  python conversor_pptx.py pasta_musicas/
  python conversor_pptx.py pasta_musicas/ --saida txts/ --recursivo
        """,
    )
    parser.add_argument("entrada", help="Arquivo .pptx/.ppt ou pasta de entrada")
    parser.add_argument("--saida", "-o", default=None,
                        help="Arquivo .txt de saida (para arquivo) ou pasta (para pasta)")
    parser.add_argument("--recursivo", "-r", action="store_true",
                        help="Buscar arquivos em subpastas (apenas para pasta)")
    args = parser.parse_args()

    entrada = Path(args.entrada)

    if entrada.is_dir():
        print(f"Convertendo pasta: {entrada}")
        resultados = converter_pasta(
            entrada,
            pasta_saida=args.saida,
            recursivo=args.recursivo,
        )
        ok  = sum(1 for v in resultados.values() if isinstance(v, Path))
        err = len(resultados) - ok
        print(f"\nConcluido: {ok} OK  |  {err} erro(s)  |  {len(resultados)} total")
    elif entrada.is_file():
        print(f"Convertendo: {entrada.name}")
        saida = converter_arquivo(entrada, caminho_saida=args.saida)
        print(f"Salvo em: {saida}")
    else:
        parser.error(f"Entrada nao encontrada: {entrada}")


if __name__ == "__main__":
    _cli()
