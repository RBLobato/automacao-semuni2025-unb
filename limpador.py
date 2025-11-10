import re
import unicodedata
import pandas as pd

ARQ_ENTRADA = "planilha_semuni.xlsx"
ARQ_SAIDA   = "planilha_semuni_limpa.xlsx"

# mapeie aqui os nomes exatos das colunas do seu Excel
COL_TITULO  = "Título do projeto"
COL_RESUMO  = "Resumo do projeto"
COL_COORD   = "Coordenador"       # opcional p/ dedupe mais estrito

# --- utilitários ---

SMALL_WORDS = {
    "a","o","as","os","de","da","do","das","dos","e","em","no","na","nos","nas",
    "por","para","com","sem","sob","sobre","entre","ao","aos","à","às","um","uma","uns","umas"
}

def remove_acentos(s: str) -> str:
    if not isinstance(s, str):
        return ""
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch))

def titlecase_pt(s: str) -> str:
    """Title Case com exceções PT; preserva palavra inicial e pós-pontuação."""
    if not isinstance(s, str): return ""
    s = re.sub(r"\s+", " ", s.strip())
    if not s: return s

    def cap_word(word: str) -> str:
        return word[:1].upper() + word[1:].lower()

    tokens = re.split(r"(\s+|-)", s)  # preserva separadores (espaço e hífen)
    out = []
    at_start = True
    for tok in tokens:
        if tok.strip() == "":  # separadores
            out.append(tok)
            continue
        # reinicia após pontuação forte
        if re.fullmatch(r"[.!?]+", tok):
            out.append(tok)
            at_start = True
            continue

        w = tok
        base = re.sub(r"[^\wÀ-ÖØ-öø-ÿ]", "", w, flags=re.UNICODE)
        small = base.lower() in SMALL_WORDS

        if at_start or not small:
            out.append(cap_word(w))
        else:
            out.append(w.lower())

        # se a palavra terminar com pontuação, próxima começa frase
        at_start = bool(re.search(r"[.!?]$", w))
    # garanta que a primeira “palavra” fique capitalizada
    out_str = "".join(out).strip()
    if out_str:
        out_str = out_str[0].upper() + out_str[1:]
    return out_str

def sentence_case_pt(s: str) -> str:
    """Primeira letra de cada frase maiúscula; resto minúsculo preservando siglas simples."""
    if not isinstance(s, str): return ""
    s = re.sub(r"\s+", " ", s.strip())
    if not s: return s

    # quebra por final de frase mantendo delimitadores
    parts = re.split(r"([.!?]\s*)", s)
    out = []
    for i in range(0, len(parts), 2):
        sent = parts[i]
        sep  = parts[i+1] if i+1 < len(parts) else ""
        sent = sent.lower()
        # capitalize 1ª letra alfabética
        m = re.search(r"[a-zá-úà-ùâ-ûãõä-üç]", sent, flags=re.I)
        if m:
            idx = m.start()
            sent = sent[:idx] + sent[idx].upper() + sent[idx+1:]
        out.append(sent + sep)
    return "".join(out).strip()

def normaliza_chave(s: str) -> str:
    """normaliza para deduplicação: sem acentos, minúsculo, espaços únicos."""
    s = str(s) if s is not None else ""
    s = remove_acentos(s)
    s = s.lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s

# --- processamento ---

df = pd.read_excel(ARQ_ENTRADA)

# Padronização de Título e Resumo
if COL_TITULO in df.columns:
    df[COL_TITULO] = df[COL_TITULO].fillna("").map(titlecase_pt)

if COL_RESUMO in df.columns:
    df[COL_RESUMO] = df[COL_RESUMO].fillna("").map(sentence_case_pt)

# Deduplicação de projetos
# chave baseada apenas no TÍTULO (robusta a acentos/caso/espaços)
df["__key_titulo__"] = df[COL_TITULO].map(normaliza_chave)

# Se quiser reforçar pelo coordenador também, ative a linha abaixo:
if COL_COORD in df.columns:
    df["__key_coord__"] = df[COL_COORD].fillna("").map(normaliza_chave)
    df["__dedup_key__"] = df["__key_titulo__"] + " | " + df["__key_coord__"]
else:
    df["__dedup_key__"] = df["__key_titulo__"]

antes = len(df)
df = df.drop_duplicates(subset="__dedup_key__", keep="first").copy()
depois = len(df)
removidos = antes - depois

# limpar colunas técnicas
df = df.drop(columns=[c for c in df.columns if c.startswith("__")], errors="ignore")

# salvar
df.to_excel(ARQ_SAIDA, index=False)
print(f"✅ Planilha limpa salva em '{ARQ_SAIDA}'. Removidos {removidos} duplicados (de {antes} → {depois}).")
