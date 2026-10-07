#!/usr/bin/env python3
"""
Devenir des docteurs formés en France, d'après OpenAlex — script unique.

Il fait tout, dans l'ordre :
  1. collecte des thèses (DOI 10.70675, type dissertation) et des profils auteurs OpenAlex,
     avec un cache disque : rien n'est re-téléchargé s'il est déjà en cache ;
  2. statut de chaque docteur à T+3, T+5, T+8 (France / étranger / entreprise / sans trace) ;
  3. trajectoires année par année et allers-retours (partis à l'étranger puis revenus) ;
  4. comparaison des thèses des docteurs partis et restés (topics, mots-clés, langue)
     et taux de départ par établissement — collecte complémentaire des métadonnées des thèses ;
  5. fichiers de résultats (CSV, Excel) et tableau de bord HTML autonome, messages compris.

Installation
    pip install httpx pandas openpyxl

Utilisation
    export OPENALEX_API_KEY=ta_cle            # Windows : set OPENALEX_API_KEY=ta_cle
    python devenir_docteurs.py                         # tout, de 2006 à l'an dernier
    python devenir_docteurs.py --from 2010 --to 2020   # une période
    python devenir_docteurs.py --sample 500            # test rapide : 500 thèses par an
    python devenir_docteurs.py --offline               # aucune requête : recalcul depuis le cache
    python devenir_docteurs.py --refresh               # re-télécharge tout le cache (automatique après 180 jours)

Résultats (dossier --out, « resultats » par défaut)
    tableau_de_bord.html     tableau de bord complet, à ouvrir dans un navigateur
    docteurs_T{k}.csv        une ligne par thèse et par horizon
    synthese.xlsx            tableaux par discipline, cohorte, pays, trajectoires, allers-retours
    partants_restes.xlsx     topics, mots-clés, établissements

Coût API (clé gratuite : 10 000 requêtes « liste » par jour) : environ 1 requête par 100 thèses,
autant pour les auteurs, et 1 par 100 thèses pour les métadonnées des partis et restés.
Une collecte interrompue reprend là où elle s'est arrêtée.
"""
from __future__ import annotations

import argparse, asyncio, calendar, collections, datetime, html, json, math, os, random, re, sys, time
from pathlib import Path

try:
    import httpx
    import pandas as pd
except ImportError:
    sys.exit("Il manque des bibliothèques : pip install httpx pandas openpyxl")

API = "https://api.openalex.org"
BASE_FILTER = "doi_starts_with:10.70675,type:dissertation"
DOM = {"GP", "MQ", "GF", "RE", "YT", "NC", "PF", "PM", "WF", "BL", "MF"}   # outre-mer = France

DISC = [
    ("Mathématiques", [26, 18]),
    ("Informatique", [17]),
    ("Physique & astronomie", [31]),
    ("Chimie & matériaux", [16, 15, 25]),
    ("Ingénierie & énergie", [22, 21]),
    ("Terre & environnement", [19, 23]),
    ("Biologie", [11, 13, 24, 28]),
    ("Médecine & santé", [27, 29, 30, 34, 35, 36]),
    ("Économie & gestion", [14, 20]),
    ("Sciences sociales & psychologie", [33, 32]),
    ("Arts & humanités", [12]),
]
FIELD2DISC = {f: n for n, fs in DISC for f in fs}
NC = "Non classé"
STATUTS = ["Académique France", "Entreprise France", "Académique étranger",
           "Entreprise étranger", "Trace interrompue", "Pas de trace après la thèse"]
TYP = ["Resté en France", "Aller-retour", "Installé à l'étranger", "Départ après la France",
       "Parti, puis plus de trace", "Retour au pays d'origine", "Pas de trace"]
EUROPE = set("GB DE CH IT BE SE NL ES LU NO AT PL DK CZ PT IE FI GR HU RO SK SI HR BG EE LV LT IS MT CY RS UA RU BA MK AL MD BY ME XK LI MC AD".split())
NAMERICA = {"US", "CA"}
INST_MERGE = {
    "Université Pierre-et-Marie-Curie": "Sorbonne Université", "Paris-Sorbonne University": "Sorbonne Université",
    "Université Paris-Sorbonne": "Sorbonne Université", "Université Paris-Sud": "Université Paris-Saclay",
    "Université Paris Descartes": "Université Paris Cité", "Université Paris Diderot": "Université Paris Cité",
    "Université Lille 1": "Université de Lille", "Université Lille 2": "Université de Lille", "Université Lille 3": "Université de Lille",
    "Université Bordeaux-I": "Université de Bordeaux", "Université Bordeaux Segalen": "Université de Bordeaux",
    "Université Montpellier 1": "Université de Montpellier", "Université Montpellier 2": "Université de Montpellier",
    "Université Blaise Pascal": "Université Clermont Auvergne", "Université d'Auvergne": "Université Clermont Auvergne",
    "Communauté Université Grenoble Alpes": "Université Grenoble Alpes", "Université Joseph Fourier": "Université Grenoble Alpes",
    "Communauté d'universités et d'établissements Université Côte d'Azur": "Université Côte d'Azur",
    "Université Nice Sophia Antipolis": "Université Côte d'Azur",
    "Lyon 1 Université": "Université Claude Bernard Lyon 1", "Sorbonne Paris Cité": "Sorbonne Paris Cité (COMUE)",
    "Institut d'Etudes Politiques de Paris": "Sciences Po", "Université Paris Sciences et Lettres": "Université PSL",
    "Université Paul-Valéry Montpellier": "Université Paul-Valéry Montpellier 3",
    "Université Toulouse III - Paul Sabatier": "Université Toulouse III Paul-Sabatier",
    "Université Toulouse-I-Capitole": "Université Toulouse Capitole",
}
COUNTRY_FR = {"US": "États-Unis", "GB": "Royaume-Uni", "DE": "Allemagne", "CH": "Suisse", "CA": "Canada", "BE": "Belgique",
              "IT": "Italie", "ES": "Espagne", "NL": "Pays-Bas", "CN": "Chine", "BR": "Brésil", "TN": "Tunisie", "MA": "Maroc",
              "DZ": "Algérie", "LB": "Liban", "SE": "Suède", "JP": "Japon", "AU": "Australie", "VN": "Viêt Nam", "LU": "Luxembourg",
              "PT": "Portugal", "AT": "Autriche", "DK": "Danemark", "NO": "Norvège", "SN": "Sénégal", "IN": "Inde", "IR": "Iran",
              "CL": "Chili", "MX": "Mexique", "CO": "Colombie", "PK": "Pakistan", "SA": "Arabie saoudite", "KR": "Corée du Sud"}


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, file=sys.stderr, flush=True)


def sid(u):
    return str(u).rsplit("/", 1)[-1] if u else None


def fr(x, d=0):
    """Pourcentage ou nombre au format français."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "–"
    s = f"{x:,.{d}f}".replace(",", "\u202f").replace(".", ",").replace("-", "−")
    return s


def pc(x, d=0):
    return fr(x, d) + " %"


def cn(c):
    return COUNTRY_FR.get(c, c)


def join_fr(items):
    items = [i for i in items if i]
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + " et " + items[-1] if items else ""


# ====================================================================== API
class ApiError(Exception):
    def __init__(self, status, msg):
        super().__init__(f"HTTP {status} : {msg}")
        self.status = status


class OpenAlex:
    def __init__(self, key, mailto, concurrency=10, transport=None):
        self.base = {k: v for k, v in (("api_key", key), ("mailto", mailto)) if v}
        self.sem = asyncio.Semaphore(concurrency)
        self.client = httpx.AsyncClient(base_url=API, timeout=httpx.Timeout(60, connect=15), transport=transport,
                                        limits=httpx.Limits(max_connections=concurrency, max_keepalive_connections=concurrency))
        self.calls, self.t0 = 0, time.time()

    async def close(self):
        await self.client.aclose()

    async def get(self, path, params=None, tries=8):
        params = {**(params or {}), **self.base}
        err = None
        for i in range(tries):
            async with self.sem:
                try:
                    r = await self.client.get(path, params=params)
                except httpx.TransportError as e:
                    err, r = e, None
            if r is None:
                await asyncio.sleep(min(30, 2 ** i) + random.random()); continue
            self.calls += 1
            if r.status_code == 200:
                return r.json()
            if r.status_code in (429, 500, 502, 503, 504):
                if r.status_code == 429 and i >= 4:
                    raise ApiError(429, "quota OpenAlex atteint (clé absente ou quota du jour épuisé). " + r.text[:200])
                ra = r.headers.get("retry-after")
                await asyncio.sleep(float(ra) if ra and ra.replace(".", "").isdigit() else min(60, 2 ** i) + random.random())
                continue
            try:
                msg = r.json().get("message") or r.text
            except Exception:
                msg = r.text
            raise ApiError(r.status_code, str(msg)[:300])
        raise ApiError(0, f"échec réseau répété ({err})")

    def rate(self):
        return self.calls / max(1e-6, time.time() - self.t0)


W_SELECT = "id,doi,publication_year,authorships,primary_topic"
A_SELECT = "id,orcid,display_name,works_count,affiliations"
M_SELECT = "id,language,keywords,topics,authorships"


def compact_work(w):
    au_list = w.get("authorships") or []
    a = next((x for x in au_list if x.get("author_position") == "first"), au_list[0] if au_list else {})
    au = a.get("author") or {}
    f = ((w.get("primary_topic") or {}).get("field") or {})
    return {"work": sid(w["id"]), "doi": (w.get("doi") or "").replace("https://doi.org/", ""),
            "year": w.get("publication_year"), "author": sid(au.get("id")),
            "name": au.get("display_name") or a.get("raw_author_name"), "orcid": sid(au.get("orcid")),
            "field": int(sid(f["id"])) if f.get("id") else None, "field_name": f.get("display_name")}


def compact_author(a):
    return {"id": sid(a["id"]), "orcid": sid(a.get("orcid")), "name": a.get("display_name"),
            "works_count": a.get("works_count"),
            "aff": [{"n": (x.get("institution") or {}).get("display_name"),
                     "c": (x.get("institution") or {}).get("country_code"),
                     "t": (x.get("institution") or {}).get("type"),
                     "ys": x.get("years") or []} for x in a.get("affiliations") or []]}


def compact_meta(w):
    a = (w.get("authorships") or [{}])[0]
    insts = a.get("institutions") or []
    kws = [k for k in (w.get("keywords") or []) if k.get("display_name")]
    return {"work": sid(w["id"]), "lang": w.get("language"),
            "kw": [k["display_name"] for k in kws], "kw_score": [round(k.get("score") or 0, 3) for k in kws],
            "topics": [t["display_name"] for t in (w.get("topics") or [])][:3],
            "inst": [i.get("display_name") for i in insts if i.get("display_name")],
            "inst_cc": [i.get("country_code") for i in insts]}


def read_jsonl(f):
    out = []
    if f.exists():
        for l in open(f, encoding="utf-8"):
            l = l.strip()
            if l:
                try:
                    out.append(json.loads(l))
                except json.JSONDecodeError:
                    pass          # ligne tronquée par une interruption
    return out


def write_jsonl(f, rows, mode="w"):
    with open(f, mode, encoding="utf-8") as fh:
        for r in rows:
            fh.write(json.dumps(r, ensure_ascii=False) + "\n")


# ---------------------------------------------------------------- thèses
async def count_by_year(oa, flt, y1, y2):
    d = await oa.get("/works", {"filter": f"{flt},publication_year:{y1}-{y2}", "group_by": "publication_year"})
    return {int(g["key"]): g["count"] for g in d.get("group_by", [])}


async def cursor_chain(oa, flt):
    out, cursor = [], "*"
    while cursor:
        d = await oa.get("/works", {"filter": flt, "select": W_SELECT, "per_page": 100, "cursor": cursor})
        out += [compact_work(w) for w in d.get("results", [])]
        cursor = d.get("meta", {}).get("next_cursor")
        if not d.get("results"):
            break
    return out


async def shard_all(oa, flt, y, n_year, cache, progress):
    f = cache / f"all_{y}.jsonl"
    if f.exists():
        rows = read_jsonl(f); progress(len(rows)); return rows
    shards = [f"{flt},publication_year:{y}"] if n_year <= 1000 else \
        [f"{flt},from_publication_date:{y}-{m:02d}-01,to_publication_date:{y}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}" for m in range(1, 13)]

    async def run(s):
        r = await cursor_chain(oa, s); progress(len(r)); return r
    parts = await asyncio.gather(*(run(s) for s in shards))
    rows = list({r["work"]: r for p in parts for r in p}.values())
    if len(rows) < n_year:
        log(f"  {y} : {len(rows)}/{n_year} par mois, complément par chaîne annuelle")
        rows = list({r["work"]: r for r in rows + await cursor_chain(oa, f"{flt},publication_year:{y}")}.values())
    write_jsonl(f, rows)
    return rows


async def shard_sample(oa, flt, y, n, seed, cache, progress):
    f = cache / f"sample_{y}_n{n}_s{seed}.jsonl"
    if f.exists():
        rows = read_jsonl(f); progress(len(rows)); return rows
    base = {"filter": f"{flt},publication_year:{y}", "select": W_SELECT, "per_page": 100, "sample": n, "seed": seed}
    try:
        res = await asyncio.gather(*(oa.get("/works", {**base, "page": p}) for p in range(1, (n + 99) // 100 + 1)))
        rows = [compact_work(w) for d in res for w in d.get("results", [])]
    except ApiError as e:
        if e.status != 400:
            raise
        rows, cursor = [], "*"
        while cursor and len(rows) < n:
            d = await oa.get("/works", {**base, "cursor": cursor})
            rows += [compact_work(w) for w in d.get("results", [])]
            cursor = d.get("meta", {}).get("next_cursor")
            if not d.get("results"):
                break
    rows = list({r["work"]: r for r in rows}.values())[:n]
    progress(len(rows)); write_jsonl(f, rows)
    return rows


# ---------------------------------------------------------------- auteurs
async def fetch_authors(oa, ids, cache):
    fa, fal = cache / "authors.jsonl", cache / "alias.json"
    authors = {a["id"]: a for a in read_jsonl(fa)}
    alias = json.loads(fal.read_text()) if fal.exists() else {}
    todo = [i for i in ids if i not in authors and i not in alias]
    log(f"Auteurs : {len(ids)} au total, {len(ids) - len(todo)} en cache, {len(todo)} à récupérer")
    if not todo:
        return authors, alias
    fh = open(fa, "a", encoding="utf-8"); done = 0

    async def batch(b):
        nonlocal done
        d = await oa.get("/authors", {"filter": "ids.openalex:" + "|".join(b), "select": A_SELECT, "per_page": 100})
        for a in d.get("results", []):
            c = compact_author(a); authors[c["id"]] = c
            fh.write(json.dumps(c, ensure_ascii=False) + "\n")
        done += len(b)
        if done % 2000 < 100 or done == len(todo):
            log(f"  auteurs {done}/{len(todo)} · {oa.rate():.1f} req/s")
    try:
        await asyncio.gather(*(batch(todo[i:i + 100]) for i in range(0, len(todo), 100)))
        missing = [i for i in todo if i not in authors]
        if missing:
            log(f"  {len(missing)} profils fusionnés : récupération unitaire")

            async def one(i):
                try:
                    a = await oa.get(f"/authors/{i}", {"select": A_SELECT})
                except ApiError as e:
                    if e.status == 404:
                        alias[i] = None; return
                    raise
                c = compact_author(a); alias[i] = c["id"]
                if c["id"] not in authors:
                    authors[c["id"]] = c; fh.write(json.dumps(c, ensure_ascii=False) + "\n")
            await asyncio.gather(*(one(i) for i in missing))
    finally:
        fh.close(); fal.write_text(json.dumps(alias))
    return authors, alias


# ---------------------------------------------------------------- métadonnées des thèses
async def fetch_meta(oa, ids, f):
    have = {m["work"]: m for m in read_jsonl(f)}
    todo = [i for i in ids if i not in have]
    log(f"Métadonnées des thèses : {len(ids)} utiles, {len(ids) - len(todo)} en cache, {len(todo)} à récupérer")
    if not todo:
        return have
    fh = open(f, "a", encoding="utf-8"); fkey = ["ids.openalex"]; done = 0

    async def batch(b):
        nonlocal done
        try:
            d = await oa.get("/works", {"filter": f"{fkey[0]}:" + "|".join(b), "select": M_SELECT, "per_page": 100})
        except ApiError as e:
            if e.status == 400 and fkey[0] == "ids.openalex":
                fkey[0] = "openalex"
                d = await oa.get("/works", {"filter": "openalex:" + "|".join(b), "select": M_SELECT, "per_page": 100})
            else:
                raise
        for w in d.get("results", []):
            c = compact_meta(w); have[c["work"]] = c
            fh.write(json.dumps(c, ensure_ascii=False) + "\n")
        done += len(b)
        if done % 5000 < 100 or done == len(todo):
            log(f"  métadonnées {done}/{len(todo)}")
    try:
        await asyncio.gather(*(batch(todo[i:i + 100]) for i in range(0, len(todo), 100)))
    finally:
        fh.close()
    return have


def cache_freshness(P, cache, root):
    """Rafraîchit tout le cache quand il est demandé (--refresh) ou trop vieux (--max-age).
    Une mise à jour interrompue reprend là où elle s'est arrêtée au lancement suivant."""
    f = cache / "_collecte.json"
    info = json.loads(f.read_text()) if f.exists() else {}
    if info.get("refresh_en_cours"):
        if not P.offline:
            log(f"Reprise de la mise à jour du cache commencée le {info['refresh_en_cours']}.")
        return
    last = datetime.date.fromisoformat(info["date"]) if info.get("date") else None
    if last is None and (cache / "authors.jsonl").exists():
        last = datetime.date.fromtimestamp((cache / "authors.jsonl").stat().st_mtime)
    age = (datetime.date.today() - last).days if last else None
    if P.offline:
        if last and P.last_obs >= last.year:
            log(f"Attention : le cache date du {last:%d/%m/%Y}. Les affiliations de {P.last_obs} y sont incomplètes ; "
                f"relance sans --offline pour le mettre à jour, ou utilise --last-obs {last.year - 1}.")
        return
    if not (P.refresh or (age is not None and age > P.max_age)):
        if age is not None:
            log(f"Cache du {last:%d/%m/%Y} ({age} jours) réutilisé ; --refresh pour tout re-télécharger.")
        return
    why = "demandée (--refresh)" if P.refresh else f"cache vieux de {age} jours (> {P.max_age})"
    log(f"Mise à jour complète du cache : {why}. Thèses, profils auteurs et métadonnées sont re-téléchargés.")
    for g in cache.iterdir():
        if g.is_file():
            g.unlink()
    (root / "theses_meta.jsonl").unlink(missing_ok=True)
    f.write_text(json.dumps({"refresh_en_cours": datetime.date.today().isoformat()}))


def mark_fresh(cache):
    (cache / "_collecte.json").write_text(json.dumps({"date": datetime.date.today().isoformat()}))


async def collect(P, cache, transport=None):
    oa = OpenAlex(P.key, P.mailto, P.concurrency, transport)
    try:
        counts = await count_by_year(oa, P.filter, P.y1, P.y2)
        years = [y for y in range(P.y1, P.y2 + 1) if counts.get(y)]
        want = sum(min(counts[y], P.sample) if P.sample else counts[y] for y in years)
        log(f"{sum(counts.values())} thèses sur {P.y1}-{P.y2} ; {want} retenues")
        got = 0

        def progress(n):
            nonlocal got
            got += n
            log(f"  thèses {got}/{want} · {oa.calls} requêtes")
        if P.sample:
            parts = await asyncio.gather(*(shard_sample(oa, P.filter, y, min(P.sample, counts[y]), P.seed, cache, progress) for y in years))
        else:
            parts = await asyncio.gather(*(shard_all(oa, P.filter, y, counts[y], cache, progress) for y in years))
        theses = [r for p in parts for r in p]
        authors, alias = await fetch_authors(oa, sorted({t["author"] for t in theses if t["author"]}), cache)
        log(f"Collecte principale : {oa.calls} requêtes en {time.time() - oa.t0:.0f} s")
        return theses, authors, alias
    finally:
        await oa.close()


def load_cache(P, cache):
    theses = []
    for y in range(P.y1, P.y2 + 1):
        files = sorted(cache.glob(f"sample_{y}_n*_s{P.seed}.jsonl")) if P.sample else [cache / f"all_{y}.jsonl"]
        for f in files:
            theses += read_jsonl(f)
    authors = {a["id"]: a for a in read_jsonl(cache / "authors.jsonl")}
    alias = json.loads((cache / "alias.json").read_text()) if (cache / "alias.json").exists() else {}
    return theses, authors, alias


# ====================================================================== classement
class Prep:
    """Affiliations d'un auteur par année ; outre-mer ramené à la France."""
    __slots__ = ("by_year", "first", "seen")

    def __init__(self, a):
        self.by_year = collections.defaultdict(list)
        self.seen = {}
        first = 9999
        for x in a["aff"]:
            c = "FR" if x["c"] in DOM else x["c"]
            for y in x["ys"]:
                self.by_year[y].append((c, x["t"]))
                first = min(first, y)
                if c and y < self.seen.get(c, 9999):
                    self.seen[c] = y
        self.first = first


def pick_country(insts, multi):
    cc = collections.Counter(c for c, _ in insts if c)
    if not cc:
        return None
    if multi == "fr":
        return "FR" if "FR" in cc else cc.most_common(1)[0][0]
    if multi == "abroad":
        ab = [x for x, _ in cc.most_common() if x != "FR"]
        return ab[0] if ab else "FR"
    return max(cc, key=lambda x: (cc[x], x == "FR"))


def classify(T, p, k, P, n_theses=1, orcid=None):
    """-> (indice de statut 0-5 | None, exclusion | None, pays, retour, année retenue)"""
    if T is None or T + k > P.last_obs:
        return None, "censuré", None, False, None
    if p is None:
        return None, "sans profil", None, False, None
    if P.orcid_only and not orcid:
        return None, "sans ORCID", None, False, None
    if not P.keep_suspect and (n_theses > 1 or p.first < T - P.max_pre):
        return None, "profil douteux", None, False, None
    lo, hi = T + P.lag + 1, T + k
    ys = [y for y in p.by_year if lo <= y <= hi]
    if not ys:
        return 5, None, None, False, None
    y = max(ys)
    if y < hi - (P.tol - 1):
        return 4, None, None, False, y
    insts = p.by_year[y]
    c = pick_country(insts, P.multi)
    if c is None:
        return None, "affiliation sans pays", None, False, y
    types = [t for cx, t in insts if cx == c]
    priv = (not P.no_private) and types and all(t == "company" for t in types)
    if c == "FR":
        return (1 if priv else 0), None, c, False, y
    return (3 if priv else 2), None, c, p.seen.get(c, 9999) <= T - P.ret_window, y


def year_loc(p, y, P):
    insts = p.by_year.get(y)
    if not insts:
        return None, None
    c = pick_country(insts, P.multi)
    if c is None:
        return None, None
    return ("F", c) if c == "FR" else ("E", c)


# ====================================================================== statistiques
def fightin_words(docs_a, docs_b, min_docs=10):
    ca = collections.Counter(t for d in docs_a for t in d)
    cb = collections.Counter(t for d in docs_b for t in d)
    terms = [t for t in set(ca) | set(cb) if ca[t] + cb[t] >= min_docs]
    cols = ["terme", "z", "part_partis_%", "part_restes_%", "n_partis", "n_restes"]
    if not terms:
        return pd.DataFrame(columns=cols)
    na, nb = sum(ca.values()), sum(cb.values())
    tot = na + nb
    a0 = max(100.0, 0.1 * tot)
    rows = []
    for t in terms:
        ai = a0 * (ca[t] + cb[t]) / tot
        da, db = na + a0 - ca[t] - ai, nb + a0 - cb[t] - ai
        if da <= 0 or db <= 0:          # terme présent partout : non informatif
            continue
        la = math.log((ca[t] + ai) / da)
        lb = math.log((cb[t] + ai) / db)
        z = (la - lb) / math.sqrt(1 / (ca[t] + ai) + 1 / (cb[t] + ai))
        rows.append((t, z, 100 * ca[t] / max(1, len(docs_a)), 100 * cb[t] / max(1, len(docs_b)), ca[t], cb[t]))
    return pd.DataFrame(rows, columns=cols).sort_values("z", ascending=False)


def wilson(k, n, z=1.96):
    if n == 0:
        return 0, 0
    p = k / n; d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d; h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return 100 * max(0, c - h), 100 * min(1, c + h)


def poisson_ratio_ci(o, e):
    """Intervalle à 95 % du ratio observé/attendu (approximation de Byar)."""
    if e <= 0:
        return 0, 0
    lo = 0 if o == 0 else o * (1 - 1 / (9 * o) - 1.96 / (3 * math.sqrt(o))) ** 3
    hi = (o + 1) * (1 - 1 / (9 * (o + 1)) + 1.96 / (3 * math.sqrt(o + 1))) ** 3
    return lo / e, hi / e


# ====================================================================== analyse principale
class Data:
    """Tout ce qui est calculé, partagé entre exports, messages et tableau de bord."""


def analyse_core(theses, authors, alias, P):
    A = Data()
    theses = [t for t in theses if t.get("year")]
    prep = {}

    def resolve(aid):
        if not aid:
            return None, None
        real = aid if aid in authors else alias.get(aid)
        if not real or real not in authors:
            return None, None
        if real not in prep:
            prep[real] = Prep(authors[real])
        return real, prep[real]
    res = [resolve(t["author"]) for t in theses]
    n_by = collections.Counter(r for r, _ in res if r)
    orc = [(authors.get(r, {}).get("orcid") if r else None) or t.get("orcid") for t, (r, _) in zip(theses, res)]
    disc = [FIELD2DISC.get(t.get("field"), NC) for t in theses]
    present = [n for n, _ in DISC if n in set(disc)] + ([NC] if NC in disc else [])
    A.theses, A.res, A.n_by, A.orc, A.disc_of, A.disc = theses, res, n_by, orc, disc, present
    A.didx = {d: i for i, d in enumerate(present)}
    years = sorted({t["year"] for t in theses})
    A.years = list(range(years[0], years[-1] + 1))
    A.yidx = {y: i for i, y in enumerate(A.years)}
    A.corpus = collections.Counter(t["year"] for t in theses)
    mx = max(A.corpus.values())
    A.minY = next(y for y in A.years if A.corpus.get(y, 0) >= 0.25 * mx)

    # ---- statut à chaque horizon
    A.cls = {}
    for k in P.horizons:
        A.cls[k] = [classify(t["year"], p, k, P, n_by.get(r, 1), o) for t, (r, p), o in zip(theses, res, orc)]
    return A


def per_thesis_frames(A, P, out):
    frames = {}
    for k in P.horizons:
        rows = []
        for t, (r, _), o, d, (st, ex, c, ret, ly) in zip(A.theses, A.res, A.orc, A.disc_of, A.cls[k]):
            lab = None if st is None else STATUTS[st]
            if lab and P.no_private:
                lab = {0: "France", 1: "France", 2: "Étranger", 3: "Étranger"}.get(st, lab)
            rows.append((t["work"], t["doi"], t["year"], d, t.get("field_name"), r or t["author"], t.get("name"), o,
                         lab, ex, c, ret, ly))
        df = pd.DataFrame(rows, columns=["work", "doi", "annee", "discipline", "field_openalex", "auteur", "nom", "orcid",
                                         "statut", "exclusion", "pays", "retour", "derniere_annee_affil"])
        df.to_csv(out / f"docteurs_T{k}.csv", index=False, sep=";", encoding="utf-8-sig")
        frames[k] = df
    return frames


def dashboard_data(A, P):
    D = {"disc": A.disc, "years": A.years, "st": STATUTS, "horizons": P.horizons, "mainH": P.main_h,
         "lastObs": P.last_obs, "minY": A.minY, "cube": {}, "ctry": {}, "excl": {}, "tot": {},
         "corpus": {str(y): A.corpus.get(y, 0) for y in A.years}, "lag": P.lag, "earlyH": P.early}
    for k in P.horizons:
        cube = [[[[0] * 6 for _ in range(2)] for _ in A.years] for _ in A.disc]
        ctry = collections.defaultdict(lambda: [0, 0])
        excl = collections.Counter()
        for t, o, d, (st, ex, c, ret, _) in zip(A.theses, A.orc, A.disc_of, A.cls[k]):
            if ex:
                excl[ex] += 1; continue
            di, yi, oi = A.didx[d], A.yidx[t["year"]], int(bool(o))
            cube[di][yi][oi][st] += 1
            if st in (2, 3):
                e = ctry[(c, di, yi, oi)]; e[0] += 1; e[1] += int(bool(ret))
        D["cube"][str(k)] = cube
        D["ctry"][str(k)] = [[c, di, yi, oi, n, r] for (c, di, yi, oi), (n, r) in ctry.items()]
        D["excl"][str(k)] = dict(excl)
        D["tot"][str(k)] = len(A.theses)
    # trajectoire T+lag+1 .. T+10
    J = list(range(P.lag + 1, 11))
    acc = collections.defaultdict(lambda: {j: [0, 0, 0, 0] for j in J})
    for t, (r, p), o, d in zip(A.theses, A.res, A.orc, A.disc_of):
        for j in J:
            Pj = argparse.Namespace(**{**vars(P), "tol": min(P.tol, j - P.lag)})
            st, ex, *_ = classify(t["year"], p, j, Pj, A.n_by.get(r, 1), o)
            if ex:
                continue
            for key in (d, "Ensemble"):
                a = acc[key][j]; a[0] += 1
                if st in (0, 1):
                    a[1] += 1; a[2] += 1
                elif st in (2, 3):
                    a[1] += 1; a[3] += 1
    traj = {"_j": J}
    for key, by in acc.items():
        traj[key] = {"observes": [by[j][0] for j in J],
                     "visibles_%": [round(100 * by[j][1] / by[j][0], 1) if by[j][0] else None for j in J],
                     "en_france_%": [round(100 * by[j][2] / by[j][0], 1) if by[j][0] else None for j in J],
                     "etranger_%": [round(100 * by[j][3] / by[j][0], 1) if by[j][0] else None for j in J]}
    D["traj"] = traj
    return D


# ====================================================================== allers-retours (année par année)
def traj_type(seq, st_end):
    """Type de trajectoire à partir de la séquence annuelle [(année, 'F'|'E', pays, retour_pays)] (années avec affiliation)."""
    locs = [x[1] for x in seq]
    if not locs:
        return 6
    if "E" not in locs:
        return 0
    fe = locs.index("E")
    if locs[-1] == "F":
        return 1
    if "F" in locs[:fe]:
        return 3
    if all(x[3] for x in seq if x[1] == "E"):
        return 5
    return 2 if st_end in (2, 3) else 4


def aller_retour(A, P):
    H = P.ar_h
    y1 = P.last_obs - H
    years = [y for y in A.years if A.minY <= y <= y1]
    AR = {"typ": TYP, "years": years, "cube": None, "early": None, "dur": [], "ctry": []}
    if not years:
        return AR, pd.DataFrame()
    yi = {y: i for i, y in enumerate(years)}
    cube = [[[[0] * 7 for _ in range(2)] for _ in years] for _ in A.disc]
    early = [[[[0] * 3 for _ in range(2)] for _ in years] for _ in A.disc]
    dur = collections.Counter(); ctry = collections.Counter(); rows = []
    for t, (r, p), o, d, (stH, exH, *_r) in zip(A.theses, A.res, A.orc, A.disc_of, A.cls[H]):
        T = t["year"]
        if T not in yi or exH:
            continue
        di, ci, oi = A.didx[d], yi[T], int(bool(o))
        seq = []
        for y in range(T + P.lag + 1, T + H + 1):
            loc, c = year_loc(p, y, P)
            if loc:
                home = loc == "E" and p.seen.get(c, 9999) <= T - P.ret_window
                seq.append((y, loc, c, home))
        typ = traj_type(seq, stH)
        if typ == 1:
            locs = [x[1] for x in seq]; fe = locs.index("E")
            ye, ce = seq[fe][0], seq[fe][2]
            yf = next(x[0] for x in seq[fe:] if x[1] == "F")
            k = min(max(yf - ye, 1), 5) - 1
            dur[(di, ci, oi, k)] += 1; ctry[(ce, di, ci, oi)] += 1
        cube[di][ci][oi][typ] += 1
        is_early = any(s[1] == "E" and not s[3] and s[0] <= T + P.early for s in seq)
        if is_early:
            early[di][ci][oi][0 if stH in (0, 1) else 1 if stH in (2, 3) else 2] += 1
        rows.append((t["work"], T, d, TYP[typ], "".join("F" if s[1] == "F" else "E" for s in seq), int(is_early)))
    AR["cube"], AR["early"] = cube, early
    AR["dur"] = [[di, ci, oi, k, n] for (di, ci, oi, k), n in dur.items()]
    AR["ctry"] = [[c, di, ci, oi, n] for (c, di, ci, oi), n in ctry.items()]
    return AR, pd.DataFrame(rows, columns=["work", "annee", "discipline", "trajectoire", "sequence_annuelle", "parti_tot"])


# ====================================================================== liste nominative (onglet Explorer)
ST_CODE = "FfEeIN"


def explorer_data(A, P, meta):
    """Une ligne par docteur, en colonnes compactes, pour l'onglet Explorer du tableau de bord."""
    hs = sorted(P.horizons)
    vt, vk = {}, {}
    cols = {k: [] for k in ("n", "a", "w", "y", "d", "s", "q", "c", "h", "t", "tp", "kw")}
    with_meta = 0
    for i, (t, (r, p), o, d) in enumerate(zip(A.theses, A.res, A.orc, A.disc_of)):
        if p is None:
            continue
        T = t["year"]
        if not P.keep_suspect and (A.n_by.get(r, 1) > 1 or p.first < T - P.max_pre):
            continue
        y1 = min(T + 10, P.last_obs)
        if y1 < T + P.lag + 1:
            continue
        seq = []
        for y in range(T + P.lag + 1, y1 + 1):
            loc, c = year_loc(p, y, P)
            seq.append((y, loc, c, loc == "E" and p.seen.get(c, 9999) <= T - P.ret_window))
        located = [x for x in seq if x[1]]
        abroad = [x for x in located if x[1] == "E"]
        if P.explorer == "mobiles" and not abroad:
            continue
        if P.explorer == "visibles" and not located:
            continue
        st_end = classify(T, p, y1 - T, P, A.n_by.get(r, 1), o)[0]
        typ = traj_type(located, st_end)
        status = "".join("-" if A.cls[k][i][1] == "censuré" else "X" if A.cls[k][i][1] else ST_CODE[A.cls[k][i][0]] for k in hs)
        ctrs = list(dict.fromkeys(x[2] for x in abroad))
        homes = list(dict.fromkeys(x[2] for x in abroad if x[3]))
        m = meta.get(t["work"]) if meta else None
        tps = [vt.setdefault(x, len(vt)) for x in (m["topics"] if m else [])]
        kws = [vk.setdefault(x, len(vk)) for x, sc in (zip(m["kw"], m["kw_score"]) if m else []) if sc >= P.min_score]
        with_meta += bool(m)
        cols["n"].append(t.get("name") or "")
        cols["a"].append(r); cols["w"].append(t["work"]); cols["y"].append(T); cols["d"].append(A.didx[d])
        cols["s"].append(status); cols["q"].append("".join(("O" if x[3] else x[1]) if x[1] else "." for x in seq))
        cols["c"].append(",".join(ctrs)); cols["h"].append(",".join(homes)); cols["t"].append(typ)
        cols["tp"].append(tps); cols["kw"].append(kws)
    n = len(cols["n"])
    log(f"Liste nominative : {n} docteurs ({P.explorer}), thèmes connus pour {with_meta}")
    return {"horizons": hs, "lag": P.lag, "typ": TYP, "vtp": list(vt), "vkw": list(vk), "meta_cov": round(with_meta / max(1, n), 3), **cols}


def explorer_ids(A, P):
    """Thèses dont il faut les métadonnées pour l'onglet Explorer (docteurs avec au moins une affiliation après la thèse)."""
    ids = []
    for t, (r, p) in zip(A.theses, A.res):
        if p is None:
            continue
        T = t["year"]
        ys = [y for y in p.by_year if T + P.lag + 1 <= y <= min(T + 10, P.last_obs)]
        if not ys:
            continue
        if P.explorer == "mobiles" and not any(year_loc(p, y, P)[0] == "E" for y in ys):
            continue
        ids.append(t["work"])
    return ids


# ====================================================================== partis / restés
def groups_pf(A, P):
    k = P.pf_h
    out = []
    for t, d, o, (st, ex, c, ret, _) in zip(A.theses, A.disc_of, A.orc, A.cls[k]):
        if st in (0, 1):
            out.append((t["work"], d, "resté", o))
        elif st in (2, 3) and not ret:
            out.append((t["work"], d, "parti", o))
    return pd.DataFrame(out, columns=["work", "discipline", "groupe", "orcid"])


def profils(G, meta, P):
    G = G[G["work"].isin(meta)].copy()
    if G.empty:
        return None, None, None
    G["kw"] = G["work"].map(lambda w: set(k for k, s in zip(meta[w]["kw"], meta[w]["kw_score"]) if s >= P.min_score))
    G["tp"] = G["work"].map(lambda w: set(meta[w]["topics"]))
    G["en"] = G["work"].map(lambda w: meta[w]["lang"] == "en")
    G["etab"] = G["work"].map(lambda w: INST_MERGE.get((meta[w]["inst"] or ["?"])[0], (meta[w]["inst"] or ["?"])[0]))
    PF = {"horizon": P.pf_h, "disciplines": {}, "noise": {}}
    tabs_kw, tabs_tp, syn = [], [], []

    def top(f, n, asc):
        f = f[(f["z"] < -2) if asc else (f["z"] > 2)].sort_values("z", ascending=asc).head(n)
        return [[t, round(z, 2), round(a, 1), round(b, 1), int(x), int(y)] for t, z, a, b, x, y in
                f[["terme", "z", "part_partis_%", "part_restes_%", "n_partis", "n_restes"]].itertuples(index=False)]
    for disc in sorted(G["discipline"].unique()):
        d = G[G["discipline"] == disc]
        Pt, R = d[d["groupe"] == "parti"], d[d["groupe"] == "resté"]
        if len(Pt) < P.min_group or len(R) < P.min_group:
            continue
        s = {"discipline": disc, "restés": len(R), "partis": len(Pt), "taux_depart_%": round(100 * len(Pt) / len(d), 1),
             "these_en_anglais_restés_%": round(100 * R["en"].mean(), 1), "these_en_anglais_partis_%": round(100 * Pt["en"].mean(), 1)}
        fw = fightin_words(list(Pt["kw"]), list(R["kw"]), P.min_docs)
        ft = fightin_words(list(Pt["tp"]), list(R["tp"]), P.min_docs)
        PF["disciplines"][disc] = {**s, "kw_partis": top(fw, 25, False), "kw_restes": top(fw, 25, True),
                                   "tp_partis": top(ft, 25, False), "tp_restes": top(ft, 25, True)}
        PF["noise"][disc] = {key: [int(len(f)), round(100 * (f["z"].abs() >= 2).mean(), 1) if len(f) else 0,
                                   round(100 * (f["z"].abs() >= 3).mean(), 2) if len(f) else 0] for key, f in (("kw", fw), ("tp", ft))}
        syn.append(s)
        fw.insert(0, "discipline", disc); ft.insert(0, "discipline", disc)
        tabs_kw.append(fw); tabs_tp.append(ft)
    # établissements : départs observés / attendus à discipline égale
    base = G.groupby("discipline")["groupe"].apply(lambda x: (x == "parti").mean()).to_dict()
    tot = G.groupby("discipline").size().to_dict()
    known = G[G["etab"] != "?"]
    cells = known.groupby(["etab", "discipline"]).agg(n=("groupe", "size"), partis=("groupe", lambda x: int((x == "parti").sum()))).reset_index()
    cells = cells[cells["n"] >= P.min_etab]
    cells["exp"] = cells["n"] * cells["discipline"].map(base)
    inst = cells.groupby("etab").agg(n=("n", "sum"), o=("partis", "sum"), e=("exp", "sum"), nd=("discipline", "nunique")).reset_index()
    inst["oe"] = inst["o"] / inst["e"]
    inst = inst.sort_values("oe", ascending=False)
    unk = G[G["etab"] == "?"]
    unk_oe = (unk["groupe"] == "parti").sum() / max(1e-9, sum(base[d] for d in unk["discipline"])) if len(unk) else None
    ET = {"inst": [[r.etab, int(r.n), int(r.o), round(r.e, 1), round(r.oe, 3), *[round(v, 3) for v in poisson_ratio_ci(r.o, r.e)], int(r.nd)]
                   for r in inst.itertuples()],
          "cells": [[r.etab, r.discipline, int(r.n), int(r.partis), round(100 * r.partis / r.n, 1), *[round(v, 1) for v in wilson(r.partis, r.n)]]
                    for r in cells.itertuples()],
          "base": {d: round(100 * v, 1) for d, v in base.items()}, "tot": {d: int(v) for d, v in tot.items()},
          "coverage": round(cells["n"].sum() / max(1, len(G)), 3), "unknown_oe": round(unk_oe, 2) if unk_oe else None,
          "min_etab": P.min_etab}
    tables = {"synthese": pd.DataFrame(syn),
              "mots_cles": pd.concat(tabs_kw).round(2) if tabs_kw else pd.DataFrame(),
              "topics": pd.concat(tabs_tp).round(2) if tabs_tp else pd.DataFrame(),
              "etablissements": inst.round(3), "etab_x_discipline": cells.round(2)}
    return PF, ET, tables


# ====================================================================== messages (insights)
class Agg:
    """Agrégations sur le cube du tableau de bord (mêmes règles que la page)."""

    def __init__(self, D):
        self.D = D

    def st(self, h, dis=None, ys=None, o=(0, 1)):
        c = self.D["cube"][str(h)]; v = [0] * 6
        dis = range(len(self.D["disc"])) if dis is None else dis
        ys = range(len(self.D["years"])) if ys is None else [self.D["years"].index(y) for y in ys if y in self.D["years"]]
        for d in dis:
            for y in ys:
                for oo in o:
                    for s in range(6):
                        v[s] += c[d][y][oo][s]
        n = sum(v); vis = sum(v[:4])
        return {"n": n, "vis": vis, "visP": 100 * vis / n if n else float("nan"), "frV": 100 * (v[0] + v[1]) / vis if vis else float("nan"),
                "abV": 100 * (v[2] + v[3]) / vis if vis else float("nan"), "privV": 100 * (v[1] + v[3]) / vis if vis else float("nan"),
                "none": 100 * v[5] / n if n else 0, "int": 100 * v[4] / n if n else 0, "v": v}

    def ctry(self, h, ys=None):
        by = collections.Counter(); ret = collections.Counter()
        yset = None if ys is None else {self.D["years"].index(y) for y in ys if y in self.D["years"]}
        for c, d, y, o, n, r in self.D["ctry"][str(h)]:
            if yset is not None and y not in yset:
                continue
            by[c] += n; ret[c] += r
        return by, ret


def retenir(D, AR, P):
    """Messages globaux (toutes disciplines, horizon principal), chacun avec son ratio n / N."""
    G = Agg(D); H = P.main_h; hs = sorted(P.horizons)
    s, smin, smax = G.st(H), G.st(hs[0]), G.st(hs[-1])
    cards = []
    e_ = html.escape

    def c(label, fig, ratio, base, body):
        cards.append(f'          <div class="dd-kpi"><p class="dd-kpi__label">{e_(label)}</p><p class="dd-kpi__value">{e_(fig)}</p>'
                     f'<p class="dd-kpi__ratio">{e_(ratio)}</p><p class="dd-kpi__base">{e_(base)}</p><p class="dd-kpi__text">{e_(body)}</p></div>')
    v = s["v"]; vis = s["vis"]; n = s["n"]; fr_ = v[0] + v[1]; ab = v[2] + v[3]; none_ = v[4] + v[5]
    if n:
        c(f"Encore visibles à {H} ans", pc(100 * vis / n), f"{fr(vis)} / {fr(n)}", "docteurs avec une affiliation récente / docteurs analysables",
          f"Les autres, {pc(100 * none_ / n)} ({fr(none_)} / {fr(n)}), n'ont plus d'affiliation récente dans leurs publications.")
    if vis:
        b1 = smin["v"][2] + smin["v"][3]; b2 = smax["v"][2] + smax["v"][3]
        c("À l'étranger", pc(100 * ab / vis), f"{fr(ab)} / {fr(vis)}", "docteurs affiliés à l'étranger / docteurs visibles",
          f"{pc(100 * fr_ / vis)} sont en France ({fr(fr_)} / {fr(vis)}). Part de l'étranger : {pc(100 * b1 / max(1, smin['vis']))} à {hs[0]} ans "
          f"({fr(b1)} / {fr(smin['vis'])}), {pc(100 * b2 / max(1, smax['vis']))} à {hs[-1]} ans ({fr(b2)} / {fr(smax['vis'])}).")
    per = [(d, G.st(H, [i])) for i, d in enumerate(D["disc"]) if d != NC]
    per = sorted([(d, x) for d, x in per if x["vis"] >= 100], key=lambda t: -t[1]["abV"])
    if len(per) >= 2:
        (dh, xh), (dl, xl) = per[0], per[-1]
        ah, al = xh["v"][2] + xh["v"][3], xl["v"][2] + xl["v"][3]
        c("Écart entre disciplines", f"{pc(xh['abV'])} / {pc(xl['abV'])}", f"{fr(ah)} / {fr(xh['vis'])} · {fr(al)} / {fr(xl['vis'])}",
          f"à l'étranger / visibles, en {dh.lower()} puis en {dl.lower()}", "Les deux extrêmes. Le filtre Discipline donne le détail des autres.")
    cby, cret = G.ctry(H); N = sum(cby.values())
    if N:
        (tc, tn), = cby.most_common(1)
        eu = sum(x for k, x in cby.items() if k in EUROPE); r = sum(cret.values())
        c(f"Premier pays : {cn(tc)}", pc(100 * tn / N), f"{fr(tn)} / {fr(N)}", "docteurs dans ce pays / docteurs à l'étranger",
          f"{pc(100 * eu / N)} sont en Europe ({fr(eu)} / {fr(N)}) ; {pc(100 * r / N)} ({fr(r)} / {fr(N)}) sont rentrés dans leur pays d'origine.")
    if AR.get("early"):
        e = [sum(o[j] for d in AR["early"] for y in d for o in y) for j in range(3)]; ne = sum(e)
        if ne:
            c(f"Partis tôt, revenus à {P.ar_h} ans", pc(100 * e[0] / ne), f"{fr(e[0])} / {fr(ne)}", "revenus en France / partis tôt à l'étranger",
              f"{pc(100 * e[1] / ne)} sont toujours à l'étranger ({fr(e[1])} / {fr(ne)}). Cohortes {AR['years'][0]}–{AR['years'][-1]}, hors retours au pays d'origine.")
    return "\n".join(cards)


# ====================================================================== page HTML
def render_html(D, AR, PF, ET, P, ntheses, out_file, EX=None):
    G = Agg(D); H = P.main_h
    ex = D["excl"][str(H)]; nobs = sum(G.st(H)["v"]) + sum(v for k, v in ex.items() if k != "censuré")
    tokens = {
        "FILTER": html.escape(P.filter), "DATE": datetime.date.today().strftime("%d/%m/%Y"),
        "HERO": f"{fr(ntheses)} thèses soutenues de {D['years'][0]} à {D['years'][-1]}. Le devenir de chaque docteur est déduit des affiliations "
                "déclarées dans ses publications : en France, à l'étranger, ou plus de publication affiliée. Les filtres ci-dessous s'appliquent aux onglets ; "
                "chacun indique ceux qu'il utilise.",
        "MAINH": str(H), "Y0": str(D["years"][0]), "Y1": str(D["years"][-1]), "YMAIN": str(P.last_obs - H), "NTHESES": fr(ntheses),
        "RETENIR": retenir(D, AR, P),
        "BLOC": ('<div class="fr-header__brand-top"><div class="fr-header__logo"><p class="fr-logo">'
                 + "<br>".join(html.escape(x.strip()) for x in P.bloc_marque.split("|")) + '</p></div></div>') if P.bloc_marque else "", "RETWIN": str(P.ret_window), "LAG": str(P.lag), "LAG1": str(P.lag + 1), "EARLY": str(P.early),
        "ARH": str(P.ar_h), "ARY0": str(AR["years"][0]) if AR["years"] else "–", "ARY1": str(AR["years"][-1]) if AR["years"] else "–",
        "PFH": str(P.pf_h), "LASTOBS": str(P.last_obs), "MINETAB": str(P.min_etab), "ETCOV": pc(100 * ET["coverage"]) if ET else "–",
        "CENS_TXT": " ; ".join(f"cohortes jusqu'en {P.last_obs - k} à T+{k}" for k in sorted(P.horizons)),
        "EXCL_TXT": (f"À T+{H}, parmi les cohortes observables, {pc(100 * ex.get('sans profil', 0) / max(1, nobs), 1)} des thèses "
                     f"({fr(ex.get('sans profil', 0))} / {fr(nobs)}) n'ont pas de profil auteur et {pc(100 * ex.get('profil douteux', 0) / max(1, nobs), 1)} "
                     f"({fr(ex.get('profil douteux', 0))} / {fr(nobs)}) sont écartées comme douteuses : plusieurs thèses françaises sur un même profil, "
                     f"ou activité commencée plus de {P.max_pre} ans avant la thèse."),
        "PARAMS": html.escape(f"années {P.y1}–{P.y2}" + (f", échantillon de {P.sample} thèses/an" if P.sample else "") + f", horizons {', '.join(map(str, P.horizons))}, "
                              f"années ignorées {P.lag}, fenêtre de présence {P.tol} ans, règle multi-pays « {P.multi} », dernière année fiable {P.last_obs}"
                              + (", docteurs avec ORCID seulement" if P.orcid_only else "") + (", sans distinction du privé" if P.no_private else "")),
    }
    page = TEMPLATE
    for k, v in tokens.items():
        page = page.replace(f"@@{k}@@", v)
    left = re.findall(r"@@[A-Z0-9_]+@@", page)
    if left:
        log(f"  avertissement : jetons non remplacés {set(left)}")

    def js(o):
        return json.dumps(o, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    page = page.replace("/*DATA*/null", js(D)).replace("/*PF*/null", js(PF) if PF else "null") \
               .replace("/*ET*/null", js(ET) if ET else "null").replace("/*AR*/null", js(AR) if AR.get("cube") else "null") \
               .replace("/*EX*/null", js(EX) if EX else "null")
    head, body = page.split("<!--BODY-->", 1)
    doc = ('<!doctype html>\n<html lang="fr" data-fr-scheme="system">\n<head>\n<meta charset="utf-8">\n'
           '<meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">\n'
           '<meta name="format-detection" content="telephone=no">\n' + head + '\n</head>\n<body>\n' + body + '\n</body>\n</html>\n')
    out_file.write_text(doc, encoding="utf-8")


# ====================================================================== exports Excel
def write_synthese(A, frames, D, AR_df, P, out):
    xl = {}
    for k, df in frames.items():
        ok = df[df["statut"].notna()]
        t = pd.crosstab(ok["discipline"], ok["statut"])
        pct = (t.div(t.sum(axis=1), axis=0) * 100).round(1).add_suffix(" (%)")
        xl[f"T{k}_discipline"] = pd.concat([t.sum(axis=1).rename("analysables"), pct, t.add_suffix(" (n)")], axis=1)
        t = pd.crosstab(ok["annee"], ok["statut"])
        xl[f"T{k}_cohorte"] = pd.concat([t.sum(axis=1).rename("analysables"), (t.div(t.sum(axis=1), axis=0) * 100).round(1)], axis=1)
        ab = ok[ok["pays"].notna() & (ok["pays"] != "FR")]
        if len(ab):
            pays = ab.groupby("pays").agg(docteurs=("work", "size"), retours=("retour", "sum")).sort_values("docteurs", ascending=False)
            pays["part_%"] = (100 * pays["docteurs"] / len(ab)).round(1)
            xl[f"T{k}_pays"] = pays
        xl[f"T{k}_exclusions"] = df["exclusion"].value_counts().rename("n").to_frame()
    rows = []
    for d, v in D["traj"].items():
        if d == "_j":
            continue
        for i, j in enumerate(D["traj"]["_j"]):
            rows.append({"discipline": d, "annees_apres_these": j, "observes": v["observes"][i], "visibles_%": v["visibles_%"][i],
                         "en_france_%": v["en_france_%"][i], "etranger_%": v["etranger_%"][i]})
    xl["trajectoire"] = pd.DataFrame(rows)
    if len(AR_df):
        t = pd.crosstab(AR_df["discipline"], AR_df["trajectoire"])
        xl["allers_retours"] = pd.concat([t, (t.div(t.sum(axis=1), axis=0) * 100).round(1).add_suffix(" (%)")], axis=1)
    xl["parametres"] = pd.DataFrame({"valeur": {k: str(v) for k, v in vars(P).items() if k != "key"}})
    with pd.ExcelWriter(out / "synthese.xlsx") as w:
        for name, d in xl.items():
            d.to_excel(w, sheet_name=name[:31])
    if len(AR_df):
        AR_df.to_csv(out / "trajectoires_annuelles.csv", index=False, sep=";", encoding="utf-8-sig")


# ====================================================================== CLI
def parse(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    g = ap.add_argument_group("collecte")
    g.add_argument("--key", default=os.environ.get("OPENALEX_API_KEY"), help="clé API (ou variable OPENALEX_API_KEY)")
    g.add_argument("--mailto", default=os.environ.get("OPENALEX_MAILTO"))
    g.add_argument("--from", dest="y1", type=int, default=2006)
    g.add_argument("--to", dest="y2", type=int, default=datetime.date.today().year - 1)
    g.add_argument("--sample", type=int, default=0, help="N thèses aléatoires par année (0 = toutes)")
    g.add_argument("--seed", type=int, default=42)
    g.add_argument("--filter", default=BASE_FILTER)
    g.add_argument("--concurrency", type=int, default=10)
    g.add_argument("--cache", default="cache_openalex")
    g.add_argument("--offline", action="store_true", help="aucune requête : uniquement le cache")
    g.add_argument("--refresh", action="store_true", help="vide et re-télécharge tout le cache (thèses, profils, métadonnées)")
    g.add_argument("--max-age", type=int, default=180, help="âge du cache (jours) au-delà duquel il est mis à jour automatiquement")
    g.add_argument("--no-meta", action="store_true", help="sans la collecte des métadonnées des thèses (pas de section « qui part »)")
    a = ap.add_argument_group("analyse")
    a.add_argument("--horizon", dest="horizons", type=int, nargs="+", default=[3, 5, 8])
    a.add_argument("--main-horizon", dest="main_h", type=int, default=5, help="horizon des messages et des comparaisons")
    a.add_argument("--lag", type=int, default=1, help="années T..T+lag ignorées (labo de thèse)")
    a.add_argument("--tol", type=int, default=2, help="fenêtre de présence en fin de période")
    a.add_argument("--last-obs", type=int, default=datetime.date.today().year - 1, help="dernière année fiable")
    a.add_argument("--multi", choices=["fr", "major", "abroad"], default="fr")
    a.add_argument("--no-private", action="store_true")
    a.add_argument("--orcid-only", action="store_true")
    a.add_argument("--keep-suspect", action="store_true")
    a.add_argument("--max-pre", type=int, default=12)
    a.add_argument("--ret-window", type=int, default=3)
    a.add_argument("--early", type=int, default=5, help="« parti tôt » : à l'étranger au plus tard à T+early")
    a.add_argument("--min-docs", type=int, default=10)
    a.add_argument("--min-score", type=float, default=0.0)
    a.add_argument("--min-group", type=int, default=50)
    a.add_argument("--min-etab", type=int, default=60)
    a.add_argument("--out", default="resultats")
    a.add_argument("--explorer", choices=["visibles", "mobiles", "tous", "non"], default="visibles",
                   help="liste nominative de l'onglet Explorer : docteurs visibles après la thèse (défaut), seulement ceux passés par l'étranger, tous, ou aucune")
    a.add_argument("--bloc-marque", default="", help="bloc-marque DSFR de l'en-tête, lignes séparées par | (ex. « Ministère|de l'Enseignement supérieur »)")
    P = ap.parse_args(argv)
    if P.y1 > P.y2:
        P.y1, P.y2 = P.y2, P.y1
    P.horizons = sorted(set(P.horizons))
    if P.main_h not in P.horizons:
        P.horizons = sorted(set(P.horizons) | {P.main_h})
    P.lag = min(P.lag, min(P.horizons) - 1)
    P.ar_h = max(P.horizons)
    P.pf_h = P.main_h
    P.early = min(P.early, P.ar_h - 1)
    return P


def main(argv=None, transport=None):
    P = parse(argv)
    root = Path(P.cache); root.mkdir(parents=True, exist_ok=True)
    cache = root / P.filter.replace(":", "=").replace(",", "+").replace("/", "_")
    cache.mkdir(parents=True, exist_ok=True)
    out = Path(P.out); out.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    cache_freshness(P, cache, root)
    try:
        if P.offline:
            theses, authors, alias = load_cache(P, cache)
            log(f"Cache : {len(theses)} thèses, {len(authors)} profils")
        else:
            if not P.key:
                log("Attention : pas de clé API, quota anonyme très bas.")
            theses, authors, alias = asyncio.run(collect(P, cache, transport))
            mark_fresh(cache)
    except ApiError as e:
        log(f"ERREUR : {e}")
        if e.status == 400:
            log("Vérifie --filter (le filtre est-il accepté par OpenAlex ?).")
        log("Ce qui a été collecté est en cache : relance la même commande pour reprendre.")
        sys.exit(1)
    if not theses:
        sys.exit("Aucune thèse (cache vide ? lance sans --offline).")

    log("Calcul des statuts…")
    A = analyse_core(theses, authors, alias, P)
    frames = per_thesis_frames(A, P, out)
    D = dashboard_data(A, P)
    log("Allers-retours…")
    AR, AR_df = aller_retour(A, P)
    write_synthese(A, frames, D, AR_df, P, out)

    PF = ET = EX = None
    meta = {}
    mf = root / "theses_meta.jsonl"
    if not P.no_meta:
        G = groups_pf(A, P)
        log(f"Partis / restés à T+{P.pf_h} : {int((G.groupe == 'parti').sum())} partis, {int((G.groupe == 'resté').sum())} restés")
        want = list(G["work"])
        if P.explorer != "non":
            want = list(dict.fromkeys(want + explorer_ids(A, P)))
        try:
            if P.offline:
                meta = {m["work"]: m for m in read_jsonl(mf)}
            else:
                async def run():
                    oa = OpenAlex(P.key, P.mailto, P.concurrency, transport)
                    try:
                        return await fetch_meta(oa, want, mf)
                    finally:
                        await oa.close()
                meta = asyncio.run(run())
            PF, ET, tables = profils(G, meta, P)
            if tables:
                with pd.ExcelWriter(out / "partants_restes.xlsx") as w:
                    for name, d in tables.items():
                        if len(d):
                            d.to_excel(w, sheet_name=name[:31], index=False)
        except ApiError as e:
            log(f"Métadonnées incomplètes ({e}) : relance pour compléter.")
            meta = {m["work"]: m for m in read_jsonl(mf)}
    if P.explorer != "non":
        EX = explorer_data(A, P, meta)

    log("Messages et tableau de bord…")
    render_html(D, AR, PF, ET, P, len(A.theses), out / "tableau_de_bord.html", EX)
    log(f"Terminé en {time.time() - t0:.0f} s. Ouvre {out / 'tableau_de_bord.html'} dans ton navigateur.")


# ====================================================================== gabarit HTML du tableau de bord
TEMPLATE = r'''<title>Devenir des docteurs formés en France</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@gouvfr/dsfr@1.15.3/dist/dsfr.min.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@gouvfr/dsfr@1.15.3/dist/utility/utility.min.css">
<style>
/* Compléments au DSFR : uniquement les graphiques et les vignettes de chiffres, avec les jetons de couleur du DSFR. */
:root{
  --c-fr: var(--blue-france-main-525, #6a6af4);
  --c-frp: var(--orange-terre-battue-main-645, #e4794a);
  --c-ab: var(--green-menthe-main-548, #009081);
  --c-abp: var(--yellow-tournesol-main-731, #c8aa39);
  --c-int: var(--grey-625-425, #929292);
  --c-none: var(--grey-925-125, #e5e5e5);
  --c-ret: var(--grey-625-425, #929292);
}
.dd-hero h1{margin-bottom:.5rem}
.dd-filters{display:flex;flex-wrap:wrap;gap:1rem 1.5rem;align-items:flex-end}
.dd-filters .fr-select-group{margin-bottom:0;min-width:12rem}
.dd-filters .fr-segmented{margin-bottom:0}
.dd-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(12.5rem,1fr));gap:1rem;margin-bottom:1.5rem}
.dd-kpi{border:1px solid var(--border-default-grey);background:var(--background-default-grey);padding:1.25rem 1.25rem 1rem;display:flex;flex-direction:column;min-width:0}
.dd-kpi__label{font-size:.75rem;line-height:1.25rem;font-weight:700;text-transform:uppercase;letter-spacing:.03em;color:var(--text-default-grey);margin:0 0 .5rem}
.dd-kpi__value{font-size:1.75rem;line-height:2.25rem;font-weight:700;color:var(--text-title-grey);margin:0;font-variant-numeric:tabular-nums}
.dd-kpi__ratio{font-size:.875rem;line-height:1.5rem;color:var(--text-default-grey);margin:0;font-variant-numeric:tabular-nums}
.dd-kpi__base{font-size:.75rem;line-height:1.25rem;color:var(--text-mention-grey);margin:.25rem 0 0}
.dd-kpi__dot{display:inline-block;width:.625rem;height:.625rem;margin-right:.5rem;vertical-align:baseline}
.dd-kpi p.dd-kpi__text{font-size:.875rem;line-height:1.5rem;color:var(--text-default-grey);margin:.5rem 0 0}
.dd-block{border:1px solid var(--border-default-grey);background:var(--background-default-grey);padding:1.5rem;margin-bottom:1.5rem;min-width:0}
.dd-block h3{font-size:1.125rem;line-height:1.5rem;margin-bottom:.5rem}
.dd-note{font-size:.875rem;line-height:1.5rem;color:var(--text-mention-grey);margin-bottom:1rem;max-width:52rem}
.dd-legend{display:flex;flex-wrap:wrap;gap:.25rem 1.25rem;font-size:.75rem;line-height:1.25rem;color:var(--text-default-grey);margin-bottom:.75rem}
.dd-legend span{display:inline-flex;align-items:center;gap:.4rem}
.dd-sw{display:inline-block;width:.75rem;height:.75rem}
.dd-two{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:1.5rem}
@media (max-width:62em){.dd-two{grid-template-columns:minmax(0,1fr)}}
.dd-chart{width:100%;min-width:0}
.dd-chart svg{display:block;width:100%;height:auto;font-family:Marianne,arial,sans-serif}
.dd-chart text{fill:var(--text-default-grey);font-size:11px}
.dd-chart text.m{fill:var(--text-mention-grey)}
.dd-chart text.v{fill:var(--text-title-grey);font-weight:700}
.dd-chart text.sel{fill:var(--text-title-blue-france);font-weight:700}
.dd-chart .grid{stroke:var(--border-default-grey);stroke-width:1}
.dd-chart .axis{stroke:var(--border-plain-grey);stroke-width:1}
.dd-chart .hit{fill:transparent}
.dd-empty{color:var(--text-mention-grey);font-size:.875rem;padding:1.5rem 0;text-align:center}
.dd-applied{display:flex;flex-wrap:wrap;gap:.5rem;align-items:center;margin-bottom:1.5rem}
.dd-applied .fr-tag{margin:0}
.dd-applied .dd-off{background:transparent;box-shadow:inset 0 0 0 1px var(--border-default-grey);color:var(--text-mention-grey)}
.dd-applied__label{font-size:.75rem;color:var(--text-mention-grey);margin:0 .25rem 0 0}
.dd-zl{display:flex;flex-direction:column}
.dd-zr{display:grid;grid-template-columns:minmax(0,1.8fr) minmax(3rem,1fr) 2.25rem 5.5rem;gap:.5rem;align-items:center;padding:.3rem .25rem;font-size:.875rem;border-bottom:1px solid var(--border-default-grey)}
.dd-zr:hover{background:var(--background-alt-grey)}
.dd-zt{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.dd-zt.strong{font-weight:700}
.dd-zb{height:.5rem;background:var(--background-contrast-grey)}
.dd-zb i{display:block;height:100%}
.dd-zv{text-align:right;font-weight:700;font-variant-numeric:tabular-nums}
.dd-zs{text-align:right;color:var(--text-mention-grey);font-size:.75rem;font-variant-numeric:tabular-nums;white-space:nowrap}
@media (max-width:40em){.dd-zr{grid-template-columns:minmax(0,1fr) 3rem 2rem}.dd-zs{display:none}}
.dd-table td .dd-q{color:var(--text-mention-grey);font-size:.75rem;margin-left:.35rem}
.dd-table tbody tr{cursor:pointer}
.dd-tip{position:fixed;pointer-events:none;z-index:1000;background:var(--background-default-grey);color:var(--text-default-grey);border:1px solid var(--border-default-grey);box-shadow:0 6px 18px rgba(0,0,18,.16);padding:.5rem .75rem;font-size:.75rem;line-height:1.25rem;min-width:12rem;max-width:20rem;opacity:0;transition:opacity .08s}
.dd-tip.on{opacity:1}
.dd-tip h4{margin:0 0 .25rem;font-size:.8125rem;line-height:1.25rem}
.dd-tip .r{display:flex;justify-content:space-between;gap:.75rem}
.dd-tip .r span:first-child{display:flex;align-items:center;gap:.4rem;color:var(--text-mention-grey)}
.dd-tip .dd-sw{width:.5rem;height:.5rem}
.dd-seq{display:inline-flex;gap:2px;vertical-align:middle}
.dd-seq i{display:inline-block;width:.6rem;height:.9rem}
#xTable td{white-space:nowrap}
#xTable td.wrap{white-space:normal;min-width:12rem}
.dd-method p{font-size:.875rem;line-height:1.5rem}
.fr-footer__content{flex:1 1 100%;max-width:100%;margin-left:0}
@media (prefers-reduced-motion:reduce){.dd-tip{transition:none}}
</style>
<!--BODY-->
<header role="banner" class="fr-header">
  <div class="fr-header__body">
    <div class="fr-container">
      <div class="fr-header__body-row">
        <div class="fr-header__brand fr-enlarge-link">
          @@BLOC@@
          <div class="fr-header__service">
            <a href="#main" title="Accueil – Devenir des docteurs"><p class="fr-header__service-title">Devenir des docteurs <span class="fr-badge fr-badge--sm fr-badge--green-menthe">BETA</span></p></a>
            <p class="fr-header__service-tagline">Insertion et mobilité des docteurs formés en France, d'après OpenAlex</p>
          </div>
        </div>
      </div>
    </div>
  </div>
</header>

<main id="main" role="main">
  <div class="fr-background-alt--blue-france dd-hero">
    <div class="fr-container fr-py-5w">
      <h1>Que deviennent les docteurs formés en France ?</h1>
      <p class="fr-text--lg fr-mb-1w">@@HERO@@</p>
      <p class="fr-text--xs fr-mb-4w" style="color:var(--text-mention-grey)">Source : OpenAlex, filtre <code>@@FILTER@@</code> · collecte du @@DATE@@</p>
      <div class="dd-filters" role="group" aria-label="Filtres">
        <fieldset class="fr-segmented fr-segmented--sm" id="segH"><legend class="fr-segmented__legend">Horizon après la thèse</legend><div class="fr-segmented__elements" id="segHel"></div></fieldset>
        <div class="fr-select-group"><label class="fr-label" for="fDisc">Discipline</label><select class="fr-select" id="fDisc"></select></div>
        <div class="fr-select-group" style="min-width:8rem"><label class="fr-label" for="fY1">Cohortes de</label><select class="fr-select" id="fY1"></select></div>
        <div class="fr-select-group" style="min-width:8rem"><label class="fr-label" for="fY2">à</label><select class="fr-select" id="fY2"></select></div>
        <button type="button" class="fr-btn fr-btn--tertiary fr-btn--sm fr-icon-refresh-line fr-btn--icon-left" id="reset">Réinitialiser</button>
      </div>
    </div>
  </div>

  <div class="fr-container fr-mt-4w fr-mb-8w">
    <div class="fr-tabs" id="tabs">
      <ul class="fr-tabs__list" role="tablist" aria-label="Sections du tableau de bord">
        <li role="presentation"><button type="button" id="tab-0" class="fr-tabs__tab" tabindex="0" role="tab" aria-selected="true" aria-controls="panel-0">Aperçu</button></li>
        <li role="presentation"><button type="button" id="tab-1" class="fr-tabs__tab" tabindex="-1" role="tab" aria-selected="false" aria-controls="panel-1">Où sont-ils ?</button></li>
        <li role="presentation"><button type="button" id="tab-2" class="fr-tabs__tab" tabindex="-1" role="tab" aria-selected="false" aria-controls="panel-2">Destinations</button></li>
        <li role="presentation"><button type="button" id="tab-3" class="fr-tabs__tab" tabindex="-1" role="tab" aria-selected="false" aria-controls="panel-3">Évolution</button></li>
        <li role="presentation"><button type="button" id="tab-4" class="fr-tabs__tab" tabindex="-1" role="tab" aria-selected="false" aria-controls="panel-4">Partir et revenir</button></li>
        <li role="presentation"><button type="button" id="tab-5" class="fr-tabs__tab" tabindex="-1" role="tab" aria-selected="false" aria-controls="panel-5">Qui part, qui reste</button></li>
        <li role="presentation"><button type="button" id="tab-6" class="fr-tabs__tab" tabindex="-1" role="tab" aria-selected="false" aria-controls="panel-6">Établissements</button></li>
        <li role="presentation"><button type="button" id="tab-8" class="fr-tabs__tab" tabindex="-1" role="tab" aria-selected="false" aria-controls="panel-8">Explorer</button></li>
        <li role="presentation"><button type="button" id="tab-7" class="fr-tabs__tab" tabindex="-1" role="tab" aria-selected="false" aria-controls="panel-7">Méthode</button></li>
      </ul>

      <div id="panel-0" class="fr-tabs__panel fr-tabs__panel--selected" role="tabpanel" aria-labelledby="tab-0" tabindex="0">
        <h2 class="fr-h4">À retenir</h2>
        <p class="dd-note">Toutes disciplines, @@MAINH@@ ans après la thèse, cohortes @@Y0@@–@@YMAIN@@. Ces chiffres sont fixes : les filtres s'appliquent aux autres onglets.</p>
        <div class="dd-kpis">
@@RETENIR@@
        </div>
        <div class="fr-alert fr-alert--warning fr-alert--sm">
          <p>« Sans trace » ne veut pas dire « abandon de la recherche ». L'outil mesure la présence dans les publications : un docteur en entreprise ou dans l'administration publie peu, et un profil auteur mal relié dans OpenAlex disparaît aussi. La part sans trace est donc un maximum pour les sorties, pas une mesure. OpenAlex ne connaît pas la nationalité des docteurs.</p>
        </div>
      </div>

      <div id="panel-1" class="fr-tabs__panel" role="tabpanel" aria-labelledby="tab-1" tabindex="0">
        <h2 class="fr-h4">Où sont les docteurs ?</h2>
        <p class="dd-note">Statut de chaque docteur à l'horizon choisi, d'après la dernière affiliation déclarée dans ses publications.</p>
        <div class="dd-applied" id="c1"></div>
        <div class="dd-kpis" id="k1"></div>
        <div class="dd-block"><h3>Statut par discipline</h3><p class="dd-note" id="n1"></p><div class="dd-legend" id="lg1"></div><div class="dd-chart" id="ch1"></div></div>
      </div>

      <div id="panel-2" class="fr-tabs__panel" role="tabpanel" aria-labelledby="tab-2" tabindex="0">
        <h2 class="fr-h4">Où partent ceux qui partent ?</h2>
        <p class="dd-note">Pays des docteurs affiliés à l'étranger à l'horizon choisi.</p>
        <div class="dd-applied" id="c2"></div>
        <div class="dd-kpis" id="k2"></div>
        <div class="dd-block"><h3>Principaux pays de destination</h3><p class="dd-note">Assiette : docteurs affiliés à l'étranger. En gris, les retours dans un pays où le docteur était déjà affilié au moins @@RETWIN@@ ans avant la thèse.</p>
          <div class="dd-legend"><span><i class="dd-sw" style="background:var(--c-ab)"></i>Nouvelle destination</span><span><i class="dd-sw" style="background:var(--c-ret)"></i>Retour au pays d'origine</span></div>
          <div class="dd-chart" id="ch2"></div></div>
      </div>

      <div id="panel-3" class="fr-tabs__panel" role="tabpanel" aria-labelledby="tab-3" tabindex="0">
        <h2 class="fr-h4">Évolution d'une cohorte à l'autre</h2>
        <p class="dd-note">Chaque point est une année de soutenance, observée au même horizon après la thèse. Les cohortes trop récentes pour l'horizon choisi n'apparaissent pas.</p>
        <div class="dd-applied" id="c3"></div>
        <div class="dd-two">
          <div class="dd-block"><h3>Encore visibles</h3><p class="dd-note">Assiette : docteurs analysables de la cohorte.</p><div class="dd-chart" id="ch3a"></div></div>
          <div class="dd-block"><h3>À l'étranger</h3><p class="dd-note">Assiette : docteurs visibles de la cohorte.</p><div class="dd-chart" id="ch3b"></div></div>
        </div>
      </div>

      <div id="panel-4" class="fr-tabs__panel" role="tabpanel" aria-labelledby="tab-4" tabindex="0">
        <h2 class="fr-h4">Partir et revenir</h2>
        <p class="dd-note">Docteurs « partis tôt » : affiliés à l'étranger au moins une année entre T+@@LAG1@@ et T+@@EARLY@@, hors retours au pays d'origine. Leur situation @@ARH@@ ans après la thèse, suivie année par année, cohortes @@ARY0@@–@@ARY1@@.</p>
        <div class="dd-applied" id="c4"></div>
        <div class="dd-kpis" id="k4"></div>
        <div class="dd-two">
          <div class="dd-block"><h3>Devenir des partis tôt, par discipline</h3><p class="dd-note">Assiette : docteurs partis tôt de chaque discipline. Trié par part revenue en France.</p><div class="dd-legend" id="lg4"></div><div class="dd-chart" id="ch4a"></div></div>
          <div class="dd-block"><h3>Durée à l'étranger avant le retour</h3><p class="dd-note">Assiette : docteurs revenus en France après une période à l'étranger. De 1 à 3 ans, le séjour ressemble à un postdoc.</p><div class="dd-chart" id="ch4b"></div></div>
        </div>
      </div>

      <div id="panel-5" class="fr-tabs__panel" role="tabpanel" aria-labelledby="tab-5" tabindex="0">
        <h2 class="fr-h4">Qui part, qui reste ?</h2>
        <p class="dd-note">Thèses des docteurs affiliés en France (restés) et à l'étranger hors retours au pays d'origine (partis), <span id="pfh"></span> ans après la thèse : langue de la thèse et thèmes OpenAlex.</p>
        <div class="dd-applied" id="c5"></div>
        <div id="p5all" class="dd-block">
          <h3>Par discipline</h3>
          <p class="dd-note">Clique sur une ligne, ou choisis une discipline dans les filtres, pour voir les thèmes qui distinguent partis et restés. « Signal » : part des thèmes nettement différents entre les deux groupes ; le hasard seul en produirait 4,6 %.</p>
          <div class="fr-table fr-table--sm fr-mb-0 dd-table"><div class="fr-table__wrapper"><div class="fr-table__container"><div class="fr-table__content"><table id="t5"></table></div></div></div></div>
        </div>
        <div id="p5one" hidden>
          <div class="dd-kpis" id="k5"></div>
          <div class="dd-block">
            <div style="display:flex;justify-content:space-between;gap:1rem;flex-wrap:wrap;align-items:flex-start">
              <div><h3>Thèmes qui distinguent partis et restés</h3><p class="dd-note" id="n5"></p></div>
              <fieldset class="fr-segmented fr-segmented--sm fr-segmented--no-legend" id="segK"><legend class="fr-segmented__legend">Type de thème</legend><div class="fr-segmented__elements">
                <div class="fr-segmented__element"><input value="tp" checked type="radio" id="k-tp" name="kind"><label class="fr-label" for="k-tp">Topics</label></div>
                <div class="fr-segmented__element"><input value="kw" type="radio" id="k-kw" name="kind"><label class="fr-label" for="k-kw">Mots-clés</label></div>
              </div></fieldset>
            </div>
            <div class="dd-two">
              <div><p class="fr-text--sm fr-text--bold fr-mb-1w">Plus fréquents chez les partis</p><div class="dd-zl" id="z5p"></div></div>
              <div><p class="fr-text--sm fr-text--bold fr-mb-1w">Plus fréquents chez les restés</p><div class="dd-zl" id="z5r"></div></div>
            </div>
          </div>
        </div>
      </div>

      <div id="panel-6" class="fr-tabs__panel" role="tabpanel" aria-labelledby="tab-6" tabindex="0">
        <h2 class="fr-h4">Établissements de soutenance</h2>
        <p class="dd-note" id="h6"></p>
        <div class="dd-applied" id="c6"></div>
        <div class="dd-block"><h3 id="t6"></h3><p class="dd-note" id="n6"></p>
          <div class="dd-legend"><span><i class="dd-sw" style="background:var(--c-ab)"></i>Nettement au-dessus de la moyenne</span><span><i class="dd-sw" style="background:var(--c-fr)"></i>Nettement en dessous</span><span><i class="dd-sw" style="background:var(--c-int)"></i>Écart non significatif</span><span style="color:var(--text-mention-grey)">trait : intervalle de confiance à 95 %</span></div>
          <div class="dd-chart" id="ch6"></div></div>
      </div>

      <div id="panel-8" class="fr-tabs__panel" role="tabpanel" aria-labelledby="tab-8" tabindex="0">
        <h2 class="fr-h4">Explorer les docteurs</h2>
        <p class="dd-note">Liste des docteurs avec leur parcours année par année après la thèse (jusqu'à 10 ans). Les filtres Discipline et Cohortes du bandeau s'appliquent, ainsi que ceux ci-dessous. Chaque nom renvoie au profil OpenAlex.</p>
        <div class="fr-alert fr-alert--info fr-alert--sm fr-mb-3w"><p>Cette liste est nominative. Les données viennent d'OpenAlex (licence CC0) mais restent des données personnelles : à utiliser pour l'analyse, sans diffusion publique du fichier.</p></div>
        <div class="dd-applied" id="c8"></div>
        <div class="fr-grid-row fr-grid-row--gutters fr-mb-1w">
          <div class="fr-col-12 fr-col-md-4"><div class="fr-input-group"><label class="fr-label" for="xTheme">Thème de la thèse<span class="fr-hint-text">Topics et mots-clés OpenAlex, en anglais. Plusieurs termes : séparer par |</span></label><input class="fr-input" type="search" id="xTheme" placeholder="ex. machine learning|neural network"></div>
            <button type="button" class="fr-btn fr-btn--tertiary-no-outline fr-btn--sm" id="xIA">Intelligence artificielle</button></div>
          <div class="fr-col-12 fr-col-md-3"><div class="fr-select-group"><label class="fr-label" for="xCtry">Pays à l'étranger<span class="fr-hint-text">Au moins une année d'affiliation</span></label><select class="fr-select" id="xCtry"></select></div></div>
          <div class="fr-col-12 fr-col-md-3"><div class="fr-select-group"><label class="fr-label" for="xTyp">Trajectoire<span class="fr-hint-text">Sur les années observées</span></label><select class="fr-select" id="xTyp"></select></div></div>
          <div class="fr-col-12 fr-col-md-2"><div class="fr-input-group"><label class="fr-label" for="xName">Nom<span class="fr-hint-text">Recherche</span></label><input class="fr-input" type="search" id="xName"></div></div>
        </div>
        <div style="display:flex;justify-content:space-between;align-items:center;gap:1rem;flex-wrap:wrap" class="fr-mb-2w">
          <p class="fr-text--lg fr-text--bold fr-mb-0" id="xCount"></p>
          <div style="display:flex;gap:.75rem;flex-wrap:wrap"><button type="button" class="fr-btn fr-btn--secondary fr-btn--sm fr-icon-download-line fr-btn--icon-left" id="xCsv">Exporter la sélection (CSV)</button><button type="button" class="fr-btn fr-btn--tertiary fr-btn--sm" id="xReset">Effacer ces filtres</button></div>
        </div>
        <p class="dd-note" id="xMeta"></p>
        <div class="dd-legend"><span><i class="dd-sw" style="background:var(--c-fr)"></i>Affilié en France</span><span><i class="dd-sw" style="background:var(--c-ab)"></i>Affilié à l'étranger</span><span><i class="dd-sw" style="background:var(--c-abp)"></i>Dans son pays d'origine</span><span><i class="dd-sw" style="background:var(--c-none)"></i>Aucune affiliation cette année</span></div>
        <div class="fr-table fr-table--sm fr-mb-2w"><div class="fr-table__wrapper"><div class="fr-table__container"><div class="fr-table__content"><table id="xTable"></table></div></div></div></div>
        <nav role="navigation" class="fr-pagination" aria-label="Pagination"><ul class="fr-pagination__list" id="xPages"></ul></nav>
      </div>

      <div id="panel-7" class="fr-tabs__panel dd-method" role="tabpanel" aria-labelledby="tab-7" tabindex="0">
        <h2 class="fr-h4">Méthode, définitions et limites</h2>
        <div class="fr-grid-row fr-grid-row--gutters">
          <div class="fr-col-12 fr-col-md-6">
            <h3 class="fr-h6">Corpus</h3>
            <p>Thèses OpenAlex sélectionnées par <code>@@FILTER@@</code> : @@NTHESES@@ thèses soutenues de @@Y0@@ à @@Y1@@. Le docteur est le premier auteur de la thèse, identifié par son identifiant auteur OpenAlex. La discipline est le domaine principal de la thèse, regroupé en 11 disciplines.</p>
            <p>@@EXCL_TXT@@</p>
            <h3 class="fr-h6">Statut à T+k</h3>
            <p>On ignore les affiliations de T à T+@@LAG@@, qui viennent des articles tirés de la thèse, puis on retient l'affiliation la plus récente jusqu'à T+k. Si plusieurs pays apparaissent la même année, la France l'emporte ; l'outre-mer compte comme la France. Sans affiliation après T+@@LAG@@ : « pas de trace » ; dernière affiliation trop ancienne : « trace interrompue ». Une cohorte n'est analysée que si T+k ≤ @@LASTOBS@@ : @@CENS_TXT@@.</p>
            <h3 class="fr-h6">Assiettes des pourcentages</h3>
            <p><b>Analysables</b> : docteurs non exclus d'une cohorte observable. <b>Visibles</b> : analysables ayant une affiliation récente. Les parts « en France » et « à l'étranger » sont calculées sur les visibles ; « encore visibles » et « sans trace », sur les analysables. Chaque pourcentage est accompagné de son ratio.</p>
          </div>
          <div class="fr-col-12 fr-col-md-6">
            <h3 class="fr-h6">Partir et revenir</h3>
            <p>Pays d'affiliation suivi année par année de T+@@LAG1@@ à T+@@ARH@@. Un retour au pays d'origine est un départ vers un pays où le docteur était affilié au moins @@RETWIN@@ ans avant la thèse.</p>
            <h3 class="fr-h6">Qui part, qui reste</h3>
            <p>Écart de fréquence d'un thème entre partis et restés, mesuré par un score z (méthode « Fightin' Words ») ; au-delà de 3, l'écart est fort.</p>
            <h3 class="fr-h6">Établissements</h3>
            <p>Départs observés rapportés aux départs attendus d'après le taux moyen de chaque discipline. Seuls les couples établissement × discipline d'au moins @@MINETAB@@ docteurs sont retenus, soit @@ETCOV@@ des docteurs.</p>
            <h3 class="fr-h6">Limites</h3>
            <p>L'outil mesure la présence dans les publications, pas l'emploi. OpenAlex ne connaît pas la nationalité : une partie des partis sont des docteurs étrangers rentrés chez eux ou partis ailleurs. Une affiliation est datée par la publication, environ un an après la prise de poste.</p>
            <p class="fr-text--xs" style="color:var(--text-mention-grey)">Paramètres : @@PARAMS@@.</p>
          </div>
        </div>
      </div>
    </div>
  </div>
</main>

<footer class="fr-footer" role="contentinfo" id="footer">
  <div class="fr-container">
    <div class="fr-footer__body">
      <div class="fr-footer__content">
        <p class="fr-footer__content-desc">Données ouvertes OpenAlex (licence CC0). Collecte du @@DATE@@.</p>
      </div>
    </div>
  </div>
</footer>
<div class="dd-tip" id="tip" role="tooltip"></div>

<script type="module" src="https://cdn.jsdelivr.net/npm/@gouvfr/dsfr@1.15.3/dist/dsfr.module.min.js"></script>
<script nomodule src="https://cdn.jsdelivr.net/npm/@gouvfr/dsfr@1.15.3/dist/dsfr.nomodule.min.js"></script>
<script>
"use strict";
const D = /*DATA*/null;
const PF = /*PF*/null;
const ET = /*ET*/null;
const AR = /*AR*/null;
const $ = id => document.getElementById(id);
const fmt = new Intl.NumberFormat("fr-FR");
const pct = (n,N,d=0) => N ? (100*n/N).toFixed(d).replace(".",",")+" %" : "–";
const rat = (n,N) => `${fmt.format(n)} / ${fmt.format(N)}`;
const esc = s => String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const LO=D.lastObs, Y0=D.years[0];
const ST_LAB=["Académique · France","Entreprise · France","Académique · étranger","Entreprise · étranger","Trace interrompue","Pas de trace"];
const ST_COL=["var(--c-fr)","var(--c-frp)","var(--c-ab)","var(--c-abp)","var(--c-int)","var(--c-none)"];
const DSH={"Sciences sociales & psychologie":"Sciences sociales & psycho."}; const dn=d=>DSH[d]||d;
let RN; try{ RN=new Intl.DisplayNames(["fr"],{type:"region"}); }catch(e){}
const cname=c=>{ try{ return RN?RN.of(c):c; }catch(e){ return c; } };
const EUROPE=new Set("GB DE CH IT BE SE NL ES LU NO AT PL DK CZ PT IE FI GR HU RO SK SI HR BG EE LV LT IS MT CY RS UA RU BA MK AL MD BY ME XK LI MC AD".split(" "));
const S={h:String(D.mainH), disc:"all", y1:Y0, y2:LO, kw:"tp"};
const ALLO=[0,1];

/* ---------- infobulles ---------- */
const tip=$("tip");
function bindTips(root){ root.querySelectorAll("[data-tip]").forEach(el=>{
  el.addEventListener("mouseenter",e=>{ tip.innerHTML=el.dataset.tip; tip.classList.add("on"); mv(e); });
  el.addEventListener("mousemove",mv); el.addEventListener("mouseleave",()=>tip.classList.remove("on")); }); }
function mv(e){ const r=tip.getBoundingClientRect(); let x=e.clientX+14,y=e.clientY+14;
  if(x+r.width>innerWidth-8) x=e.clientX-r.width-14; if(y+r.height>innerHeight-8) y=e.clientY-r.height-14; tip.style.left=x+"px"; tip.style.top=y+"px"; }
const row=(l,v,c)=>`<div class="r"><span>${c?`<i class="dd-sw" style="background:${c}"></i>`:""}${l}</span><span>${v}</span></div>`;

/* ---------- filtres ---------- */
const dList=()=>S.disc==="all"?D.disc.map((_,i)=>i):[D.disc.indexOf(S.disc)];
const discLab=()=>S.disc==="all"?"Toutes disciplines":S.disc;
function applied(id, o){
  const t=[];
  t.push(o.h===true?`<p class="fr-tag fr-tag--sm">Horizon : ${S.h} ans</p>`:`<p class="fr-tag fr-tag--sm dd-off" title="Ce filtre ne s'applique pas ici">Horizon fixe : ${o.h} ans</p>`);
  t.push(o.disc?`<p class="fr-tag fr-tag--sm">${esc(discLab())}</p>`:`<p class="fr-tag fr-tag--sm dd-off">Toutes disciplines (filtre non applicable)</p>`);
  if(o.coh===true) t.push(`<p class="fr-tag fr-tag--sm">Cohortes ${S.y1}–${Math.min(S.y2, LO-+S.h)}</p>`);
  else if(Array.isArray(o.coh)) t.push(`<p class="fr-tag fr-tag--sm">Cohortes ${o.coh[0]}–${o.coh[1]}</p>`);
  else t.push(`<p class="fr-tag fr-tag--sm dd-off">Toutes cohortes (filtre non applicable)</p>`);
  $(id).innerHTML=`<span class="dd-applied__label">Filtres appliqués :</span>`+t.join("");
}
function kpis(id, tiles){
  $(id).innerHTML=tiles.map(t=>`<div class="dd-kpi"><p class="dd-kpi__label">${t.c?`<i class="dd-kpi__dot" style="background:${t.c}"></i>`:""}${t.l}</p><p class="dd-kpi__value">${t.v}</p>${t.r?`<p class="dd-kpi__ratio">${t.r}</p>`:""}<p class="dd-kpi__base">${t.s||""}</p></div>`).join("");
}

/* ---------- agrégations ---------- */
function yIdx(h){ const last=LO-+h; return D.years.map((y,i)=>[y,i]).filter(([y])=>y>=S.y1&&y<=S.y2&&y<=last).map(([,i])=>i); }
function cube(h, dis, ys){ const c=D.cube[String(h)], v=[0,0,0,0,0,0];
  for(const d of dis) for(const y of ys) for(const o of ALLO) for(let s=0;s<6;s++) v[s]+=c[d][y][o][s]; return v; }
function st(v){ const n=v.reduce((a,b)=>a+b,0), vis=v[0]+v[1]+v[2]+v[3]; return {v,n,vis,fr:v[0]+v[1],ab:v[2]+v[3],none:v[4]+v[5]}; }
const W0=el=>el.clientWidth||(document.querySelector(".fr-container").clientWidth-80)||900;

/* ---------- graphiques ---------- */
function stackRows(el, rows, labs, cols, o){
  if(!rows.length){ el.innerHTML='<div class="dd-empty">Effectifs insuffisants avec ces filtres.</div>'; return; }
  const W=Math.max(300,W0(el)), rh=28, nar=W<560, m={l:nar?Math.min(150,W*.38):200,r:nar?56:130,t:2,b:22};
  const H=m.t+m.b+rh*rows.length+6, iw=W-m.l-m.r, maxC=Math.floor((m.l-10)/6.3), cut=t=>t.length>maxC?t.slice(0,maxC-1)+"…":t;
  let s=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.label)}">`;
  [0,25,50,75,100].forEach(v=>{const x=m.l+iw*v/100; s+=`<line class="grid" x1="${x}" x2="${x}" y1="${m.t}" y2="${H-m.b}"/><text x="${x}" y="${H-6}" text-anchor="middle" class="m">${v} %</text>`;});
  rows.forEach((r,k)=>{
    const tot=r.v.reduce((a,b)=>a+b,0), y=m.t+k*rh+(r.total?6:0)+4, h=rh-10; let acc=0;
    if(r.total) s+=`<line class="axis" x1="0" x2="${W}" y1="${y-5}" y2="${y-5}"/>`;
    s+=`<text x="${m.l-8}" y="${y+h/2+4}" text-anchor="end" class="${r.sel?"sel":r.total?"v":""}">${esc(cut(r.label))}</text>`;
    r.v.forEach((v,j)=>{ if(!v) return; const w=iw*v/tot; s+=`<rect x="${m.l+acc}" y="${y}" width="${w}" height="${h}" fill="${cols[j]}" opacity="${S.disc!=="all"&&!r.sel&&!r.total?.4:1}"/>`; acc+=w; });
    const kk=o.key(r.v); s+=`<text x="${W-m.r+8}" y="${y+h/2+4}" class="v">${pct(kk[0],kk[1])}${nar?"":`<tspan class="m" dx="6" style="font-weight:400">${fmt.format(kk[0])} / ${fmt.format(kk[1])}</tspan>`}</text>`;
    s+=`<rect class="hit" x="0" y="${y-5}" width="${W}" height="${rh}" data-tip="${esc(`<h4>${esc(r.full||r.label)}</h4>`+r.v.map((v,j)=>row(labs[j],`${pct(v,tot,1)} · ${rat(v,tot)}`,cols[j])).join(""))}"/>`;
  });
  el.innerHTML=s+`</svg>`; bindTips(el);
}
function lineChart(el, pts, o){
  if(pts.length<2){ el.innerHTML='<div class="dd-empty">Pas assez de cohortes avec ces filtres.</div>'; return; }
  const W=Math.max(280,W0(el)), H=230, m={l:42,r:18,t:18,b:28}, iw=W-m.l-m.r, ih=H-m.t-m.b;
  const vals=pts.map(p=>100*p.n/p.N); let lo=Math.max(0,Math.floor((Math.min(...vals)-5)/10)*10), hi=Math.min(100,Math.ceil((Math.max(...vals)+5)/10)*10); if(hi-lo<20) hi=Math.min(100,lo+20);
  const xa=pts[0].y, xb=pts.at(-1).y, X=y=>m.l+iw*(xb===xa?.5:(y-xa)/(xb-xa)), Y=v=>m.t+ih*(1-(v-lo)/(hi-lo));
  let s=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.label)}">`;
  for(let v=lo;v<=hi;v+=10) s+=`<line class="grid" x1="${m.l}" x2="${W-m.r}" y1="${Y(v)}" y2="${Y(v)}"/><text x="${m.l-6}" y="${Y(v)+4}" text-anchor="end" class="m">${v} %</text>`;
  const every=Math.ceil(pts.length/Math.max(2,Math.floor(iw/46)));
  pts.forEach((p,i)=>{ if(i%every===0||i===pts.length-1) s+=`<text x="${X(p.y)}" y="${H-7}" text-anchor="middle" class="m">${p.y}</text>`; });
  const d=pts.map((p,i)=>(i?"L":"M")+X(p.y).toFixed(1)+","+Y(100*p.n/p.N).toFixed(1)).join("");
  s+=`<path d="${d} L${X(xb)},${m.t+ih} L${X(xa)},${m.t+ih}Z" fill="${o.col}" fill-opacity=".12"/><path d="${d}" fill="none" stroke="${o.col}" stroke-width="2.5" stroke-linejoin="round"/>`;
  const L=pts.at(-1); s+=`<circle cx="${X(L.y)}" cy="${Y(100*L.n/L.N)}" r="4" fill="${o.col}"/><text x="${X(L.y)-8}" y="${Y(100*L.n/L.N)-10}" text-anchor="end" class="v">${pct(L.n,L.N)}</text>`;
  const bw=iw/Math.max(1,pts.length-1);
  pts.forEach(p=>{ s+=`<rect class="hit" x="${X(p.y)-bw/2}" y="${m.t}" width="${bw}" height="${ih}" data-tip="${esc(`<h4>Cohorte ${p.y}</h4>${row(o.nLab,pct(p.n,p.N,1),o.col)}${row("Ratio",rat(p.n,p.N))}${row("Assiette",o.base)}`)}"/>`; });
  el.innerHTML=s+`</svg>`; bindTips(el);
}
function hbars(el, rows, o){
  if(!rows.length){ el.innerHTML='<div class="dd-empty">Aucune donnée avec ces filtres.</div>'; return; }
  const W=Math.max(300,W0(el)), rh=26, nar=W<560, m={l:nar?110:170,r:nar?70:160,t:2,b:4}, H=m.t+m.b+rh*rows.length, iw=W-m.l-m.r, mx=Math.max(...rows.filter(r=>!r.noBar).map(r=>r.n));
  let s=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.label)}">`;
  rows.forEach((r,i)=>{ const y=m.t+i*rh+4, h=rh-10, w1=iw*(r.n-(r.sub||0))/mx, w2=iw*(r.sub||0)/mx;
    s+=`<text x="${m.l-8}" y="${y+h/2+4}" text-anchor="end">${esc(r.label)}</text>`;
    if(!r.noBar){ s+=`<rect x="${m.l}" y="${y}" width="${Math.max(0,w1)}" height="${h}" fill="${o.col}"/>`; if(w2>0) s+=`<rect x="${m.l+w1}" y="${y}" width="${w2}" height="${h}" fill="var(--c-ret)"/>`; }
    const xx=r.noBar?m.l:m.l+w1+w2+8;
    s+=`<text x="${xx}" y="${y+h/2+4}" class="v">${pct(r.n,o.N,1)}${nar?"":`<tspan class="m" dx="6" style="font-weight:400">${fmt.format(r.n)} / ${fmt.format(o.N)}</tspan>`}</text>`;
    s+=`<rect class="hit" x="0" y="${y-4}" width="${W}" height="${rh}" data-tip="${esc(`<h4>${esc(r.label)}</h4>${row(o.nLab,`${pct(r.n,o.N,1)} · ${rat(r.n,o.N)}`)}${r.sub!=null?row("dont retours au pays",`${pct(r.sub,r.n)} · ${rat(r.sub,r.n)}`,"var(--c-ret)"):""}`)}"/>`; });
  el.innerHTML=s+`</svg>`; bindTips(el);
}
function forest(el, rows, o){
  if(!rows.length){ el.innerHTML='<div class="dd-empty">Aucun établissement avec assez de docteurs.</div>'; return; }
  const W=Math.max(300,W0(el)), rh=24, nar=W<560, m={l:nar?Math.min(160,W*.42):280,r:nar?56:140,t:4,b:26}, H=m.t+m.b+rh*rows.length, iw=W-m.l-m.r;
  const X=v=>m.l+iw*(Math.min(o.hi,Math.max(o.lo,v))-o.lo)/(o.hi-o.lo), maxC=Math.floor((m.l-10)/6.3), cut=t=>{ if(nar) t=t.replace(/^Université /,"Univ. "); return t.length>maxC?t.slice(0,maxC-1)+"…":t; };
  let s=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.label)}">`;
  o.ticks.forEach(v=>{ s+=`<line class="grid" x1="${X(v)}" x2="${X(v)}" y1="${m.t}" y2="${H-m.b}"/><text x="${X(v)}" y="${H-8}" text-anchor="middle" class="m">${o.ft(v)}</text>`; });
  s+=`<line x1="${X(o.ref)}" x2="${X(o.ref)}" y1="${m.t}" y2="${H-m.b}" stroke="var(--text-default-grey)" stroke-width="1.2" stroke-dasharray="4 3"/>`;
  rows.forEach((r,i)=>{ const y=m.t+i*rh+rh/2, col=r.lo>o.ref?"var(--c-ab)":r.hi<o.ref?"var(--c-fr)":"var(--c-int)", strong=col!=="var(--c-int)";
    s+=`<text x="${m.l-8}" y="${y+4}" text-anchor="end" class="${strong?"v":""}">${esc(cut(r.label))}</text>`;
    s+=`<line x1="${X(r.lo)}" x2="${X(r.hi)}" y1="${y}" y2="${y}" stroke="${col}" stroke-width="2" opacity=".6"/><rect x="${X(r.v)-4.5}" y="${y-4.5}" width="9" height="9" fill="${col}"/>`;
    s+=`<text x="${W-m.r+8}" y="${y+4}" class="${strong?"v":"m"}">${o.fv(r)}</text>`;
    s+=`<rect class="hit" x="0" y="${y-rh/2}" width="${W}" height="${rh}" data-tip="${esc(r.tip)}"/>`; });
  el.innerHTML=s+`</svg>`; bindTips(el);
}

/* ---------- 1 · où sont-ils ---------- */
function r1(){
  const h=S.h, ys=yIdx(h), s=st(cube(h,dList(),ys));
  applied("c1",{h:true,disc:true,coh:true});
  kpis("k1",[
    {l:"Docteurs analysables", v:fmt.format(s.n), s:`cohortes observables à ${h} ans, hors exclus`},
    {l:"Encore visibles", c:"var(--text-default-grey)", v:pct(s.vis,s.n), r:rat(s.vis,s.n), s:"affiliation récente / analysables"},
    {l:"En France", c:ST_COL[0], v:pct(s.fr,s.vis), r:rat(s.fr,s.vis), s:`en France / visibles · dont entreprise ${pct(s.v[1],s.vis)}`},
    {l:"À l'étranger", c:ST_COL[2], v:pct(s.ab,s.vis), r:rat(s.ab,s.vis), s:"à l'étranger / visibles"},
    {l:"Sans trace récente", c:ST_COL[4], v:pct(s.none,s.n), r:rat(s.none,s.n), s:"sans affiliation récente / analysables"},
  ]);
  $("n1").textContent="Assiette : docteurs analysables de chaque discipline. À droite : part encore visible (visibles / analysables). La discipline choisie dans les filtres est mise en avant.";
  $("lg1").innerHTML=ST_LAB.map((l,i)=>`<span><i class="dd-sw" style="background:${ST_COL[i]}"></i>${l}</span>`).join("");
  const vr=v=>(v[0]+v[1]+v[2]+v[3])/Math.max(1,v.reduce((a,b)=>a+b,0));
  const rows=D.disc.map((d,i)=>({label:dn(d),full:d,sel:S.disc===d,v:cube(h,[i],ys)})).filter(r=>r.v.reduce((a,b)=>a+b,0)>=30).sort((a,b)=>vr(b.v)-vr(a.v));
  rows.push({label:"Ensemble",total:true,v:cube(h,D.disc.map((_,i)=>i),ys)});
  stackRows($("ch1"),rows,ST_LAB,ST_COL,{label:"Statut par discipline",key:v=>[v[0]+v[1]+v[2]+v[3],v.reduce((a,b)=>a+b,0)]});
}

/* ---------- 2 · destinations ---------- */
function r2(){
  const h=S.h, ys=new Set(yIdx(h)), ds=new Set(dList()), by={}; let N=0,R=0,eu=0;
  for(const [c,d,y,o,n,r] of D.ctry[h]){ if(!ds.has(d)||!ys.has(y)) continue; (by[c]||(by[c]={n:0,r:0})).n+=n; by[c].r+=r; N+=n; R+=r; if(EUROPE.has(c)) eu+=n; }
  applied("c2",{h:true,disc:true,coh:true});
  const list=Object.entries(by).sort((a,b)=>b[1].n-a[1].n), top=list[0];
  kpis("k2",[
    {l:"Docteurs à l'étranger", v:fmt.format(N), s:`affiliés hors de France à ${h} ans`},
    {l:"Premier pays", c:"var(--c-ab)", v:top?pct(top[1].n,N):"–", r:top?`${cname(top[0])} · ${rat(top[1].n,N)}`:"", s:"premier pays / docteurs à l'étranger"},
    {l:"Europe", c:"var(--c-ab)", v:pct(eu,N), r:rat(eu,N), s:"en Europe / docteurs à l'étranger"},
    {l:"Retours au pays d'origine", c:"var(--c-ret)", v:pct(R,N), r:rat(R,N), s:"retours / docteurs à l'étranger"},
  ]);
  const rows=list.slice(0,12).map(([c,o])=>({label:cname(c),n:o.n,sub:o.r}));
  const rest=list.slice(12).reduce((a,[,o])=>a+o.n,0); if(rest) rows.push({label:`Autres (${list.length-12} pays)`,n:rest,noBar:true});
  hbars($("ch2"),rows,{label:"Pays de destination",col:"var(--c-ab)",N,nLab:"Part des docteurs à l'étranger"});
}

/* ---------- 3 · évolution ---------- */
function r3(){
  const h=S.h, ds=dList(), min=S.disc==="all"?300:60;
  applied("c3",{h:true,disc:true,coh:true});
  const pts=yIdx(h).map(i=>({y:D.years[i],s:st(cube(h,ds,[i]))})).filter(p=>p.y>=D.minY&&p.s.n>=min);
  lineChart($("ch3a"),pts.map(p=>({y:p.y,n:p.s.vis,N:p.s.n})),{label:"Encore visibles",col:"var(--c-fr)",nLab:"Encore visibles",base:"analysables de la cohorte"});
  lineChart($("ch3b"),pts.filter(p=>p.s.vis>=40).map(p=>({y:p.y,n:p.s.ab,N:p.s.vis})),{label:"À l'étranger",col:"var(--c-ab)",nLab:"À l'étranger",base:"visibles de la cohorte"});
}

/* ---------- 4 · partir et revenir ---------- */
function r4(){
  if(!AR) return;
  const ys=AR.years.map((y,i)=>[y,i]).filter(([y])=>y>=S.y1&&y<=S.y2).map(([,i])=>i);
  const sum=(arr,dis,k)=>{ const v=new Array(k).fill(0); for(const d of dis) for(const y of ys) for(const o of ALLO) for(let j=0;j<k;j++) v[j]+=arr[d][y][o][j]; return v; };
  const H8=D.horizons.at(-1), coh=ys.length?[AR.years[ys[0]],AR.years[ys.at(-1)]]:[AR.years[0],AR.years.at(-1)];
  applied("c4",{h:H8,disc:true,coh});
  const e=sum(AR.early,dList(),3), n=e[0]+e[1]+e[2];
  const dur=[0,0,0,0,0], ds=new Set(dList()), yset=new Set(ys); for(const [d,y,o,k,c] of AR.dur){ if(ds.has(d)&&yset.has(y)) dur[k]+=c; }
  const nd=dur.reduce((a,b)=>a+b,0);
  kpis("k4",[
    {l:"Partis tôt", v:fmt.format(n), s:`à l'étranger entre T+${D.lag+1} et T+${D.earlyH}, hors retours au pays d'origine`},
    {l:`Revenus en France à ${H8} ans`, c:"var(--c-fr)", v:pct(e[0],n), r:rat(e[0],n), s:"revenus / partis tôt"},
    {l:"Toujours à l'étranger", c:"var(--c-ab)", v:pct(e[1],n), r:rat(e[1],n), s:"à l'étranger / partis tôt"},
    {l:"Sans trace", c:"var(--c-int)", v:pct(e[2],n), r:rat(e[2],n), s:"sans affiliation récente / partis tôt"},
    {l:"Séjours de 3 ans ou moins", c:"var(--c-frp)", v:pct(dur[0]+dur[1]+dur[2],nd), r:rat(dur[0]+dur[1]+dur[2],nd), s:"séjours ≤ 3 ans / allers-retours"},
  ]);
  const labs=["Revenus en France","Toujours à l'étranger","Sans trace"], cols=["var(--c-fr)","var(--c-ab)","var(--c-none)"];
  $("lg4").innerHTML=labs.map((l,i)=>`<span><i class="dd-sw" style="background:${cols[i]}"></i>${l}</span>`).join("")+`<span style="color:var(--text-mention-grey)">à droite : revenus / partis tôt</span>`;
  const fr0=v=>v[0]/Math.max(1,v[0]+v[1]+v[2]);
  const rows=D.disc.map((d,i)=>({label:dn(d),full:d,sel:S.disc===d,v:sum(AR.early,[i],3)})).filter(r=>r.v.reduce((a,b)=>a+b,0)>=30).sort((a,b)=>fr0(b.v)-fr0(a.v));
  rows.push({label:"Ensemble",total:true,v:sum(AR.early,D.disc.map((_,i)=>i),3)});
  stackRows($("ch4a"),rows,labs,cols,{label:"Devenir des partis tôt",key:v=>[v[0],v[0]+v[1]+v[2]]});
  const el=$("ch4b"); if(nd<15){ el.innerHTML='<div class="dd-empty">Trop peu d\'allers-retours avec ces filtres.</div>'; return; }
  const lab=["1 an","2 ans","3 ans","4 ans","5 ans et +"], W=Math.max(280,W0(el)), H=230, m={l:42,r:10,t:22,b:28}, iw=W-m.l-m.r, ih=H-m.t-m.b, mx=Math.max(...dur.map(x=>100*x/nd)), top=Math.ceil(mx/10)*10, Y=p=>m.t+ih*(1-p/top), bw=iw/5;
  let s=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Durée à l'étranger">`;
  for(let p=0;p<=top;p+=10) s+=`<line class="grid" x1="${m.l}" x2="${W-m.r}" y1="${Y(p)}" y2="${Y(p)}"/><text x="${m.l-6}" y="${Y(p)+4}" text-anchor="end" class="m">${p} %</text>`;
  dur.forEach((x,i)=>{ const p=100*x/nd, X=m.l+i*bw+bw*.2, w=bw*.6;
    s+=`<rect x="${X}" y="${Y(p)}" width="${w}" height="${m.t+ih-Y(p)}" fill="${i<3?"var(--c-frp)":"var(--c-int)"}"/><text x="${X+w/2}" y="${Y(p)-6}" text-anchor="middle" class="v">${pct(x,nd)}</text><text x="${X+w/2}" y="${H-8}" text-anchor="middle" class="m">${lab[i]}</text>`;
    s+=`<rect class="hit" x="${m.l+i*bw}" y="${m.t}" width="${bw}" height="${ih}" data-tip="${esc(`<h4>${lab[i]} à l'étranger</h4>${row("Part des allers-retours",pct(x,nd,1))}${row("Ratio",rat(x,nd))}`)}"/>`; });
  el.innerHTML=s+`</svg>`; bindTips(el);
}

/* ---------- 5 · qui part, qui reste ---------- */
const SIG={net:["Signal net","fr-badge--success"],faible:["Signal faible","fr-badge--warning"],bruit:["Proche du hasard",""]};
function sig(d){ const v=PF.noise[d]; if(!v) return ["bruit","–"]; const x=v.tp[1]; return [x>=6.5?"net":x>=4?"faible":"bruit", `${x.toFixed(1).replace(".",",")} %`]; }
function r5(){
  if(!PF) return;
  $("pfh").textContent=PF.horizon;
  applied("c5",{h:PF.horizon,disc:true,coh:false});
  const one=S.disc!=="all" && PF.disciplines[S.disc];
  $("p5all").hidden=!!one; $("p5one").hidden=!one;
  const en=v=>[Math.round(v["these_en_anglais_partis_%"]*v.partis/100), Math.round(v["these_en_anglais_restés_%"]*v["restés"]/100)];
  if(!one){
    let h=`<thead><tr><th scope="col">Discipline</th><th scope="col">Partis / (partis + restés)</th><th scope="col">Thèse en anglais · partis</th><th scope="col">Thèse en anglais · restés</th><th scope="col">Signal thématique</th></tr></thead><tbody>`;
    let P=0,R=0,EP=0,ER=0;
    Object.values(PF.disciplines).sort((a,b)=>b["taux_depart_%"]-a["taux_depart_%"]).forEach(v=>{ const [ep,er]=en(v), [cl,x]=sig(v.discipline); P+=v.partis; R+=v["restés"]; EP+=ep; ER+=er;
      h+=`<tr data-d="${esc(v.discipline)}"><td>${esc(v.discipline)}</td><td>${pct(v.partis,v.partis+v["restés"])}<span class="dd-q">${rat(v.partis,v.partis+v["restés"])}</span></td><td>${pct(ep,v.partis)}<span class="dd-q">${rat(ep,v.partis)}</span></td><td>${pct(er,v["restés"])}<span class="dd-q">${rat(er,v["restés"])}</span></td><td><p class="fr-badge fr-badge--sm fr-badge--no-icon ${SIG[cl][1]}">${SIG[cl][0]} · ${x}</p></td></tr>`; });
    h+=`<tr><td><b>Ensemble</b></td><td><b>${pct(P,P+R)}</b><span class="dd-q">${rat(P,P+R)}</span></td><td><b>${pct(EP,P)}</b><span class="dd-q">${rat(EP,P)}</span></td><td><b>${pct(ER,R)}</b><span class="dd-q">${rat(ER,R)}</span></td><td></td></tr></tbody>`;
    $("t5").innerHTML=h;
    $("t5").querySelectorAll("tbody tr[data-d]").forEach(tr=>tr.addEventListener("click",()=>{ S.disc=tr.dataset.d; $("fDisc").value=S.disc; render(); }));
    return;
  }
  const v=PF.disciplines[S.disc], [ep,er]=en(v), [cl,x]=sig(S.disc), N=v.partis+v["restés"];
  kpis("k5",[
    {l:"Partis", c:"var(--c-ab)", v:pct(v.partis,N), r:rat(v.partis,N), s:"partis / (partis + restés)"},
    {l:"Thèse en anglais · partis", c:"var(--c-ab)", v:pct(ep,v.partis), r:rat(ep,v.partis), s:"thèses en anglais / partis"},
    {l:"Thèse en anglais · restés", c:"var(--c-fr)", v:pct(er,v["restés"]), r:rat(er,v["restés"]), s:"thèses en anglais / restés"},
    {l:"Signal thématique", v:SIG[cl][0], r:`${x} des topics nettement différents`, s:"4,6 % attendus par hasard"},
  ]);
  $("n5").textContent=`Score z : au-delà de 3, l'écart est fort (en gras). À droite : nombre de thèses concernées chez les partis (sur ${fmt.format(v.partis)}) · chez les restés (sur ${fmt.format(v["restés"])}).`;
  const P5=v[S.kw+"_partis"]||[], R5=v[S.kw+"_restes"]||[], mz=Math.max(3,...P5.map(t=>Math.abs(t[1])),...R5.map(t=>Math.abs(t[1])));
  const list=(id,items,col)=>{ const el=$(id); if(!items.length){ el.innerHTML='<p class="fr-text--sm" style="color:var(--text-mention-grey)">Aucun thème avec un écart notable.</p>'; return; }
    el.innerHTML=items.slice(0,12).map(([t,z,a,b,na,nb])=>{ const strong=Math.abs(z)>=3;
      return `<div class="dd-zr" data-tip="${esc(`<h4>${esc(t)}</h4>${row("Score z",z.toFixed(1).replace(".",","))}${row("Partis",`${a.toFixed(1).replace(".",",")} % · ${rat(na,v.partis)}`,"var(--c-ab)")}${row("Restés",`${b.toFixed(1).replace(".",",")} % · ${rat(nb,v["restés"])}`,"var(--c-fr)")}`)}"><span class="dd-zt${strong?" strong":""}">${esc(t)}</span><span class="dd-zb"><i style="width:${Math.max(4,100*Math.abs(z)/mz)}%;background:${col};opacity:${strong?1:.55}"></i></span><span class="dd-zv">${Math.abs(z).toFixed(1).replace(".",",")}</span><span class="dd-zs">${fmt.format(na)} · ${fmt.format(nb)}</span></div>`; }).join(""); bindTips(el); };
  list("z5p",P5,"var(--c-ab)"); list("z5r",R5,"var(--c-fr)");
}

/* ---------- 6 · établissements ---------- */
const ETS={"École des hautes études en sciences sociales":"EHESS","Institut National des Sciences Appliquées de Lyon":"INSA Lyon","Institut National des Sciences Appliquées de Toulouse":"INSA Toulouse","Institut Supérieur de l'Aéronautique et de l'Espace (ISAE-SUPAERO)":"ISAE-SUPAERO","Conservatoire National des Arts et Métiers":"Cnam","Institut National Polytechnique de Toulouse":"INP Toulouse"};
const etn=n=>ETS[n]||n;
function r6(){
  if(!ET) return;
  applied("c6",{h:PF?PF.horizon:D.mainH,disc:true,coh:false});
  const nar=W0($("ch6"))<560;
  if(S.disc==="all" || ET.base[S.disc]==null){
    $("h6").textContent="Chaque établissement est comparé au taux de départ moyen des disciplines où il forme : un ratio de 1,3 signifie 30 % de départs de plus qu'attendu.";
    $("t6").textContent="Départs observés / départs attendus, à discipline égale";
    $("n6").textContent=`Assiette : docteurs partis et restés des couples établissement × discipline d'au moins ${ET.min_etab} docteurs. À droite : écart à la moyenne, puis départs observés / attendus. Établissements avec au moins 20 départs attendus. Choisis une discipline pour voir les taux bruts.`;
    const rows=ET.inst.filter(r=>r[3]>=20).map(([i,n,o,e,oe,lo,hi,nd])=>({label:etn(i),v:oe,lo,hi,o,e,
      tip:`<h4>${esc(i)}</h4>${row("Observés / attendus",`${fmt.format(o)} / ${fmt.format(Math.round(e))}`)}${row("Ratio",oe.toFixed(2).replace(".",","))}${row("Intervalle à 95 %",`${lo.toFixed(2).replace(".",",")} – ${hi.toFixed(2).replace(".",",")}`)}${row("Docteurs (partis + restés)",fmt.format(n))}${row("Disciplines",nd)}`}));
    const hi=Math.max(2,Math.ceil(Math.max(...rows.map(r=>Math.min(r.hi,4)))*2)/2);
    forest($("ch6"),rows,{label:"Ratio par établissement",lo:0,hi,ref:1,ticks:[0,.5,1,1.5,2,2.5,3,3.5,4].filter(t=>t<=hi),ft:v=>v===1?"1 = moyenne":String(v).replace(".",","),
      fv:r=>{ const d=Math.round(100*(r.v-1)); return `${d>0?"+":d<0?"−":""}${Math.abs(d)} %`+(nar?"":`   ${fmt.format(r.o)} / ${fmt.format(Math.round(r.e))}`); }});
    return;
  }
  const d=S.disc, base=ET.base[d];
  $("h6").textContent=`Taux de départ de chaque établissement en ${d.toLowerCase()}, comparé à la moyenne de la discipline.`;
  $("t6").textContent=`Partis / (partis + restés) par établissement · ${d}`;
  $("n6").textContent=`Assiette : docteurs partis et restés de l'établissement dans cette discipline (au moins ${ET.min_etab}). Pointillés : moyenne de la discipline, ${String(base).replace(".",",")} % (${fmt.format(ET.tot[d])} docteurs).`;
  const rows=ET.cells.filter(c=>c[1]===d).sort((a,b)=>b[4]-a[4]).map(([i,_,n,p,r,lo,hi])=>({label:etn(i),v:r,lo,hi,n,p,
    tip:`<h4>${esc(i)}</h4>${row("Taux de départ",`${String(r).replace(".",",")} %`)}${row("Partis / docteurs",rat(p,n))}${row("Intervalle à 95 %",`${String(lo).replace(".",",")} – ${String(hi).replace(".",",")} %`)}${row("Moyenne de la discipline",`${String(base).replace(".",",")} %`)}`}));
  if(!rows.length){ $("ch6").innerHTML='<div class="dd-empty">Aucun établissement avec assez de docteurs dans cette discipline.</div>'; return; }
  const top=Math.min(100,Math.ceil((Math.max(...rows.map(r=>r.hi))+2)/10)*10), lo=Math.max(0,Math.floor((Math.min(...rows.map(r=>r.lo))-2)/10)*10), stp=nar?20:10, ticks=[]; for(let v=Math.ceil(lo/stp)*stp;v<=top;v+=stp) ticks.push(v);
  forest($("ch6"),rows,{label:"Taux de départ",lo,hi:top,ref:base,ticks,ft:v=>v+" %",fv:r=>`${Math.round(r.v)} %`+(nar?"":`   ${r.p} / ${r.n}`)});
}


/* ---------- explorer ---------- */
const EX = /*EX*/null;
const X={theme:"",ctry:"",typ:"",name:"",page:0}, PER=50;
const IA_RE="artificial intelligence|machine learning|deep learning|neural network|reinforcement learning|natural language processing|computer vision|large language model|intelligence artificielle|apprentissage automatique";
let xIdx=null;
function xInit(){
  if(!EX){ $("tab-8").parentElement.hidden=true; return; }
  const cc={}; EX.c.forEach(s=>{ if(s) s.split(",").forEach(c=>cc[c]=(cc[c]||0)+1); });
  $("xCtry").innerHTML=`<option value="">Tous les pays</option>`+Object.entries(cc).sort((a,b)=>b[1]-a[1]).map(([c,n])=>`<option value="${c}">${esc(cname(c))} (${fmt.format(n)})</option>`).join("");
  $("xTyp").innerHTML=`<option value="">Toutes</option>`+EX.typ.map((t,i)=>`<option value="${i}">${esc(t)}</option>`).join("");
  const deb=(f)=>{ let h; return ()=>{ clearTimeout(h); h=setTimeout(f,250); }; };
  $("xTheme").addEventListener("input",deb(()=>{ X.theme=$("xTheme").value.trim(); X.page=0; r8(); }));
  $("xName").addEventListener("input",deb(()=>{ X.name=$("xName").value.trim(); X.page=0; r8(); }));
  $("xCtry").addEventListener("change",e=>{ X.ctry=e.target.value; X.page=0; r8(); });
  $("xTyp").addEventListener("change",e=>{ X.typ=e.target.value; X.page=0; r8(); });
  $("xIA").addEventListener("click",()=>{ $("xTheme").value=IA_RE; X.theme=IA_RE; X.page=0; r8(); });
  $("xReset").addEventListener("click",()=>{ Object.assign(X,{theme:"",ctry:"",typ:"",name:"",page:0}); $("xTheme").value=""; $("xName").value=""; $("xCtry").value=""; $("xTyp").value=""; r8(); });
  $("xCsv").addEventListener("click",xCsv);
  $("xPages").addEventListener("click",e=>{ const a=e.target.closest("[data-p]"); if(!a) return; e.preventDefault(); X.page=+a.dataset.p; r8(); $("xCount").scrollIntoView({block:"start"}); });
}
function xFilter(){
  const ds=S.disc==="all"?null:D.disc.indexOf(S.disc);
  let re=null; if(X.theme){ try{ re=new RegExp(X.theme,"i"); }catch(e){ re=new RegExp(X.theme.replace(/[.*+?^${}()[\]\\]/g,"\\$&"),"i"); } }
  const tpHit=re?EX.vtp.map(t=>re.test(t)):null, kwHit=re?EX.vkw.map(t=>re.test(t)):null;
  const nm=X.name?X.name.toLowerCase():"";
  const out=[];
  for(let i=0;i<EX.n.length;i++){
    if(ds!==null&&EX.d[i]!==ds) continue;
    const y=EX.y[i]; if(y<S.y1||y>S.y2) continue;
    if(X.typ!==""&&EX.t[i]!==+X.typ) continue;
    if(X.ctry&&!(","+EX.c[i]+",").includes(","+X.ctry+",")) continue;
    if(nm&&!EX.n[i].toLowerCase().includes(nm)) continue;
    if(re&&!(EX.tp[i].some(k=>tpHit[k])||EX.kw[i].some(k=>kwHit[k]))) continue;
    out.push(i);
  }
  out.sort((a,b)=>EX.y[b]-EX.y[a]||EX.n[a].localeCompare(EX.n[b],"fr"));
  return out;
}
function seqHtml(i){
  const y0=EX.y[i]+EX.lag+1, home=new Set(EX.h[i]?EX.h[i].split(","):[]), cs=EX.c[i]?EX.c[i].split(","):[];
  return `<span class="dd-seq">`+[...EX.q[i]].map((ch,k)=>{ const yr=y0+k; const col=ch==="F"?"var(--c-fr)":ch==="E"?"var(--c-ab)":ch==="O"?"var(--c-abp)":"var(--c-none)";
    return `<i style="background:${col}" title="${yr} : ${ch==="F"?"France":ch==="E"?"étranger":ch==="O"?"pays d'origine":"aucune affiliation"}"></i>`; }).join("")+`</span>`;
}
function r8(){
  if(!EX) return;
  applied("c8",{h:"–",disc:true,coh:[S.y1,S.y2]});
  $("c8").querySelector(".dd-off")?.remove();
  xIdx=xFilter(); const n=xIdx.length, pages=Math.max(1,Math.ceil(n/PER)); if(X.page>=pages) X.page=pages-1;
  $("xCount").textContent=`${fmt.format(n)} docteur${n>1?"s":""}`;
  $("xMeta").textContent=X.theme?`Le filtre Thème ne porte que sur les thèses dont les thèmes OpenAlex ont été collectés (${pct(Math.round(EX.meta_cov*1000),1000)} des docteurs de la liste).`:"";
  const hs=EX.horizons;
  let h=`<thead><tr><th scope="col">Docteur</th><th scope="col">Soutenance</th><th scope="col">Discipline</th><th scope="col">Parcours T+${EX.lag+1} → T+10</th><th scope="col">Pays à l'étranger</th><th scope="col">Trajectoire</th>${hs.map(k=>`<th scope="col">${k} ans</th>`).join("")}<th scope="col">Thèse</th></tr></thead><tbody>`;
  const SL={F:"France",f:"France (entr.)",E:"Étranger",e:"Étranger (entr.)",I:"Interrompue",N:"Sans trace","-":"–",X:"exclu"};
  xIdx.slice(X.page*PER,(X.page+1)*PER).forEach(i=>{
    const home=new Set(EX.h[i]?EX.h[i].split(","):[]);
    const cs=EX.c[i]?EX.c[i].split(",").map(c=>`${esc(cname(c))}${home.has(c)?" <span class=\"fr-text--xs\" style=\"color:var(--text-mention-grey)\">(origine)</span>":""}`).join(", "):"–";
    h+=`<tr><td><a href="https://openalex.org/${EX.a[i]}" target="_blank" rel="noopener" title="Profil OpenAlex">${esc(EX.n[i]||EX.a[i])}</a></td><td>${EX.y[i]}</td><td>${esc(dn(D.disc[EX.d[i]]))}</td><td>${seqHtml(i)}</td><td class="wrap">${cs}</td><td>${esc(EX.typ[EX.t[i]])}</td>${[...EX.s[i]].map(ch=>`<td>${SL[ch]||ch}</td>`).join("")}<td><a href="https://openalex.org/${EX.w[i]}" target="_blank" rel="noopener">${EX.w[i]}</a></td></tr>`;
  });
  if(!n) h+=`<tr><td colspan="${7+hs.length}">Aucun docteur ne correspond à ces filtres.</td></tr>`;
  $("xTable").innerHTML=h+"</tbody>";
  const p=X.page, btn=(k,lab,cls)=>`<li><a class="fr-pagination__link ${cls||""}" href="#" data-p="${k}" ${k===p?'aria-current="page"':""} title="Page ${k+1}">${lab}</a></li>`;
  let pg=""; if(pages>1){ pg+=p>0?btn(p-1,"Précédente","fr-pagination__link--prev fr-pagination__link--lg-label"):""; 
    const set=new Set([0,pages-1,p-1,p,p+1].filter(k=>k>=0&&k<pages)); let last=-1;
    [...set].sort((a,b)=>a-b).forEach(k=>{ if(last>=0&&k>last+1) pg+=`<li><span class="fr-pagination__link fr-displayed-lg">…</span></li>`; pg+=btn(k,String(k+1)); last=k; });
    pg+=p<pages-1?btn(p+1,"Suivante","fr-pagination__link--next fr-pagination__link--lg-label"):""; }
  $("xPages").innerHTML=pg;
}
function xCsv(){
  if(!xIdx) return;
  const hs=EX.horizons, q=v=>{ v=String(v??""); return /[";\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; };
  const head=["nom","profil_openalex","these_openalex","annee_soutenance","discipline","trajectoire","pays_etranger","pays_origine","parcours_annuel_debut","parcours_annuel",...hs.map(k=>`statut_${k}_ans`),"topics","mots_cles"];
  const lines=[head.join(";")];
  xIdx.forEach(i=>lines.push([EX.n[i],"https://openalex.org/"+EX.a[i],"https://openalex.org/"+EX.w[i],EX.y[i],D.disc[EX.d[i]],EX.typ[EX.t[i]],EX.c[i],EX.h[i],EX.y[i]+EX.lag+1,EX.q[i],...[...EX.s[i]],EX.tp[i].map(k=>EX.vtp[k]).join(" | "),EX.kw[i].map(k=>EX.vkw[k]).join(" | ")].map(q).join(";")));
  const blob=new Blob(["\ufeff"+lines.join("\n")],{type:"text/csv;charset=utf-8"}), url=URL.createObjectURL(blob), a=document.createElement("a");
  a.href=url; a.download="docteurs_selection.csv"; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),2000);
}
/* ---------- initialisation ---------- */
$("segHel").innerHTML=D.horizons.map(h=>`<div class="fr-segmented__element"><input value="${h}" ${String(h)===S.h?"checked":""} type="radio" id="h-${h}" name="horizon"><label class="fr-label" for="h-${h}">${h} ans</label></div>`).join("");
$("fDisc").innerHTML=`<option value="all">Toutes les disciplines</option>`+D.disc.map(d=>`<option value="${esc(d)}">${esc(d)}</option>`).join("");
const yo=D.years.map(y=>`<option value="${y}">${y}</option>`).join(""); $("fY1").innerHTML=yo; $("fY2").innerHTML=yo; $("fY1").value=String(Y0); $("fY2").value=String(LO);
$("segH").addEventListener("change",e=>{ if(e.target.name==="horizon"){ S.h=e.target.value; render(); } });
$("segK").addEventListener("change",e=>{ if(e.target.name==="kind"){ S.kw=e.target.value; r5(); } });
$("fDisc").addEventListener("change",e=>{ S.disc=e.target.value; render(); });
$("fY1").addEventListener("change",e=>{ S.y1=+e.target.value; if(S.y1>S.y2){ S.y2=S.y1; $("fY2").value=S.y2; } render(); });
$("fY2").addEventListener("change",e=>{ S.y2=+e.target.value; if(S.y2<S.y1){ S.y1=S.y2; $("fY1").value=S.y1; } render(); });
$("reset").addEventListener("click",()=>{ Object.assign(S,{h:String(D.mainH),disc:"all",y1:Y0,y2:LO}); $("fDisc").value="all"; $("fY1").value=String(Y0); $("fY2").value=String(LO);
  const r=$("h-"+S.h); if(r) r.checked=true; render(); });
[[AR,"tab-4"],[PF,"tab-5"],[ET,"tab-6"]].forEach(([o,id])=>{ if(!o) $(id).parentElement.hidden=true; });
// onglets : le DSFR gère l'affichage ; repli si son script n'est pas chargé (hors ligne)
document.querySelectorAll(".fr-tabs__tab").forEach(t=>t.addEventListener("click",()=>{
  setTimeout(()=>{
    if(!window.dsfr){ document.querySelectorAll(".fr-tabs__tab").forEach(x=>{ const on=x===t; x.setAttribute("aria-selected",on); x.tabIndex=on?0:-1; $(x.getAttribute("aria-controls")).classList.toggle("fr-tabs__panel--selected",on); }); }
    render();
  },60);
}));
xInit();
function render(){ r1(); r2(); r3(); r4(); r5(); r6(); r8(); }
let rz; addEventListener("resize",()=>{ clearTimeout(rz); rz=setTimeout(render,150); });
render();
</script>
'''


if __name__ == "__main__":
    main()
