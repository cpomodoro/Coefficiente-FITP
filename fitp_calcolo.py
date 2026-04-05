#!/usr/bin/env python3
"""
Calcolatore Classifica FITP 2027
Uso: python fitp_calcolo.py <file_excel> [--classifica 3.4] [--sesso M] [--bonus-camp 0]
"""

import sys
import argparse
import math
from pathlib import Path
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    sys.exit("Errore: installa pandas con  pip install pandas openpyxl")

# ---------------------------------------------------------------------------
# Costanti regolamento
# ---------------------------------------------------------------------------

# Ordine dal MIGLIORE (indice 0) al PEGGIORE (indice alto)
CLASSI = [
    "2.1","2.2","2.3","2.4","2.5","2.6","2.7","2.8",
    "3.1","3.2","3.3","3.4","3.5",
    "4.1","4.2","4.3","4.4","4.5","4.6","4.NC"
]
GRUPPI = {c: i for i, c in enumerate(CLASSI)}

# Vittorie base per classifica (Tabella 1)
VBASE = {
    "4.NC":6,"4.6":6,"4.5":6,"4.4":6,"4.3":7,"4.2":7,"4.1":7,
    "3.5":8,"3.4":8,"3.3":8,"3.2":9,"3.1":9,
    "2.8":10,"2.7":10,"2.6":11,"2.5":11,"2.4":12,"2.3":14,"2.2":15,"2.1":16,
}

# Soglie [promo_M, promo_F, retro_M, retro_F] — None = non applicabile
SOGLIE = {
    "4.NC":[80,80,None,None],
    "4.6":[110,110,60,60],     "4.5":[210,210,90,90],
    "4.4":[300,300,120,120],   "4.3":[380,380,190,170],
    "4.2":[450,450,240,220],   "4.1":[520,500,275,250],
    "3.5":[580,560,370,320],   "3.4":[610,570,400,345],
    "3.3":[640,600,420,380],   "3.2":[670,630,470,420],
    "3.1":[720,660,500,450],
    "2.8":[750,690,525,470],   "2.7":[770,720,560,510],
    "2.6":[820,760,600,550],   "2.5":[880,820,650,600],
    "2.4":[950,910,680,640],   "2.3":[1020,960,730,680],
    "2.2":[1080,1020,750,720], "2.1":[None,None,780,760],
}

# ---------------------------------------------------------------------------
# Funzioni di calcolo
# ---------------------------------------------------------------------------

def punt_vittoria(diff: int) -> int:
    """Punti per una vittoria in base al diff di graduatoria (negativo = avv. più forte)."""
    if diff <= -2: return 120
    if diff == -1: return 90
    if diff == 0:  return 60
    if diff == 1:  return 30
    if diff == 2:  return 20
    if diff == 3:  return 15
    return 0

def desc_rel(diff: int, vittoria: bool) -> str:
    """Descrizione testuale della relazione di classifica."""
    if vittoria:
        if diff <= -2: return f"{abs(diff)} grad. sup. → 120 pt"
        if diff == -1: return "1 grad. sup. → 90 pt"
        if diff == 0:  return "pari → 60 pt"
        if diff == 1:  return "1 grad. inf. → 30 pt"
        if diff == 2:  return "2 grad. inf. → 20 pt"
        if diff == 3:  return "3 grad. inf. → 15 pt"
        return "4+ grad. inf. → 0 pt"
    else:
        if diff <= -1: return f"{abs(diff)} grad. sup. → non entra in formula"
        if diff == 0:  return "pari → E"
        if diff == 1:  return "1 grad. inf. → I"
        return "2+ grad. inf. → G"

def vitt_supplementari(formula_val: int, classe: str) -> int:
    """Calcola le vittorie supplementari dalla formula V-E-2I-3G (Tabella 2)."""
    c = classe
    if c in {"4.NC","4.6","4.5","4.4","4.3","4.2","4.1"}:
        if formula_val >= 21: return 4
        if formula_val >= 16: return 3
        if formula_val >= 11: return 2
        if formula_val >= 4:  return 1
        return 0
    if c in {"3.5","3.4","3.3","3.2","3.1"}:
        if formula_val >= 25: return 4
        if formula_val >= 19: return 3
        if formula_val >= 13: return 2
        if formula_val >= 5:  return 1
        return 0
    if c in {"2.8","2.7","2.6","2.5"}:
        if formula_val < 0:
            if formula_val <= -21: return -2
            if formula_val <= -5:  return -1
            return 0
        if formula_val >= 29: return 4
        if formula_val >= 22: return 3
        if formula_val >= 15: return 2
        if formula_val >= 6:  return 1
        return 0
    # 2.4 – 2.1
    if formula_val < 0:
        if formula_val <= -31: return -2
        if formula_val <= -15: return -1
        return 0
    if formula_val >= 41: return 5
    if formula_val >= 33: return 4
    if formula_val >= 25: return 3
    if formula_val >= 17: return 2
    if formula_val >= 7:  return 1
    return 0

def next_classe(c: str):
    i = CLASSI.index(c)
    return CLASSI[i+1] if i < len(CLASSI)-1 else None

def prev_classe(c: str):
    i = CLASSI.index(c)
    return CLASSI[i-1] if i > 0 else None

# ---------------------------------------------------------------------------
# Normalizzazione classifica da float Excel (3.4 → "3.4", 4.1 → "4.1")
# ---------------------------------------------------------------------------

def normalizza_classifica(val) -> str:
    if pd.isna(val):
        return None
    s = str(val).strip()
    # rimuovi .0 finale se presente (es. "3.40" → "3.4")
    try:
        f = float(s)
        # Gestisci 4.NC
        if s.upper() == "4.NC":
            return "4.NC"
        # Converti in stringa pulita: "3.4", "4.1", ecc.
        # I valori nel file sono tipo 3.4, 3.5, 4.1 — mantieni una decimale
        formatted = f"{f:.1f}"
        if formatted in GRUPPI:
            return formatted
        # prova anche senza decimale (es. "4.NC")
        return s
    except ValueError:
        return s.upper() if s.upper() == "4.NC" else s

# ---------------------------------------------------------------------------
# Parsing del file Excel
# ---------------------------------------------------------------------------

ESITI_VALIDI = {
    "W", "L",
    "W per assenza avv.", "W per assenza",
    "L per assenza", "L per assenza propria",
    "L per ritiro", "W per ritiro",
}

def normalizza_esito(s: str) -> str:
    """
    Normalizza le varianti di esito al codice interno W/WA/WR/L/LR/LA.
    Supporta sia il vecchio formato (W, L, W per assenza avv., ...)
    sia il nuovo formato (Win, Loss, Win - assenza avv., Win - ritiro avv.,
    Loss - assenza mia, Loss - ritiro mio).
    """
    s = str(s).strip()
    sl = s.lower()
    is_win  = sl.startswith("w")
    is_loss = sl.startswith("l")
    has_assenza = "assenza" in sl
    has_ritiro  = "ritiro"  in sl
    if is_win:
        if has_assenza: return "WA"
        if has_ritiro:  return "WR"
        return "W"
    if is_loss:
        if has_assenza: return "LA"
        if has_ritiro:  return "LR"
        return "L"
    return s.upper()

def normalizza_vet(s: str) -> str:
    s = str(s).strip().lower()
    if s in ("no","","nan"): return "no"
    if any(x in s for x in ["30","35","40","45"]): return "over30_45"
    if any(x in s for x in ["50","55","60","65"]): return "over50_65"
    if any(x in s for x in ["70","75","80"]): return "over70_80"
    return "no"

def leggi_excel(path: str, classe_partenza: str, sesso: str, bonus_camp: int):
    """
    Legge il file Excel e restituisce la lista di partite.
    Tollerante al formato: ignora le colonne non necessarie al calcolo,
    supporta nuove etichette esiti (Win, Win - assenza avv., Win - ritiro avv.,
    Loss, Loss - assenza mia, Loss - ritiro mio) e il vecchio formato.
    """
    try:
        df = pd.read_excel(path, sheet_name=0, header=0)
        # Verifica che le colonne abbiano senso; se no, prova con header nella riga 1
        cols_lower = [str(c).strip().lower() for c in df.columns]
        if not any(c in cols_lower for c in ["classifica", "esito"]):
            df = pd.read_excel(path, sheet_name=0, header=None)
            df.columns = [str(v).strip() if pd.notna(v) else f"col{i}"
                          for i, v in enumerate(df.iloc[0])]
            df = df.iloc[1:].reset_index(drop=True)
    except Exception as e:
        sys.exit(f"Errore lettura file Excel: {e}")

    # Mappa colonne (case-insensitive, tollerante agli spazi)
    col_map = {str(c).strip().lower(): c for c in df.columns}

    def get_col(*names):
        for n in names:
            if n.lower() in col_map:
                return col_map[n.lower()]
        return None

    col_n        = get_col("n.", "n", "#")
    col_cl       = get_col("classifica", "class.")
    col_esito    = get_col("esito", "risultato")
    col_tipo     = get_col("tipo", "type")
    col_fmt      = get_col("punteggio", "puteggio", "formato", "format")
    col_vet      = get_col("torneo veterani", "veterani", "vet")
    col_torneo   = get_col("vittoria torneo", "torneo vinto")
    col_migliore = get_col("classifica miglior partecipante avversario",
                           "miglior partecipante", "classifica miglior")
    col_npart    = get_col("numero partecipanti", "n partecipanti", "partecipanti")

    partite = []
    for _, row in df.iterrows():
        # Salta righe completamente vuote
        if all(pd.isna(v) or str(v).strip() in ("", "nan") for v in row):
            continue

        n = len(partite) + 1
        if col_n and pd.notna(row.get(col_n, float('nan'))):
            try:
                n = int(float(str(row[col_n])))
            except (ValueError, TypeError):
                pass

        cl    = normalizza_classifica(row[col_cl]) if col_cl else None
        esito = normalizza_esito(str(row[col_esito])) if col_esito else "L"
        tipo  = str(row[col_tipo]).strip().lower() if col_tipo else "singolare"
        ridotto = (str(row[col_fmt]).strip().lower() == "ridotto") if col_fmt else False

        vet_val = str(row.get(col_vet, "no")) if col_vet else "no"
        vet = normalizza_vet(vet_val)

        torneo_val = str(row.get(col_torneo, "no")) if col_torneo else "no"
        torneo_vinto = torneo_val.strip().lower() == "si"

        migliore = None
        if col_migliore:
            mig_val = row.get(col_migliore, float('nan'))
            if pd.notna(mig_val):
                migliore = normalizza_classifica(mig_val)

        n_part = None
        if col_npart:
            np_val = row.get(col_npart, float('nan'))
            if pd.notna(np_val):
                try:
                    n_part = int(float(str(np_val)))
                except (ValueError, TypeError):
                    pass

        if cl is None or cl not in GRUPPI:
            print(f"  [!] Riga {n}: classifica '{cl}' non riconosciuta, riga saltata.")
            continue

        partite.append({
            "n": n, "cl": cl, "esito": esito, "tipo": tipo,
            "ridotto": ridotto, "vet": vet,
            "torneo_vinto": torneo_vinto,
            "migliore": migliore, "n_part": n_part,
            "note": "",
        })

    return partite


# ---------------------------------------------------------------------------
# Calcolo coefficiente
# ---------------------------------------------------------------------------

def calcola_coeff(classe: str, sesso: str, partite: list, bonus_camp: int):
    """
    Calcola il coefficiente di rendimento per la classe data.
    Restituisce un dizionario con tutti i dettagli.
    """
    idx_promo = 0 if sesso == "M" else 1
    idx_retro = 2 if sesso == "M" else 3
    g_curr = GRUPPI[classe]
    n_base = VBASE[classe]
    cat = 2 if classe.startswith("2.") else 3 if classe.startswith("3.") else 4

    V = E = I = G = ass_pi = 0
    avv_pari_inf = 0
    sconfitte_pari_inf = False
    vitt_sing = []
    tornei_vinti = []

    for m in partite:
        if m["tipo"] not in ("singolare", "singles", "s"):
            continue  # doppio: non considerato per ora

        g_avv = GRUPPI[m["cl"]]
        diff = g_avv - g_curr        # negativo = avv. più forte
        is_pari_inf = (g_avv >= g_curr)
        is_migliore = (g_avv < g_curr)
        esito = m["esito"]

        # --- Sconfitta per assenza propria ---
        if esito == "LA":
            if is_pari_inf:
                ass_pi += 1
                if ass_pi >= 3:
                    sconfitte_pari_inf = True
            continue

        # --- Vittoria (normale o per ritiro avversario ad incontro iniziato) ---
        if esito in ("W", "WR"):
            V += 1
            p = punt_vittoria(diff)
            if m["vet"] == "over50_65": p = math.floor(p * 0.8)
            if m["vet"] == "over70_80": p = math.floor(p * 0.6)
            vitt_sing.append({
                "n": m["n"], "cl": m["cl"], "esito": esito,
                "ridotto": m["ridotto"], "valore_bruto": p, "diff": diff,
                "note": m["note"],
            })
            if is_pari_inf:
                avv_pari_inf += 1

        # --- Vittoria per assenza avversario ---
        elif esito == "WA":
            V += 1
            vitt_sing.append({
                "n": m["n"], "cl": m["cl"], "esito": "WA",
                "ridotto": m["ridotto"], "valore_bruto": 0, "diff": diff,
                "note": "Vittoria per assenza avv.: 0 pt, conta in V",
            })
            if is_pari_inf:
                avv_pari_inf += 1

        # --- Sconfitta (normale o ritiro mio ad incontro iniziato) ---
        elif esito in ("L", "LR"):
            if is_pari_inf:
                avv_pari_inf += 1
                sconfitte_pari_inf = True
            if not is_migliore:
                if diff == 0:   E += 1
                elif diff == 1: I += 1
                else:           G += 1

        # --- Torneo vinto ---
        if m["torneo_vinto"] and m["migliore"] and m["migliore"] in GRUPPI:
            tornei_vinti.append({
                "n": m["n"], "migliore": m["migliore"],
                "n_part": m["n_part"],
            })

    # Assenze pari/inf nella formula (dalla 3ª in poi)
    ass_use = max(0, ass_pi - 2)
    if ass_use >= 1: E += 1
    if ass_use >= 2: I += 1
    for k in range(3, ass_use + 1): G += 1

    formula_val = V - E - 2*I - 3*G
    n_suppl = vitt_supplementari(formula_val, classe)
    n_totali = max(0, n_base + n_suppl)

    # Ordina vittorie: valore desc, a parità intero prima
    vitt_sing.sort(key=lambda v: (-v["valore_bruto"], v["ridotto"]))

    # Calcola soglie e cap ridotto
    soglie = SOGLIE.get(classe, [None, None, None, None])
    promo_min = soglie[idx_promo]
    max_rid = math.floor(promo_min * 0.5) if promo_min else 9999

    # Identifica le usate e calcola taglio dal cap
    usate = vitt_sing[:n_totali]
    somma_ridotti = sum(v["valore_bruto"] for v in usate if v["ridotto"])
    ecc_rimanente = max(0, somma_ridotti - max_rid)

    # Applica taglio alle ridotte di minor valore tra le usate
    for v in reversed(usate):
        v["taglio"] = 0
        if v["ridotto"] and ecc_rimanente > 0:
            t = min(v["valore_bruto"], ecc_rimanente)
            v["taglio"] = t
            ecc_rimanente -= t

    # Assegna rank e punti effettivi
    punti_vitt = 0
    rid_prog = 0
    for i, v in enumerate(vitt_sing):
        v["rank"] = i + 1
        v["used"] = (i < n_totali)
        v["capped"] = v.get("taglio", 0) > 0
        if v["used"]:
            v["punti_eff"] = v["valore_bruto"] - v.get("taglio", 0)
            punti_vitt += v["punti_eff"]
            if v["ridotto"]:
                rid_prog += v["punti_eff"]
                v["rid_prog"] = rid_prog
        else:
            v["punti_eff"] = None
            v["taglio"] = 0

    # Bonus assenza sconfitte (art. 3.4a)
    ha_bonus_assenza = (classe != "4.NC") and (not sconfitte_pari_inf) and (avv_pari_inf >= 5)
    bonus_assenza = (50 if cat == 4 else 100) if ha_bonus_assenza else 0

    # Bonus vittoria torneo (art. 3.4b) — max 2
    bonus_tornei = 0
    tornei_detail = []
    min_part = 16 if sesso == "M" else 8
    for t in tornei_vinti[:2]:
        n_part = t["n_part"] or 0
        if n_part < min_part:
            tornei_detail.append({
                "n": t["n"], "migliore": t["migliore"],
                "n_part": n_part, "bonus": 0,
                "motivo": f"solo {n_part} partecipanti (servono ≥{min_part})",
            })
            continue
        g_mig = GRUPPI[t["migliore"]]
        diff_t = g_mig - g_curr
        p_base = punt_vittoria(diff_t)
        bonus = math.floor(p_base * 0.5)
        bonus_tornei += bonus
        tornei_detail.append({
            "n": t["n"], "migliore": t["migliore"],
            "n_part": n_part, "bonus": bonus,
            "p_base": p_base, "motivo": None,
        })

    coeff = punti_vitt + bonus_assenza + bonus_tornei + bonus_camp

    return {
        "classe": classe,
        "coeff": coeff,
        "punti_vitt": punti_vitt,
        "bonus_assenza": bonus_assenza,
        "ha_bonus_assenza": ha_bonus_assenza,
        "bonus_tornei": bonus_tornei,
        "tornei_detail": tornei_detail,
        "bonus_camp": bonus_camp,
        "V": V, "E": E, "I": I, "G": G,
        "ass_pi": ass_pi,
        "formula_val": formula_val,
        "n_base": n_base,
        "n_suppl": n_suppl,
        "n_totali": n_totali,
        "vitt_sing": vitt_sing,
        "promo_min": promo_min,
        "retro_max": soglie[idx_retro],
        "max_rid": max_rid,
        "rid_prog": rid_prog,
        "avv_pari_inf": avv_pari_inf,
        "sconfitte_pari_inf": sconfitte_pari_inf,
    }

def calcola_con_promozioni(classe_start: str, sesso: str, partite: list, bonus_camp: int):
    """Esegue il calcolo step-by-step con promozioni successive."""
    classe = classe_start
    risultato = None
    idx_promo = 0 if sesso == "M" else 1

    for _ in range(20):
        risultato = calcola_coeff(classe, sesso, partite, bonus_camp)
        promo_min = risultato["promo_min"]
        if promo_min and risultato["coeff"] >= promo_min:
            nxt = next_classe(classe)
            if nxt:
                classe = nxt
                continue
        break

    risultato["classe_start"] = classe_start
    risultato["classe_finale"] = classe

    # Determina esito
    g_f = GRUPPI[classe]
    g_s = GRUPPI[classe_start]
    retro_max = risultato["retro_max"]
    coeff = risultato["coeff"]

    if g_f > g_s:
        risultato["esito"] = f"PROMOZIONE a {classe}"
        risultato["esito_tipo"] = "promo"
    elif g_f < g_s:
        risultato["esito"] = f"RETROCESSIONE a {classe}"
        risultato["esito_tipo"] = "retro"
    elif retro_max and coeff <= retro_max:
        prv = prev_classe(classe)
        risultato["esito"] = f"RETROCESSIONE a {prv}" if prv else "RETROCESSIONE"
        risultato["esito_tipo"] = "retro"
    else:
        risultato["esito"] = f"MANTENIMENTO {classe}"
        risultato["esito_tipo"] = "stay"

    return risultato

# ---------------------------------------------------------------------------
# Output a schermo
# ---------------------------------------------------------------------------

def stampa_risultati(r: dict, partite: list):
    W = 100
    sep = "─" * W
    sep2 = "═" * W

    def line(s=""): print(s)
    def title(s): print(f"\n{sep2}\n  {s}\n{sep2}")
    def section(s): print(f"\n{sep}\n  {s}\n{sep}")

    title("CALCOLO CLASSIFICA FITP 2027")
    line(f"  Classifica di partenza : {r['classe_start']}")
    line(f"  Classifica conseguita  : {r['classe_finale']}")
    line(f"  Esito                  : {r['esito']}")

    section("RIEPILOGO NUMERICO")
    line(f"  Coefficiente totale    : {r['coeff']} pt")
    line(f"  di cui punti vittorie  : {r['punti_vitt']} pt")
    line(f"  di cui bonus totali    : {r['bonus_assenza'] + r['bonus_tornei'] + r['bonus_camp']} pt")
    line()
    line(f"  Formula V-E-2I-3G      : {r['V']} - {r['E']} - 2×{r['I']} - 3×{r['G']} = {r['formula_val']}")
    line(f"  Vittorie base          : {r['n_base']}")
    line(f"  Vittorie supplementari : {'+' if r['n_suppl']>=0 else ''}{r['n_suppl']}")
    line(f"  Vittorie considerate   : {r['n_totali']} (su {r['V']} totali)")
    line()
    promo = r['promo_min']
    retro = r['retro_max']
    line(f"  Soglia promozione      : {promo if promo else '— (già al massimo)'} pt")
    line(f"  Soglia retrocessione   : {retro if retro else '— (già al minimo)'} pt")
    if promo:
        mancanti = max(0, promo - r['coeff'])
        line(f"  Mancanti promozione    : {mancanti if mancanti else 'RAGGIUNTA!'} pt")
    if retro:
        margine = r['coeff'] - retro
        seg = f"+{margine} pt (al sicuro)" if margine > 0 else f"{margine} pt (IN RETROCESSIONE!)"
        line(f"  Margine su retrocessione: {seg}")

    section("DETTAGLIO PARTITE — ordinate per valore (le prime {n} sono usate nel calcolo)".format(n=r['n_totali']))

    hdr = f"{'Ord':>4}  {'#':>3}  {'Class':5}  {'Esito':<22}  {'Fmt':7}  {'Valore':>7}  {'Pt usati':>9}  Note"
    line(hdr)
    line("─" * W)

    sconfitte_display = []
    for v in r["vitt_sing"]:
        rank_s = f"[{v['rank']:>2}]" if v["used"] else f" {v['rank']:>2} "
        esito_s = {
            "W": "Vittoria", "WR": "Vittoria (ritiro avv.)",
            "WA": "Vittoria per assenza avv.",
        }.get(v["esito"], v["esito"])
        fmt_s = "Ridotto" if v["ridotto"] else "Intero"
        valore_s = f"{v['valore_bruto']:>5} pt"

        if v["used"]:
            if v["capped"]:
                pt_s = f"{v['punti_eff']:>5} pt (↓{v['taglio']})"
            else:
                pt_s = f"{v['punti_eff']:>5} pt"
        else:
            pt_s = f"{'—':>9}"

        note_parts = []
        note_parts.append(desc_rel(v["diff"], vittoria=True))
        if v["used"] and v["ridotto"] and "rid_prog" in v:
            note_parts.append(f"rid.cum.={v['rid_prog']}/{r['max_rid']}")
            if v["capped"]:
                note_parts.append("CAP APPLICATO")
        if v.get("note"):
            note_parts.append(v["note"])
        note_s = " | ".join(note_parts)

        flag = "*" if v["used"] else " "
        line(f"{flag}{rank_s}  {v['n']:>3}  {v['cl']:5}  {esito_s:<22}  {fmt_s:7}  {valore_s}  {pt_s:<12}  {note_s}")

    line()
    line("  * = vittoria usata nel calcolo")

    # Sconfitte
    line()
    line("  SCONFITTE E ASSENZE:")
    line("─" * W)
    sconfitte_partite = [m for m in partite if m["esito"] in ("L","LR","LA")]
    for m in sconfitte_partite:
        g_avv = GRUPPI[m["cl"]]
        diff = g_avv - GRUPPI[r["classe"]]
        esito_s = {
            "L": "Sconfitta", "LR": "Sconfitta (ritiro mio)",
            "LA": "Assenza propria",
        }.get(m["esito"], m["esito"])
        fmt_s = "Ridotto" if m["ridotto"] else "Intero"
        note_s = desc_rel(diff, vittoria=False)
        if m.get("note"): note_s += " | " + m["note"]
        line(f"       {m['n']:>3}  {m['cl']:5}  {esito_s:<22}  {fmt_s:7}  {'0':>7} pt  {'':9}  {note_s}")

    section("BONUS")
    # Bonus assenza sconfitte
    if r["ha_bonus_assenza"]:
        motivo = f"ASSEGNATO — {r['avv_pari_inf']} avv. pari/inf incontrati, nessuna sconfitta"
    else:
        if r["sconfitte_pari_inf"]:
            motivo = "non assegnato — presenti sconfitte vs pari/inf"
        elif r["avv_pari_inf"] < 5:
            motivo = f"non assegnato — solo {r['avv_pari_inf']}/5 avv. pari/inf richiesti"
        else:
            motivo = "non applicabile (4.NC)"
    line(f"  Assenza sconfitte (art. 3.4a): {r['bonus_assenza']:>4} pt  [{motivo}]")

    # Bonus torneo
    if r["tornei_detail"]:
        for t in r["tornei_detail"]:
            if t["motivo"]:
                line(f"  Torneo vinto #{ t['n'] } (art. 3.4b):  {t['bonus']:>4} pt  [non assegnato — {t['motivo']}]")
            else:
                line(f"  Torneo vinto #{ t['n'] } (art. 3.4b):  {t['bonus']:>4} pt  [miglior partecipante {t['migliore']}, {t['p_base']}pt × 50%]")
    else:
        line(f"  Vittoria torneo (art. 3.4b):   {0:>4} pt  [nessun torneo vinto indicato]")

    line(f"  Campionati individuali:        {r['bonus_camp']:>4} pt  [inserito manualmente]")

    # Barra cap ridotto
    section("CAP TORNEI A PUNTEGGIO RIDOTTO")
    rid = r["rid_prog"]
    max_r = r["max_rid"]
    pct = min(100, round(rid / max_r * 100)) if max_r < 9999 else 0
    bar_len = 50
    filled = round(pct / 100 * bar_len)
    bar = "█" * filled + "░" * (bar_len - filled)
    line(f"  [{bar}] {pct}%")
    line(f"  {rid} pt usati su {max_r} pt disponibili (50% di {r['promo_min'] or '?'})")
    if rid > max_r:
        line(f"  ATTENZIONE: limite superato di {rid - max_r} pt → applicato taglio")

    line(f"\n{'═'*W}\n")

# ---------------------------------------------------------------------------
# Output su file HTML (report leggibile)
# ---------------------------------------------------------------------------

def genera_html(r: dict, partite: list, output_path: str):
    sconfitte_partite = [m for m in partite if m["esito"] in ("L","LR","LA")]
    g_curr = GRUPPI[r["classe"]]

    esito_color = {"promo": "#27500A", "stay": "#185FA5", "retro": "#A32D2D"}[r["esito_tipo"]]
    esito_bg    = {"promo": "#EAF3DE", "stay": "#E6F1FB", "retro": "#FCEBEB"}[r["esito_tipo"]]

    def row_class(v):
        if not v["used"]: return "unused"
        if v["capped"]:   return "capped"
        return "used"

    # Righe vittorie
    vitt_rows = ""
    for v in r["vitt_sing"]:
        rc = row_class(v)
        rank_s = f'<span class="rank{"_cap" if v["capped"] else ""}">{v["rank"]}</span>' if v["used"] else f'<span style="color:#aaa">{v["rank"]}</span>'
        esito_s = {"W":"Vittoria","WR":"Vittoria (ritiro avv.)","WA":"Vittoria per assenza avv."}.get(v["esito"], v["esito"])
        pill = f'<span class="pill pill-w">{esito_s}</span>'
        fmt_s = "Ridotto" if v["ridotto"] else "Intero"

        if v["used"]:
            if v["capped"]:
                pt_s = f'<strong>{v["punti_eff"]} pt</strong><br><small>↓ tagliati {v["taglio"]} pt (cap)</small>'
            else:
                pt_s = f'<strong>{v["punti_eff"]} pt</strong>'
        else:
            pt_s = f'<span style="color:#aaa">{v["valore_bruto"]} pt (non usata)</span>'

        note_parts = [desc_rel(v["diff"], vittoria=True)]
        if v["used"] and v["ridotto"] and "rid_prog" in v:
            note_parts.append(f"rid. cum.: {v['rid_prog']}/{r['max_rid']} pt")
            if v["capped"]: note_parts.append("<strong>CAP APPLICATO</strong>")
        if v.get("note"): note_parts.append(v["note"])

        vitt_rows += f"""<tr class="{rc}">
          <td style="text-align:center">{rank_s}</td>
          <td style="color:#888;font-size:12px">{v['n']}</td>
          <td>{v['cl']}</td><td>{pill}</td>
          <td style="font-size:12px">{fmt_s}</td>
          <td>{v['valore_bruto']} pt</td>
          <td>{pt_s}</td>
          <td style="font-size:11px">{"&nbsp;·&nbsp;".join(note_parts)}</td>
        </tr>"""

    # Righe sconfitte
    loss_rows = ""
    for m in sconfitte_partite:
        g_avv = GRUPPI[m["cl"]]
        diff = g_avv - g_curr
        esito_s = {"L":"Sconfitta","LR":"Sconfitta (ritiro mio)","LA":"Assenza propria"}.get(m["esito"], m["esito"])
        pill_cls = "pill-la" if m["esito"] == "LA" else "pill-l"
        fmt_s = "Ridotto" if m["ridotto"] else "Intero"
        note_s = desc_rel(diff, vittoria=False)
        if m.get("note"): note_s += " · " + m["note"]
        loss_rows += f"""<tr class="loss">
          <td></td><td style="color:#888;font-size:12px">{m['n']}</td>
          <td>{m['cl']}</td>
          <td><span class="pill {pill_cls}">{esito_s}</span></td>
          <td style="font-size:12px">{fmt_s}</td>
          <td style="color:#aaa">0 pt</td><td></td>
          <td style="font-size:11px">{note_s}</td>
        </tr>"""

    # Bonus rows
    bonus_rows = ""
    if r["ha_bonus_assenza"]:
        motivo_ass = f"Assegnato — {r['avv_pari_inf']} avv. pari/inf, nessuna sconfitta"
        ba_color = "#27500A"
    else:
        motivo_ass = "Non assegnato — " + ("sconfitte vs pari/inf presenti" if r["sconfitte_pari_inf"] else f"solo {r['avv_pari_inf']}/5 avv. richiesti")
        ba_color = "#A32D2D"
    bonus_rows += f'<tr><td>Bonus assenza sconfitte (art. 3.4a)</td><td style="font-weight:500">{r["bonus_assenza"]} pt</td><td style="color:{ba_color}">{motivo_ass}</td></tr>'

    for t in r["tornei_detail"]:
        if t["motivo"]:
            bonus_rows += f'<tr><td>Torneo vinto #{t["n"]} (art. 3.4b)</td><td style="font-weight:500">0 pt</td><td style="color:#A32D2D">Non assegnato — {t["motivo"]}</td></tr>'
        else:
            bonus_rows += f'<tr><td>Torneo vinto #{t["n"]} (art. 3.4b)</td><td style="font-weight:500">{t["bonus"]} pt</td><td style="color:#27500A">Miglior partecipante {t["migliore"]}, {t["p_base"]}pt × 50%</td></tr>'
    if not r["tornei_detail"]:
        bonus_rows += f'<tr><td>Vittoria torneo (art. 3.4b)</td><td style="font-weight:500">0 pt</td><td style="color:#888">Nessun torneo vinto indicato</td></tr>'
    bonus_rows += f'<tr><td>Bonus campionati individuali</td><td style="font-weight:500">{r["bonus_camp"]} pt</td><td style="color:#888">Inserito manualmente</td></tr>'

    # Barra cap
    rid = r["rid_prog"]
    max_r = r["max_rid"]
    pct = min(100, round(rid / max_r * 100)) if max_r < 9999 else 0
    bar_color = "#E24B4A" if pct >= 100 else "#BA7517" if pct >= 80 else "#639922"

    promo = r["promo_min"]
    retro = r["retro_max"]
    mancanti = max(0, promo - r["coeff"]) if promo else None
    margine = r["coeff"] - retro if retro else None

    now = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = f"""<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Calcolo Classifica FITP 2027</title>
<style>
  body{{font-family:system-ui,Arial,sans-serif;max-width:1100px;margin:0 auto;padding:24px;color:#1a1a1a;background:#f5f5f0}}
  h1{{font-size:22px;font-weight:600;margin-bottom:4px}}
  .sub{{font-size:13px;color:#666;margin-bottom:24px}}
  .metrics{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:24px}}
  .metric{{background:#eee;border-radius:8px;padding:14px 16px}}
  .metric .lbl{{font-size:12px;color:#666;margin-bottom:4px}}
  .metric .val{{font-size:24px;font-weight:600}}
  .metric .sub2{{font-size:11px;color:#888;margin-top:3px}}
  .card{{background:#fff;border:1px solid #ddd;border-radius:10px;padding:16px 20px;margin-bottom:16px}}
  .sec{{font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.05em;margin-bottom:12px}}
  table{{width:100%;border-collapse:collapse;font-size:13px}}
  th{{text-align:left;padding:6px 10px;border-bottom:1px solid #eee;font-size:11px;color:#888;font-weight:600}}
  td{{padding:6px 10px;border-bottom:1px solid #f0f0f0;vertical-align:middle}}
  tr:last-child td{{border-bottom:none}}
  tr.used{{background:#C0DD97}}
  tr.used td{{color:#173404}}
  tr.capped{{background:#FAC775}}
  tr.capped td{{color:#412402}}
  tr.unused{{background:#f5f5f5}}
  tr.unused td{{color:#888}}
  tr.loss td{{color:#aaa}}
  .rank{{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:50%;font-size:12px;font-weight:600;background:#639922;color:#fff}}
  .rank_cap{{display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:50%;font-size:12px;font-weight:600;background:#BA7517;color:#fff}}
  .pill{{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}}
  .pill-w{{background:#27500A;color:#EAF3DE}}
  .pill-wa{{background:#0C447C;color:#E6F1FB}}
  .pill-l{{background:#791F1F;color:#FCEBEB}}
  .pill-la{{background:#633806;color:#FAEEDA}}
  .detail-row{{display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid #f0f0f0;font-size:13px}}
  .detail-row:last-child{{border-bottom:none}}
  .dlabel{{color:#666}}.dval{{font-weight:600}}
  .badge{{display:inline-block;padding:4px 16px;border-radius:8px;font-size:14px;font-weight:600;background:{esito_bg};color:{esito_color}}}
  .ridbar{{height:14px;border-radius:7px;background:#ddd;overflow:hidden;margin:8px 0}}
  .ridfill{{height:100%;border-radius:7px;background:{bar_color};width:{pct}%}}
  .sep-row td{{background:#f8f8f8;font-size:11px;color:#888;padding:4px 10px;font-style:italic}}
  .bonus-table td{{padding:7px 10px;border-bottom:1px solid #f0f0f0;font-size:13px}}
  .bonus-table tr:last-child td{{border-bottom:none}}
  @media(max-width:700px){{.metrics{{grid-template-columns:1fr 1fr}}}}
</style>
</head>
<body>
<h1>Calcolo Classifica FITP 2027</h1>
<div class="sub">Generato il {now}</div>

<div class="metrics">
  <div class="metric"><div class="lbl">Coefficiente totale</div><div class="val">{r['coeff']}</div><div class="sub2">punti</div></div>
  <div class="metric"><div class="lbl">Punti da vittorie</div><div class="val">{r['punti_vitt']}</div><div class="sub2">{r['n_totali']} vittorie su {r['V']} totali</div></div>
  <div class="metric"><div class="lbl">Formula V-E-2I-3G</div><div class="val">{r['formula_val']}</div><div class="sub2">V={r['V']} E={r['E']} I={r['I']} G={r['G']}</div></div>
  <div class="metric"><div class="lbl">Vittorie supplementari</div><div class="val">+{r['n_suppl']}</div><div class="sub2">base {r['n_base']} → tot {r['n_totali']}</div></div>
</div>

<div class="card">
<div class="sec">Dettaglio partite — ordinate per valore (le prime {r['n_totali']} sono usate nel calcolo)</div>
<table>
<thead><tr>
  <th style="width:36px">Ord.</th><th style="width:26px">#</th>
  <th style="width:48px">Class.</th><th style="width:180px">Esito</th>
  <th style="width:60px">Formato</th><th style="width:70px">Valore</th>
  <th style="width:100px">Pt usati</th><th>Note</th>
</tr></thead>
<tbody>
<tr class="sep-row"><td colspan="8">Le {r['n_totali']} vittorie migliori considerate ({r['n_base']} base + {r['n_suppl']} supplementari)</td></tr>
{vitt_rows}
<tr class="sep-row"><td colspan="8">Sconfitte e assenze</td></tr>
{loss_rows}
</tbody>
</table>
<div style="margin-top:12px">
  <div style="font-size:12px;color:#666">Punti da tornei ridotti — massimale {max_r} pt (50% di {promo or '?'})</div>
  <div class="ridbar"><div class="ridfill"></div></div>
  <div style="font-size:12px;color:#666">{rid} pt su {max_r} pt disponibili ({pct}%){' — LIMITE RAGGIUNTO' if pct >= 100 else ''}</div>
</div>
</div>

<div class="card">
<div class="sec">Bonus</div>
<table class="bonus-table"><tbody>{bonus_rows}</tbody></table>
</div>

<div class="card">
<div class="sec">Riepilogo e soglie</div>
<div class="detail-row"><span class="dlabel">Classifica di partenza</span><span class="dval">{r['classe_start']}</span></div>
<div class="detail-row"><span class="dlabel">Classifica conseguita</span><span class="dval">{r['classe_finale']}</span></div>
<div class="detail-row"><span class="dlabel">Soglia promozione</span><span class="dval">{promo if promo else '— (già al massimo)'} pt</span></div>
<div class="detail-row"><span class="dlabel">Soglia retrocessione</span><span class="dval">{retro if retro else '— (già al minimo)'} pt</span></div>
<div class="detail-row"><span class="dlabel">Punti mancanti promozione</span><span class="dval">{'RAGGIUNTA!' if mancanti == 0 else f'{mancanti} pt ancora necessari' if mancanti is not None else '—'}</span></div>
<div class="detail-row"><span class="dlabel">Margine su retrocessione</span><span class="dval" style="color:{'#27500A' if margine and margine > 0 else '#A32D2D'}">{f'+{margine} pt (al sicuro)' if margine and margine > 0 else f'{margine} pt (IN RETROCESSIONE!)' if margine is not None else '—'}</span></div>
<div class="detail-row"><span class="dlabel">Esito</span><span class="dval"><span class="badge">{r['esito']}</span></span></div>
</div>

</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Calcolatore Classifica FITP 2027",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Esempi:
  python fitp_calcolo.py Matches.xlsx
  python fitp_calcolo.py Matches.xlsx --classifica 3.4 --sesso M
  python fitp_calcolo.py Matches.xlsx --classifica 3.4 --bonus-camp 45 --output risultato.html
        """
    )
    parser.add_argument("file", help="File Excel con le partite")
    parser.add_argument("--classifica", default="3.4",
                        help="Classifica di partenza (default: 3.4)")
    parser.add_argument("--sesso", choices=["M","F"], default="M",
                        help="Sesso del giocatore: M o F (default: M)")
    parser.add_argument("--bonus-camp", type=int, default=0, dest="bonus_camp",
                        help="Punti bonus campionati individuali (default: 0)")
    parser.add_argument("--output", default=None,
                        help="Percorso file HTML di output (default: stesso nome del file .xlsx)")
    args = parser.parse_args()

    # Validazione classifica
    if args.classifica not in GRUPPI:
        sys.exit(f"Errore: classifica '{args.classifica}' non valida. Usa una tra: {', '.join(CLASSI)}")

    # Percorso output
    input_path = Path(args.file)
    output_path = args.output or str(Path(".") / (input_path.stem + "_risultato.html"))

    print(f"\nLettura file: {args.file}")
    partite = leggi_excel(args.file, args.classifica, args.sesso, args.bonus_camp)
    print(f"Lette {len(partite)} partite.")

    print(f"Calcolo classifica di partenza {args.classifica} (sesso {args.sesso})...")
    risultato = calcola_con_promozioni(args.classifica, args.sesso, partite, args.bonus_camp)

    stampa_risultati(risultato, partite)

    genera_html(risultato, partite, output_path)
    print(f"Report HTML salvato in: {output_path}\n")

if __name__ == "__main__":
    main()
