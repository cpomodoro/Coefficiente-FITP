import streamlit as st
import io
import contextlib
import tempfile
import os
from pathlib import Path

from fitp_calcolo import (
    CLASSI, GRUPPI, leggi_excel, calcola_con_promozioni,
    stampa_risultati, genera_html, desc_rel
)

st.set_page_config(
    page_title="Calcolatore Classifica FITP 2027",
    page_icon="🎾",
    layout="centered"
)

st.title("🎾 Calcolatore Classifica FITP 2027")

# ── Parametri ────────────────────────────────────────────────────────
st.subheader("Parametri")

col1, col2, col3 = st.columns(3)
with col1:
    classifica = st.selectbox("Classifica di partenza", CLASSI, index=CLASSI.index("3.4"))
with col2:
    sesso = st.radio("Sesso", ["M", "F"], horizontal=True)
with col3:
    bonus_camp = st.number_input("Bonus campionati (pt)", min_value=0, value=0, step=5)

# Template Excel incorporato come base64
_TEMPLATE_B64 = "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"

# ── Istruzioni formato file ──────────────────────────────────────────
with st.expander("ℹ️  Formato del file Excel — clicca per leggere le istruzioni"):
    st.markdown("""
Il file Excel deve contenere un foglio con una riga di intestazione e una riga per ogni partita.

**Colonne obbligatorie** (il nome deve corrispondere esattamente):

| Colonna | Valori accettati |
|---|---|
| `Classifica` | Classifica dell'avversario: `2.1` … `2.8`, `3.1` … `3.5`, `4.1` … `4.6`, `4.NC` |
| `Esito` | `Win` · `Win - assenza avv.` · `Win - ritiro avv.` · `Loss` · `Loss - assenza mia` · `Loss - ritiro mio` |
| `Tipo` | `Singolare` (il doppio non è considerato nel calcolo) |
| `Punteggio` | `Intero` oppure `Ridotto` |
| `Torneo Veterani` | `No` · `Over 30-45` · `Over 50-65` · `Over 70-80` |
| `Vittoria Torneo` | `Si` oppure `No` |

**Colonne necessarie solo se "Vittoria Torneo" = Si:**

| Colonna | Descrizione |
|---|---|
| `Classifica miglior partecipante avversario` | Classifica del miglior avversario che ha partecipato al torneo |
| `Numero partecipanti` | Numero totale di partecipanti effettivi (serve almeno 16 per il bonus maschile, 8 per il femminile) |

**Colonne opzionali** (vengono ignorate dal calcolo ma puoi includerle):
`n.` · `Data` · `Torneo` · `Superficie` · `Avversario` · `Età` · `Risultato` · qualsiasi altra colonna

**Note importanti:**
- `Win - assenza avv.` conta come vittoria nella formula V-E-2I-3G ma vale **0 punti**
- `Loss - ritiro mio` (ad incontro iniziato) vale come sconfitta normale
- `Loss - assenza mia` non entra nella formula per le prime 2 volte contro avversari pari/inferiori
- Le partite di doppio vengono ignorate
    """)
    import base64 as _b64
    _tpl_bytes = _b64.b64decode(_TEMPLATE_B64)
    st.download_button(
        label="⬇️  Scarica template Excel",
        data=_tpl_bytes,
        file_name="Template_Partite_FITP.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ── Caricamento file ─────────────────────────────────────────────────
st.subheader("File partite")
uploaded = st.file_uploader(
    "Carica il file Excel con le partite",
    type=["xlsx", "xls"],
)

# ── Calcola ──────────────────────────────────────────────────────────
if uploaded and st.button("Calcola", use_container_width=True):

    try:
        file_bytes = uploaded.read()
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name

        with st.spinner("Lettura file e calcolo in corso..."):
            partite = leggi_excel(tmp_path, classifica, sesso, bonus_camp)
            risultato = calcola_con_promozioni(classifica, sesso, partite, bonus_camp)

        # ── Metriche principali ───────────────────────────────────────
        st.divider()
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Coefficiente totale", f"{risultato['coeff']} pt")
        col2.metric("Punti da vittorie",   f"{risultato['punti_vitt']} pt")
        col3.metric("Formula V-E-2I-3G",   risultato['formula_val'],
                    help=f"V={risultato['V']}  E={risultato['E']}  I={risultato['I']}  G={risultato['G']}")
        col4.metric("Vittorie considerate",
                    f"{risultato['n_totali']}",
                    help=f"Base {risultato['n_base']} + supplementari {risultato['n_suppl']}")

        # ── Esito ─────────────────────────────────────────────────────
        tipo = risultato["esito_tipo"]
        esito_txt = risultato["esito"]
        if tipo == "promo":
            st.success(f"✅  {esito_txt}")
        elif tipo == "retro":
            st.error(f"⚠️  {esito_txt}")
        else:
            st.info(f"➡️  {esito_txt}")

        # ── Soglie ────────────────────────────────────────────────────
        promo = risultato["promo_min"]
        retro = risultato["retro_max"]
        coeff = risultato["coeff"]

        col1, col2, col3 = st.columns(3)
        col1.metric("Soglia promozione",     f"{promo} pt" if promo else "—")
        col2.metric("Soglia retrocessione",  f"{retro} pt" if retro else "—")
        mancanti = max(0, promo - coeff) if promo else None
        col3.metric("Mancanti alla promozione",
                    "Raggiunta! 🎉" if mancanti == 0 else f"{mancanti} pt" if mancanti is not None else "—")

        # ── Cap tornei ridotti ────────────────────────────────────────
        rid = risultato["rid_prog"]
        max_r = risultato["max_rid"]
        if max_r < 9999:
            pct = min(100, round(rid / max_r * 100))
            st.caption(f"Punti da tornei ridotti: {rid}/{max_r} pt ({pct}%)")
            st.progress(pct / 100)

        # ── Dettaglio vittorie ────────────────────────────────────────
        st.subheader("Dettaglio partite")

        vitt = risultato["vitt_sing"]
        n_tot = risultato["n_totali"]

        # Vittorie usate
        st.markdown("**Vittorie usate nel calcolo**")
        rows_used = []
        for v in vitt:
            if not v["used"]: continue
            rows_used.append({
                "Ord.": v["rank"],
                "#": v["n"],
                "Class.": v["cl"],
                "Esito": {"W":"Vittoria","WR":"Vittoria (ritiro avv.)","WA":"Vittoria per assenza avv."}.get(v["esito"], v["esito"]),
                "Formato": "Ridotto" if v["ridotto"] else "Intero",
                "Valore": f"{v['valore_bruto']} pt",
                "Pt usati": f"{v['punti_eff']} pt" + (" ⚠️ cap" if v["capped"] else ""),
            })
        if rows_used:
            st.dataframe(rows_used, use_container_width=True, hide_index=True)

        # Vittorie non usate
        unused = [v for v in vitt if not v["used"]]
        if unused:
            with st.expander(f"Vittorie non usate ({len(unused)})"):
                rows_unused = []
                for v in unused:
                    rows_unused.append({
                        "Ord.": v["rank"],
                        "#": v["n"],
                        "Class.": v["cl"],
                        "Valore": f"{v['valore_bruto']} pt",
                        "Formato": "Ridotto" if v["ridotto"] else "Intero",
                    })
                st.dataframe(rows_unused, use_container_width=True, hide_index=True)

        # Sconfitte
        sconfitte = [m for m in partite if m["esito"] in ("L","LR","LA")]
        if sconfitte:
            with st.expander(f"Sconfitte e assenze ({len(sconfitte)})"):
                g_curr = GRUPPI[risultato["classe"]]
                rows_s = []
                for m in sconfitte:
                    diff = GRUPPI[m["cl"]] - g_curr
                    rows_s.append({
                        "#": m["n"],
                        "Class.": m["cl"],
                        "Esito": {"L":"Sconfitta","LR":"Sconfitta (ritiro mio)","LA":"Assenza propria"}.get(m["esito"]),
                        "Formato": "Ridotto" if m["ridotto"] else "Intero",
                        "Nella formula": desc_rel(diff, vittoria=False),
                    })
                st.dataframe(rows_s, use_container_width=True, hide_index=True)

        # ── Bonus ────────────────────────────────────────────────────
        st.subheader("Bonus")
        if risultato["ha_bonus_assenza"]:
            st.success(f"✅ Bonus assenza sconfitte (art. 3.4a): +{risultato['bonus_assenza']} pt")
        else:
            motivo = "sconfitte vs pari/inf presenti" if risultato["sconfitte_pari_inf"] \
                     else f"solo {risultato['avv_pari_inf']}/5 avversari pari/inf richiesti"
            st.warning(f"Bonus assenza sconfitte: non assegnato — {motivo}")

        for t in risultato["tornei_detail"]:
            if t["motivo"]:
                st.warning(f"Torneo vinto #{t['n']}: non assegnato — {t['motivo']}")
            else:
                st.success(f"✅ Torneo vinto #{t['n']} (art. 3.4b): +{t['bonus']} pt")
        if not risultato["tornei_detail"]:
            st.info("Nessun torneo vinto indicato (art. 3.4b)")

        if bonus_camp > 0:
            st.success(f"✅ Bonus campionati individuali: +{bonus_camp} pt")

        # ── Download report HTML ──────────────────────────────────────
        st.divider()
        html_path = tempfile.mktemp(suffix=".html")
        genera_html(risultato, partite, html_path)
        with open(html_path, "r", encoding="utf-8") as f:
            html_content = f.read()
        os.unlink(html_path)

        nome_output = Path(uploaded.name).stem + "_risultato.html"
        st.download_button(
            label="⬇️  Scarica report HTML",
            data=html_content,
            file_name=nome_output,
            mime="text/html",
            use_container_width=True
        )

    except Exception as ex:
        st.error(f"Errore durante il calcolo: {ex}")
        import traceback
        st.code(traceback.format_exc())

    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass

elif not uploaded:
    st.info("Carica il file Excel per procedere con il calcolo.")
