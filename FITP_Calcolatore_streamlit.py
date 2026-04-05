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
