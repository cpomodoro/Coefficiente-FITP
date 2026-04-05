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
