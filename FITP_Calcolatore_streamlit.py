import streamlit as st
import io
import contextlib
import tempfile
import os
from pathlib import Path

from fitp_calcolo import (
    CLASSI, leggi_excel, calcola_con_promozioni,
    stampa_risultati, genera_html
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

# ── Caricamento file ─────────────────────────────────────────────────
st.subheader("File partite")
uploaded = st.file_uploader(
    "Carica il file Excel con le partite",
    type=["xlsx", "xls"],
    help="Il file deve contenere le colonne: Classifica, Esito, Tipo, Punteggio, Torneo Veterani, Vittoria Torneo"
)

# ── Calcola ──────────────────────────────────────────────────────────
if uploaded and st.button("▶  Calcola", type="primary", use_container_width=True):

    # Salva il file caricato in un file temporaneo
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    try:
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
                from fitp_calcolo import GRUPPI, desc_rel
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
