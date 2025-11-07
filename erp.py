def build_report(trans_df: pd.DataFrame, gr_df: pd.DataFrame, en_df: pd.DataFrame) -> pd.DataFrame:
    """Compare Greek vs English OCR tokens, produce a clean translation mapping."""
    from rapidfuzz import fuzz

    def norm(s):
        s = str(s).strip()
        s = re.sub(r"[\u00A0\s]+", " ", s)
        return s.lower()

    rows = []
    for _, tr in trans_df.iterrows():
        greek_expected = str(tr["Greek"]).strip()
        eng_expected   = str(tr["English"]).strip()

        # --- Find actual OCR tokens ---
        gr_hits = gr_df[gr_df["text"].str.contains(greek_expected[:4], case=False, na=False)]
        en_hits = en_df[en_df["text"].str.contains(eng_expected[:4], case=False, na=False)]

        if gr_hits.empty:
            rows.append({
                "Greek (Expected)": greek_expected,
                "Greek (Detected)": "",
                "English (Expected)": eng_expected,
                "English (Detected)": "",
                "Status": 4,
                "Status Description": "Field Not Found on the Report View",
                "Confidence": 0.0,
                "Reason": "Greek header not visible in screenshots",
                "Screenshot": ""
            })
            continue

        # take first Greek detection
        g = gr_hits.iloc[0]
        gtxt, gx, gy, gconf, gsrc = g["text"], g["x"], g["y"], g["conf"], g.get("source","")

        # find best English match
        best_row, best_score = None, 0
        for _, e in en_df.iterrows():
            sim_txt = fuzz.token_set_ratio(gtxt, e["text"]) / 100
            dist = math.sqrt((gx - e["x"])**2 + (gy - e["y"])**2)
            pos_score = max(0, 1 - dist)
            total = 0.7*sim_txt + 0.3*pos_score
            if total > best_score:
                best_score = total
                best_row = e

        if best_row is None or best_score < 0.65:
            rows.append({
                "Greek (Expected)": greek_expected,
                "Greek (Detected)": gtxt,
                "English (Expected)": eng_expected,
                "English (Detected)": "",
                "Status": 3,
                "Status Description": "Field Not Translated",
                "Confidence": round(best_score,3),
                "Reason": "No strong English match detected",
                "Screenshot": gsrc
            })
            continue

        # found an English match
        e = best_row
        estatus, reason = 1, ""
        if best_score >= 0.9:
            estatus, reason = 1, "Good match"
        elif best_score >= 0.75:
            estatus, reason = 2, "Possible synonym / slightly different"
        else:
            estatus, reason = 0, "Low certainty match"

        rows.append({
            "Greek (Expected)": greek_expected,
            "Greek (Detected)": gtxt,
            "English (Expected)": eng_expected,
            "English (Detected)": e["text"],
            "Status": estatus,
            "Status Description": {
                1:"Translated_Correct",
                2:"Translated_Not Accurate",
                0:"Pending Review",
                3:"Field Not Translated",
                4:"Field Not Found on the Report View"
            }[estatus],
            "Confidence": round(best_score,3),
            "Reason": reason,
            "Screenshot": e.get("source","")
        })

    out = pd.DataFrame(rows)

    # sort — problems first
    order = pd.CategoricalDtype(
        ["Field Not Found on the Report View",
         "Field Not Translated",
         "Translated_Not Accurate",
         "Pending Review",
         "Translated_Correct"], ordered=True)
    out["Status Description"] = out["Status Description"].astype(order)
    out = out.sort_values(["Status Description","Greek (Expected)"]).reset_index(drop=True)
    return out
