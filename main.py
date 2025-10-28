@app.post("/analyze")
async def analyze(request: Request):
    try:
        try:
            data = await request.json()
        except Exception:
            body = await request.body()
            return JSONResponse({"keyword": "OTHER", "error": f"Invalid JSON: {body[:100].decode(errors='ignore')}"})

        filename = data.get("filename", "unknown")
        content_b64 = data.get("content", "")
        if not content_b64:
            return JSONResponse({"keyword": "OTHER", "error": "Empty content"})

        file_bytes = base64.b64decode(content_b64)
        raw_text = extract_text_from_file(file_bytes, filename)
        text = normalize(raw_text)

        # --------------------------------------------------------------
        # DEBUG (keep it – helps you see why a file fails)
        # --------------------------------------------------------------
        print(f"\n[DEBUG] File: {filename}")
        print(f"[DEBUG] Raw len: {len(raw_text)} | Norm len: {len(text)}")
        print(f"[DEBUG] Sample: {text[:300]}")
        print(f"[DEBUG] IKOS?: {'IKOS' in text} | PORTO PETRO?: {'PORTO PETRO' in text}")

        # --------------------------------------------------------------
        # 1. FILENAME FAST-PATH
        # --------------------------------------------------------------
        filename_norm = normalize(filename)
        if any(k in filename_norm for k in ["ANDALUSIA", "ODISIA", "ESTEPONA", "COSTA DEL SOL"]):
            return JSONResponse({"keyword": "ANDALUSIA", "error": ""})
        if "PORTO PETRO" in filename_norm:
            return JSONResponse({"keyword": "PORTO PETRO", "error": ""})

        # --------------------------------------------------------------
        # 2. LOCAL KEYWORD DETECTION
        # --------------------------------------------------------------
        local_result = detect_ikos_hotel(text)
        if local_result:
            return JSONResponse({"keyword": local_result, "error": ""})

        # --------------------------------------------------------------
        # 3. LLM FALLBACK (only if we really have no clue)
        # --------------------------------------------------------------
        if len(text) > 50:
            prompt = f"""
You are a JSON classifier. Return **only** valid JSON.

Text (first 6000 chars):
{text[:6000]}

Rules (exact phrase match):
- "ANDALUSIA"  → contains  any ANDALUSIA or andalusia
- "PORTO PETRO" → contains the exact phrase "PORTO PETRO"
- "IKOS SPANISH HOTEL MANAGEMENT" → contains IKOS SPANISH HOTEL MANAGEMENT or ikos spanish hotel management or ISHM
- otherwise → "OTHER"

Return ONLY:
{{"keyword": "ANDALUSIA" | "PORTO PETRO" | "IKOS SPANISH HOTEL MANAGEMENT" | "OTHER"}}
"""
            try:
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": prompt}],
                    temperature=0,
                    response_format={"type": "json_object"}
                )
                raw = response.choices[0].message.content.strip()
                parsed = json.loads(raw)
                keyword = parsed.get("keyword", "OTHER").upper()
                valid = ["ANDALUSIA", "PORTO PETRO", "IKOS SPANISH HOTEL MANAGEMENT", "OTHER"]
                if keyword not in valid:
                    keyword = "OTHER"
                return JSONResponse({"keyword": keyword, "error": ""})
            except Exception as e:
                return JSONResponse({"keyword": "OTHER", "error": f"LLM failed: {str(e)[:100]}"})

        return JSONResponse({"keyword": "OTHER", "error": "Empty PDF - no text extracted"})
    except Exception as e:
        return JSONResponse({"keyword": "OTHER", "error": f"Server error: {str(e)[:100]}"})
