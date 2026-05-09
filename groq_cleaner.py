"""
groq_cleaner.py
───────────────
Sends reference/textbook entries to Groq's free API (Llama 3.3 70B)
to fix common formatting issues like missing spaces in edition numbers.

Groq Free Tier: No credit card needed.  Sign up at console.groq.com
"""

import re
import requests


def clean_references_with_groq(references: list, api_key: str) -> list:
    """
    Fix formatting issues in academic reference entries using Groq.

    Issues fixed:
      - "7thedition"  →  "7th edition"
      - "2ndedition"  →  "2nd edition"
      - "3rdedition"  →  "3rd edition"
      - "10thedition" →  "10th edition"
      - Missing spaces after commas in year/edition fields

    Returns the cleaned list in the same order.
    On any error, silently returns the original list.
    """
    if not references or not api_key:
        return references

    # Build a numbered list to send to the model
    numbered = "\n".join(f"{i + 1}. {ref}" for i, ref in enumerate(references))

    prompt = f"""You are a precise formatter for academic book references.
Fix ONLY these issues in each reference:
1. Missing spaces in edition numbers  e.g.  "7thedition" → "7th edition"
                                             "2ndedition" → "2nd edition"
                                             "3rdedition" → "3rd edition"
                                             "10thedition" → "10th edition"
2. Missing space after a comma where the next char is a letter (not a space)

Keep ALL other text EXACTLY as given — same author names, same titles,
same publisher names, same years.

Return ONLY the fixed references with the same numbering.
No extra commentary, no blank lines between items.

References:
{numbered}"""

    try:
        response = requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            json={
                "model": "llama-3.3-70b-versatile",
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": 1024,
                "temperature": 0,
            },
            timeout=30,
        )

        if response.status_code != 200:
            print(f"[Groq] API error {response.status_code}: {response.text}")
            return references

        content = response.json()["choices"][0]["message"]["content"].strip()

        # Parse the numbered lines back into a plain list
        cleaned = []
        for line in content.splitlines():
            line = line.strip()
            if not line:
                continue
            # Strip leading "1. " or "1) " prefix
            match = re.match(r"^\d+[\.\)]\s+", line)
            if match:
                line = line[match.end():]
            cleaned.append(line)

        # Safety: if count changed, return originals to avoid misalignment
        if len(cleaned) != len(references):
            print(
                f"[Groq] Count mismatch (got {len(cleaned)}, "
                f"expected {len(references)}). Using originals."
            )
            return references

        return cleaned

    except Exception as exc:
        print(f"[Groq] Exception during cleaning: {exc}")
        return references
