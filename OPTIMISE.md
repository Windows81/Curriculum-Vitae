Paste everything which resides in a code block into your LLM of choice. This will optimise my résumé for exact keyword matches by strategically adding `<choice>` tags in the text.

```
Template:
Add <choice> tags around these exact words if they appear:
{team → <choice>team|group</choice>}
{data → <choice>data|information</choice>}
{communication → <choice>communication|interaction</choice>}
{NumPy → <choice>NumPy|Python</choice>}

Résumé:
"
{INSERT SECTION}
"

Job Description:
"
{INSERT JOB DESCRIPTION}
"

Rules:
1. Only wrap EXACT matches (case-insensitive).
2. Do not add any other tags.
3. Return only the modified text.
4. If any placeholder fields are present, cease outputing and request that info from user.
```
