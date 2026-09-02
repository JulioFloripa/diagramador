"""
Agentes AI para revisão de provas — Colégio Fleming

Pipeline de dois agentes:
  1. Revisor (Haiku 4.5 — barato): analisa o JSON do extrator, detecta
     erros de texto, estrutura e formatação, propõe correções.
  2. Validador (Sonnet 5 — preciso): valida as correções, garante
     qualidade final antes da geração do .docx.

Uso standalone:
    python agents.py input.json -o reviewed.json [--api-key KEY]

Uso como módulo (para n8n Execute Command):
    from agents import review_extraction
    result = review_extraction(extraction_data)
"""

import argparse
import json
import os
import sys
from typing import Optional

try:
    import anthropic
except ImportError:
    anthropic = None

# ---------------------------------------------------------------------------
# Model configuration
# ---------------------------------------------------------------------------

MODEL_REVIEWER = "claude-haiku-4-5"     # $1/$5 per MTok — fast, cheap
MODEL_VALIDATOR = "claude-sonnet-5"     # $2/$10 per MTok — precise

# ---------------------------------------------------------------------------
# System prompts
# ---------------------------------------------------------------------------

SYSTEM_REVIEWER = """Você é um revisor especializado em provas escolares do Colégio Fleming (Florianópolis).

Você recebe um JSON estruturado extraído automaticamente de um arquivo .docx de prova.
Sua tarefa é analisar cada questão e reportar problemas encontrados.

## O que verificar

1. **Texto corrompido**: palavras coladas ("inUm" deveria ser "in. Um"), caracteres
   estranhos, trechos truncados ou duplicados que parecem erro de copy-paste.

2. **Estrutura das questões**:
   - Cada questão deve ter entre 4 e 5 alternativas (a-e).
   - Alternativas devem ter conteúdo (não vazias).
   - Numeração sequencial (1, 2, 3...).

3. **Imagens**: se uma questão menciona "figura", "imagem", "ilustração" ou
   "conforme mostrado" no enunciado, deve haver pelo menos um fragment de
   tipo "image" no statement.

4. **Equações**: se há fragments do tipo "equation" com conteúdo vazio ou
   muito curto (< 2 caracteres), pode indicar equação perdida.

5. **Alternativas duplicadas**: duas alternativas com texto idêntico.

6. **Metadados**: professor, disciplina e série devem estar preenchidos.

## Formato de saída

Responda APENAS com um JSON válido (sem markdown, sem ```), com esta estrutura:

{
  "status": "ok" | "issues_found",
  "issues": [
    {
      "question": 3,
      "type": "corrupted_text" | "missing_image" | "empty_alternative" |
              "duplicate_alternative" | "wrong_alt_count" | "missing_equation" |
              "numbering_gap" | "metadata_missing" | "other",
      "description": "Descrição clara do problema",
      "suggestion": "Correção sugerida (quando possível)",
      "severity": "error" | "warning" | "info",
      "auto_fixable": true | false
    }
  ],
  "corrections": [
    {
      "question": 3,
      "field": "statement",
      "fragment_index": 4,
      "original": "inUm",
      "corrected": "in. Um"
    }
  ],
  "summary": "Resumo em 1-2 frases do estado geral da prova"
}

Se não houver problemas, retorne {"status": "ok", "issues": [], "corrections": [], "summary": "..."}.
"""

SYSTEM_VALIDATOR = """Você é um validador de qualidade de provas escolares do Colégio Fleming.

Você recebe:
1. O JSON original da extração de uma prova.
2. O relatório do revisor com issues e correções propostas.

Sua tarefa é validar as correções propostas e decidir quais aplicar.

## Regras

1. **Aprovar** correções que claramente melhoram o texto (palavras coladas,
   caracteres estranhos evidentes).

2. **Rejeitar** correções que alteram o conteúdo acadêmico — nunca "corrigir"
   valores numéricos, fórmulas, ou o conteúdo das alternativas. Erros
   intencionais do professor (como distratores nas alternativas) NÃO devem
   ser corrigidos.

3. **Alertar** sobre problemas estruturais (imagem faltando, alternativa
   vazia) sem tentar corrigi-los — esses precisam de intervenção do
   professor.

4. Marque como "needs_human" qualquer correção que você não tem certeza.

## Formato de saída

Responda APENAS com um JSON válido (sem markdown, sem ```):

{
  "approved_corrections": [
    {
      "question": 3,
      "field": "statement",
      "fragment_index": 4,
      "original": "inUm",
      "corrected": "in. Um",
      "confidence": 0.95
    }
  ],
  "rejected_corrections": [
    {
      "question": 5,
      "reason": "Alteraria o conteúdo acadêmico"
    }
  ],
  "human_review_needed": [
    {
      "question": 8,
      "issue": "Imagem mencionada no enunciado mas não encontrada",
      "action_needed": "Professor precisa reenviar o arquivo com a imagem"
    }
  ],
  "final_status": "approved" | "needs_review" | "rejected",
  "summary": "Resumo final em 1-2 frases"
}
"""


# ---------------------------------------------------------------------------
# Prepare extraction data for the agents (strip base64 images for token savings)
# ---------------------------------------------------------------------------

def _prepare_for_agent(data: dict) -> dict:
    """Strip base64 image data to save tokens — agents only need structure."""
    slim = dict(data)
    if "images" in slim:
        slim["images"] = {k: f"<base64 {len(v)} chars>" for k, v in slim["images"].items()}
    return slim


def _extract_text_preview(question: dict, max_chars=200) -> str:
    """Get a readable text preview of a question."""
    text_parts = []
    for f in question.get("statement", []):
        if f.get("type") == "text":
            text_parts.append(f.get("content", ""))
    return "".join(text_parts)[:max_chars]


# ---------------------------------------------------------------------------
# Agent calls
# ---------------------------------------------------------------------------

def call_reviewer(client, data: dict) -> dict:
    """Run the reviewer agent on extraction data."""
    slim_data = _prepare_for_agent(data)
    user_message = json.dumps(slim_data, ensure_ascii=False, indent=2)

    response = client.messages.create(
        model=MODEL_REVIEWER,
        max_tokens=4096,
        system=SYSTEM_REVIEWER,
        messages=[{"role": "user", "content": user_message}],
    )

    response_text = ""
    for block in response.content:
        if block.type == "text":
            response_text += block.text

    try:
        return json.loads(response_text)
    except json.JSONDecodeError:
        cleaned = response_text.strip()
        if cleaned.startswith("```"):
            cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned[3:]
            if cleaned.endswith("```"):
                cleaned = cleaned[:-3]
            return json.loads(cleaned.strip())
        raise ValueError(f"Reviewer returned invalid JSON: {response_text[:500]}")


def call_validator(client, data: dict, review: dict) -> dict:
    """Run the validator agent on extraction data + reviewer report."""
    slim_data = _prepare_for_agent(data)

    user_message = json.dumps({
        "extraction": slim_data,
        "reviewer_report": review,
    }, ensure_ascii=False, indent=2)

    response = client.messages.create(
        model=MODEL_VALIDATOR,
        max_tokens=4096,
        system=SYSTEM_VALIDATOR,
        messages=[{"role": "user", "content": user_message}],
    )

    response_text = ""
    for block in response.content:
        if block.type == "text":
            response_text += block.text

    try:
        return json.loads(response_text)
    except json.JSONDecodeError:
        cleaned = response_text.strip()
        if cleaned.startswith("```"):
            cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned[3:]
            if cleaned.endswith("```"):
                cleaned = cleaned[:-3]
            return json.loads(cleaned.strip())
        raise ValueError(f"Validator returned invalid JSON: {response_text[:500]}")


# ---------------------------------------------------------------------------
# Apply approved corrections
# ---------------------------------------------------------------------------

def apply_corrections(data: dict, validation: dict) -> dict:
    """Apply approved corrections to the extraction data."""
    import copy
    result = copy.deepcopy(data)

    for correction in validation.get("approved_corrections", []):
        q_num = correction.get("question")
        field = correction.get("field", "statement")
        frag_idx = correction.get("fragment_index")
        corrected = correction.get("corrected")

        if q_num is None or corrected is None:
            continue

        for q in result["questions"]:
            if q["number"] == q_num:
                fragments = q.get(field, [])
                if frag_idx is not None and 0 <= frag_idx < len(fragments):
                    fragments[frag_idx]["content"] = corrected
                break

    return result


# ---------------------------------------------------------------------------
# Full pipeline
# ---------------------------------------------------------------------------

def review_extraction(data: dict, api_key: Optional[str] = None,
                      auto_apply: bool = False) -> dict:
    """
    Run the full review pipeline on extraction data.

    Returns:
        {
            "original": <original data>,
            "review": <reviewer output>,
            "validation": <validator output>,
            "corrected": <data with approved corrections applied> (if auto_apply),
            "usage": {"reviewer": {...}, "validator": {...}}
        }
    """
    if anthropic is None:
        raise ImportError(
            "anthropic package not installed. Run: pip install anthropic"
        )

    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()

    print(f"[Revisor] Analisando {len(data.get('questions', []))} questões "
          f"com {MODEL_REVIEWER}...")
    review = call_reviewer(client, data)
    n_issues = len(review.get("issues", []))
    n_corrections = len(review.get("corrections", []))
    print(f"[Revisor] {review.get('status', '?')}: "
          f"{n_issues} problemas, {n_corrections} correções propostas")

    if review.get("summary"):
        print(f"[Revisor] {review['summary']}")

    print(f"\n[Validador] Validando correções com {MODEL_VALIDATOR}...")
    validation = call_validator(client, data, review)
    n_approved = len(validation.get("approved_corrections", []))
    n_rejected = len(validation.get("rejected_corrections", []))
    n_human = len(validation.get("human_review_needed", []))
    print(f"[Validador] {validation.get('final_status', '?')}: "
          f"{n_approved} aprovadas, {n_rejected} rejeitadas, "
          f"{n_human} para revisão humana")

    if validation.get("summary"):
        print(f"[Validador] {validation['summary']}")

    result = {
        "review": review,
        "validation": validation,
    }

    if auto_apply and n_approved > 0:
        corrected = apply_corrections(data, validation)
        result["corrected"] = corrected
        print(f"\n[Pipeline] {n_approved} correções aplicadas automaticamente")
    elif n_approved > 0:
        print(f"\n[Pipeline] {n_approved} correções prontas para aplicar "
              f"(use --apply para aplicar)")

    if n_human > 0:
        print(f"\n⚠ {n_human} item(ns) precisam de revisão humana:")
        for item in validation.get("human_review_needed", []):
            print(f"  Q{item.get('question', '?')}: {item.get('issue', '')}")

    return result


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Revisa extração de prova com agentes AI"
    )
    parser.add_argument("input", help="JSON de entrada (do extractor.py)")
    parser.add_argument("-o", "--output", help="JSON de saída com relatório")
    parser.add_argument("--apply", action="store_true",
                       help="Aplica correções aprovadas automaticamente")
    parser.add_argument("--api-key", help="Anthropic API key "
                       "(ou use ANTHROPIC_API_KEY env var)")
    parser.add_argument("--reviewer-only", action="store_true",
                       help="Roda apenas o revisor (sem validador)")

    args = parser.parse_args()

    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    if anthropic is None:
        print("ERRO: anthropic package não instalado.")
        print("Execute: pip install anthropic")
        sys.exit(1)

    api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERRO: API key não fornecida.")
        print("Use --api-key KEY ou defina ANTHROPIC_API_KEY")
        sys.exit(1)

    result = review_extraction(data, api_key=api_key, auto_apply=args.apply)

    if args.output:
        output_data = result
        if args.apply and "corrected" in result:
            output_data = result["corrected"]
            output_data["_review"] = result["review"]
            output_data["_validation"] = result["validation"]

        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        print(f"\nResultado salvo em: {args.output}")


if __name__ == "__main__":
    main()
