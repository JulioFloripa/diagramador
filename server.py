"""
API HTTP para o pipeline de diagramação de provas — Colégio Fleming

Endpoints:
    POST /process   — recebe .docx, retorna .docx formatado
    GET  /health    — health check

Uso:
    pip install -r requirements.txt
    python server.py [--port 5050] [--host 0.0.0.0]

Variáveis de ambiente:
    ANTHROPIC_API_KEY  — chave da API Anthropic (obrigatória para revisão AI)
    DIAGRAMADOR_PORT   — porta (default 5050)
"""

import io
import json
import os
import sys
import tempfile
import traceback
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

from extractor import extract_questions
from generator import generate_exam

try:
    from agents import review_extraction
    HAS_AGENTS = True
except ImportError:
    HAS_AGENTS = False


class DiagramadorHandler(BaseHTTPRequestHandler):

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/health":
            self._json_response(200, {
                "status": "ok",
                "agents_available": HAS_AGENTS,
                "has_api_key": bool(os.environ.get("ANTHROPIC_API_KEY")),
            })
        else:
            self._json_response(404, {"error": "Not found"})

    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/process":
            self._handle_process()
        else:
            self._json_response(404, {"error": "Not found"})

    def _handle_process(self):
        content_length = int(self.headers.get("Content-Length", 0))
        if content_length == 0:
            self._json_response(400, {"error": "No file uploaded"})
            return

        body = self.rfile.read(content_length)
        params = parse_qs(urlparse(self.path).query)
        skip_review = params.get("skip_review", ["false"])[0].lower() == "true"

        try:
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_in:
                tmp_in.write(body)
                tmp_in_path = tmp_in.name

            extraction = extract_questions(tmp_in_path)

            if not skip_review and HAS_AGENTS and os.environ.get("ANTHROPIC_API_KEY"):
                result = review_extraction(extraction, auto_apply=True)
                if "corrected" in result:
                    extraction = result["corrected"]
                review_info = {
                    "review": result.get("review", {}),
                    "validation": result.get("validation", {}),
                }
            else:
                review_info = {"skipped": True}

            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_out:
                tmp_out_path = tmp_out.name

            generate_exam(extraction, tmp_out_path)

            with open(tmp_out_path, "rb") as f:
                docx_bytes = f.read()

            os.unlink(tmp_in_path)
            os.unlink(tmp_out_path)

            professor = extraction.get("professor", "Professor")
            disciplina = extraction.get("disciplina", "")
            serie = extraction.get("serie", "")
            n_questions = len(extraction.get("questions", []))

            filename = f"PROVA_{disciplina}_{serie}.docx".replace(" ", "_")

            self.send_response(200)
            self.send_header("Content-Type",
                             "application/vnd.openxmlformats-officedocument"
                             ".wordprocessingml.document")
            self.send_header("Content-Disposition",
                             f'attachment; filename="{filename}"')
            self.send_header("X-Professor", professor)
            self.send_header("X-Disciplina", disciplina)
            self.send_header("X-Serie", serie)
            self.send_header("X-Questions", str(n_questions))
            self.send_header("X-Review", json.dumps(review_info,
                                                     ensure_ascii=False)[:2000])
            self.send_header("Content-Length", str(len(docx_bytes)))
            self.end_headers()
            self.wfile.write(docx_bytes)

        except Exception as e:
            traceback.print_exc()
            self._json_response(500, {
                "error": str(e),
                "type": type(e).__name__,
            })

    def _json_response(self, status, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        print(f"[{self.log_date_time_string()}] {format % args}")


def main():
    import argparse
    parser = argparse.ArgumentParser(description="API do Diagramador Fleming")
    parser.add_argument("--port", type=int,
                        default=int(os.environ.get("DIAGRAMADOR_PORT", "5050")))
    parser.add_argument("--host", default="0.0.0.0")
    args = parser.parse_args()

    server = HTTPServer((args.host, args.port), DiagramadorHandler)
    print(f"Diagramador API rodando em http://{args.host}:{args.port}")
    print(f"  Agentes AI: {'disponíveis' if HAS_AGENTS else 'indisponíveis'}")
    print(f"  API key: {'configurada' if os.environ.get('ANTHROPIC_API_KEY') else 'NÃO configurada'}")
    print(f"\nEndpoints:")
    print(f"  GET  /health   — status do serviço")
    print(f"  POST /process  — envia .docx, recebe .docx formatado")
    print(f"       ?skip_review=true para pular revisão AI")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nEncerrando...")
        server.shutdown()


if __name__ == "__main__":
    main()
