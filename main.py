"""논문 PDF → PPT 자동 생성 파이프라인."""

import os
import re
import sys
import json
import glob
import fitz  # PyMuPDF
import anthropic
from ppt_builder import build_presentation

# ── 설정 ──────────────────────────────────────────────────
TARGET_DATE = "260329"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "paper_review-format.pptx")
WORKSPACE_DIR = os.path.join(BASE_DIR, "workspaces", TARGET_DATE)
ASSETS_DIR = os.path.join(WORKSPACE_DIR, "assets")
OUTPUT_DIR = os.path.join(WORKSPACE_DIR, "output")
OUTPUT_PATH = os.path.join(OUTPUT_DIR, f"{TARGET_DATE}_presentation.pptx")
STRUCTURE_PATH = os.path.join(OUTPUT_DIR, "slide_structure.json")

os.makedirs(OUTPUT_DIR, exist_ok=True)


# ── 1. PDF 텍스트 추출 ───────────────────────────────────
def extract_text(pdf_path: str) -> tuple[str, list[str]]:
    """PDF에서 전체 텍스트와 figure/table 캡션을 추출한다."""
    doc = fitz.open(pdf_path)
    pages = []
    for i, page in enumerate(doc):
        pages.append(f"--- PAGE {i+1} ---\n{page.get_text()}")
    full_text = "\n\n".join(pages)

    # 캡션 추출
    caption_pattern = r"((?:Figure|Fig\.|Table)\s+\d+[.:]\s*.+?)(?:\n|$)"
    captions = re.findall(caption_pattern, full_text, re.IGNORECASE)

    return full_text, captions


# ── 2. 섹션 파싱 ─────────────────────────────────────────
def parse_sections(full_text: str) -> dict:
    """텍스트를 주요 섹션별로 분리한다."""
    section_headers = [
        "abstract", "introduction", "related work", "background",
        "method", "approach", "model", "architecture", "proposed",
        "experiment", "result", "evaluation",
        "ablation", "analysis",
        "conclusion", "discussion", "limitation", "future work",
    ]

    pattern = r"(?:^|\n)(\d+\.?\s*)?(" + "|".join(section_headers) + r")s?\b"
    matches = list(re.finditer(pattern, full_text, re.IGNORECASE))

    if not matches:
        return {"full_text": full_text}

    sections = {}
    for i, match in enumerate(matches):
        name = match.group(2).lower().strip()
        start = match.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(full_text)
        content = full_text[start:end].strip()
        # 같은 이름의 섹션이 있으면 합치기
        if name in sections:
            sections[name] += "\n" + content
        else:
            sections[name] = content

    return sections


# ── 3. Assets 스캔 ────────────────────────────────────────
def scan_assets(assets_dir: str) -> list[str]:
    """assets 디렉토리의 이미지 파일 목록을 반환한다."""
    if not os.path.exists(assets_dir):
        return []
    exts = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg")
    files = [f for f in os.listdir(assets_dir)
             if f.lower().endswith(exts)]
    return sorted(files)


# ── 4. LLM으로 슬라이드 구조 생성 ────────────────────────
SYSTEM_PROMPT = """You are an academic presentation assistant. You read research papers and produce structured JSON for PowerPoint slides. Output ONLY valid JSON, no markdown fences, no commentary."""

USER_PROMPT_TEMPLATE = """다음 논문을 바탕으로 발표 슬라이드 구조를 JSON으로 생성해주세요.

## 사용 가능한 이미지 파일
{asset_list}

## 논문에서 발견된 Figure/Table 캡션
{captions_text}

## 논문 내용
{paper_text}

## 슬라이드 구성 규칙
1. 다음 순서로 슬라이드를 생성하세요:
   - 슬라이드 1: 제목 (논문 제목, 저자, 학회/저널)
   - 슬라이드 2: 연구 배경 (기존 연구의 문제점, motivation)
   - 슬라이드 3: 전체 아키텍처/방법론 요약
   - 슬라이드 4-6: 아키텍처/방법론 세부 설명 (필요한 만큼 2-4장)
   - 슬라이드 7-9: 실험 결과 (필요한 만큼 2-4장)
   - Ablation study 슬라이드 (논문에 ablation이 있는 경우에만 포함)
   - 결론 슬라이드
   - 마무리 인사 슬라이드

2. 각 content 슬라이드:
   - section: 섹션 라벨 (예: "Background", "Architecture", "Experiments")
   - title: 슬라이드 제목
   - bullets: 핵심 내용 3-5개 (한국어, 간결하게)
   - assets: 해당 슬라이드에 넣을 이미지 파일명 리스트 (캡션 매칭)
   - speaker_notes: 발표자 참고 노트 (한국어, 선택)

3. 이미지 매칭: "fig1.png"은 "Figure 1", "table2.png"은 "Table 2"에 매칭합니다.
   사용 가능한 파일에만 매칭하세요. 각 파일은 한 슬라이드에서만 사용하세요.

4. bullets는 한국어로 작성해주세요. 핵심만 간결하게.

## 출력 JSON 스키마
{{
  "slides": [
    {{
      "slide_type": "title",
      "title": "논문 제목",
      "subtitle": "저자1, 저자2, ...",
      "venue": "학회명",
      "date": "연도"
    }},
    {{
      "slide_type": "content",
      "section": "섹션명",
      "title": "슬라이드 제목",
      "bullets": ["포인트1", "포인트2", "포인트3"],
      "assets": ["fig1.png"],
      "speaker_notes": "발표자 노트"
    }},
    {{
      "slide_type": "closing",
      "title": "감사합니다",
      "subtitle": "Q&A"
    }}
  ]
}}
"""


def generate_slide_structure(sections: dict, captions: list[str],
                             asset_files: list[str], full_text: str) -> dict:
    """Claude API를 호출하여 슬라이드 구조 JSON을 생성한다."""
    # 텍스트가 너무 길면 섹션별 요약 사용, 아니면 전체 텍스트 사용
    if len(full_text) > 80000:
        paper_text = "\n\n".join(
            f"### {name}\n{content[:3000]}" for name, content in sections.items()
        )
    else:
        paper_text = full_text

    asset_list = "\n".join(f"- {f}" for f in asset_files) if asset_files else "(없음)"
    captions_text = "\n".join(f"- {c}" for c in captions) if captions else "(캡션 없음)"

    user_prompt = USER_PROMPT_TEMPLATE.format(
        asset_list=asset_list,
        captions_text=captions_text,
        paper_text=paper_text,
    )

    client = anthropic.Anthropic()
    message = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[
            {"role": "user", "content": user_prompt}
        ],
        system=SYSTEM_PROMPT,
    )

    response_text = message.content[0].text.strip()

    # JSON 파싱 (혹시 마크다운 펜스가 있으면 제거)
    if response_text.startswith("```"):
        response_text = re.sub(r"^```(?:json)?\n?", "", response_text)
        response_text = re.sub(r"\n?```$", "", response_text)

    slide_data = json.loads(response_text)
    return slide_data


# ── 메인 파이프라인 ───────────────────────────────────────
def main():
    # PDF 찾기
    pdf_files = glob.glob(os.path.join(WORKSPACE_DIR, "*.pdf"))
    if not pdf_files:
        print(f"❌ PDF 파일을 찾을 수 없습니다: {WORKSPACE_DIR}")
        sys.exit(1)

    pdf_path = pdf_files[0]
    print(f"📄 PDF: {os.path.basename(pdf_path)}")

    # 기존 슬라이드 구조가 있으면 재사용할지 확인
    if os.path.exists(STRUCTURE_PATH):
        answer = input(f"기존 슬라이드 구조 파일이 있습니다. 재사용할까요? (y/n): ").strip().lower()
        if answer == "y":
            with open(STRUCTURE_PATH, "r", encoding="utf-8") as f:
                slide_data = json.load(f)
            print("♻️  기존 슬라이드 구조를 재사용합니다.")
            prs = build_presentation(slide_data, TEMPLATE_PATH, ASSETS_DIR)
            prs.save(OUTPUT_PATH)
            print(f"✅ PPT 생성 완료: {OUTPUT_PATH}")
            return

    # 1) 텍스트 추출
    print("📖 텍스트 추출 중...")
    full_text, captions = extract_text(pdf_path)
    print(f"   캡션 {len(captions)}개 발견")

    # 2) 섹션 파싱
    sections = parse_sections(full_text)
    print(f"   섹션 {len(sections)}개 파싱됨: {', '.join(sections.keys())}")

    # 3) Assets 스캔
    asset_files = scan_assets(ASSETS_DIR)
    print(f"🖼️  Assets: {asset_files if asset_files else '(없음)'}")

    # 4) LLM으로 슬라이드 구조 생성
    print("🤖 슬라이드 구조 생성 중 (Claude API 호출)...")
    slide_data = generate_slide_structure(sections, captions, asset_files, full_text)

    # 중간 결과 저장 (디버깅 및 재사용)
    with open(STRUCTURE_PATH, "w", encoding="utf-8") as f:
        json.dump(slide_data, f, ensure_ascii=False, indent=2)
    print(f"💾 슬라이드 구조 저장: {STRUCTURE_PATH}")

    # 5) PPT 생성
    print("📊 PPT 생성 중...")
    prs = build_presentation(slide_data, TEMPLATE_PATH, ASSETS_DIR)
    prs.save(OUTPUT_PATH)
    print(f"✅ PPT 생성 완료: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
