"""
Writer — Claude API를 통한 content_map 생성 (v2: 정밀 매핑)
"""
import os
import json
import httpx

OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "")
MODEL = "anthropic/claude-sonnet-4-6"


def build_structured_prompt(analysis: dict) -> str:
    """분석 결과를 shape ID별 원본 텍스트가 명확한 구조적 프롬프트로 변환"""
    lines = []
    for slide in analysis["slides"]:
        lines.append(f"\n## 슬라이드 {slide['number']}")

        if slide["shapes"]:
            lines.append("### 도형 (shapes)")
            lines.append("| shape_id | 이름 | 원본 텍스트 (줄별) | 글자수 |")
            lines.append("|----------|------|---------------------|--------|")
            for sh in slide["shapes"]:
                for i, txt in enumerate(sh["text"]):
                    char_count = len(txt)
                    escaped = txt.replace("|", "\\|")[:60]
                    if i == 0:
                        lines.append(f"| {sh['id']} | {sh['name'][:20]} | {escaped} | {char_count} |")
                    else:
                        lines.append(f"| (계속) | | {escaped} | {char_count} |")

        if slide["tables"]:
            lines.append("### 표 (tables)")
            for tb in slide["tables"]:
                lines.append(f"표 이름: {tb['name']} ({tb['row_count']}행 x {tb['col_count']}열)")
                lines.append(f"| 행 | 셀 | 원본 텍스트 | 글자수(상한) |")
                lines.append(f"|----|----|------------|-------------|")
                for r_i, row in enumerate(tb["rows"]):
                    role = "헤더" if r_i == 0 else f"행{r_i+1}"
                    for c_i, cell in enumerate(row):
                        lines.append(f"| {role} | 셀{c_i+1} | {cell[:40]} | {len(cell)} |")

        if slide["images"]:
            lines.append(f"[이미지 {slide['images']}개 — 교체 대상 아님]")

    return "\n".join(lines)


async def generate_content_map(template_analysis: dict, markitdown_text: str,
                                topic: str, extra_info: str = "") -> dict:
    """정밀 content_map 생성"""
    structured = build_structured_prompt(template_analysis)
    total_shapes = sum(len(s["shapes"]) for s in template_analysis["slides"])

    system_prompt = """당신은 PPTX 템플릿 자동 채우기 전문가입니다.
주어진 템플릿의 각 shape_id별 원본 텍스트를 분석하고, 새 주제에 맞는 교체 텍스트를 생성합니다.

## 핵심 규칙
1. **shape_id 정확 매칭** — 분석표의 shape_id를 그대로 사용
2. **글자수 엄격 제한** — 반드시 원본 글자수 이하로 작성. 분석표의 "글자수" 열이 상한선임. 초과 시 텍스트 박스가 넘침. 짧게 쓰는 것은 괜찮지만 길게 쓰는 것은 절대 금지
3. **줄 수 유지** — 원본이 2줄이면 교체도 2줄 (배열 길이 동일)
4. **슬라이드/행/열 수 고정** — 절대 추가·삭제 금지
5. **이미지 제외** — 이미지 관련 도형은 content_map에 포함하지 않음
6. **표 셀 정밀 매핑** — 헤더 포함 모든 행의 모든 셀을 2차원 배열로 반환. 각 셀도 원본 글자수 이하로
7. **전문적 톤** — 비즈니스 실무에 적합한 전문 용어 사용
8. **핵심만 간결하게** — 글자수 제한 내에서 핵심 키워드 중심으로 압축 표현

## 응답 형식 (JSON만, 다른 텍스트 없이)
```json
{
  "슬라이드번호(문자열)": {
    "shapes": {
      "shape_id": ["줄1", "줄2"],
      "shape_id": "단일줄 텍스트"
    },
    "tables": {
      "표이름": [
        ["행1셀1", "행1셀2", "행1셀3"],
        ["행2셀1", "행2셀2", "행2셀3"]
      ]
    }
  }
}
```

주의: shape_id가 숫자여도 문자열로 표기. 교체 불필요한 shape는 생략 가능."""

    # 도형 0개면 markitdown 텍스트를 주 데이터소스로 사용
    if total_shapes == 0:
        user_prompt = f"""## 템플릿 텍스트 (markitdown 추출 — 도형 구조 미감지)
이 템플릿은 도형 ID를 감지하지 못했습니다. markitdown으로 추출한 텍스트를 기반으로
각 슬라이드의 텍스트를 새 주제에 맞게 교체해 주세요.
슬라이드 번호를 key로, shapes는 빈 dict로, 교체할 텍스트가 있으면 추정 shape id를 사용하세요.

{markitdown_text[:4000]}"""
    else:
        user_prompt = f"""## 템플릿 구조 (shape_id별 원본 텍스트)
{structured}

## markitdown 전체 텍스트 (참고용, 일부)
{markitdown_text[:2000]}"""

    user_prompt += f"""

## 새 주제
{topic}

## 추가 지시사항
{extra_info if extra_info else '없음'}

위 템플릿의 모든 텍스트 도형과 표 셀을 새 주제에 맞게 교체한 content_map을 생성하세요.
shape_id를 정확히 매칭하고, 각 텍스트의 원본 글자수를 참고하여 비슷한 길이로 작성하세요."""

    async with httpx.AsyncClient(timeout=180) as client:
        resp = await client.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {OPENROUTER_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MODEL,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.5,
                "max_tokens": 16000,
            },
        )
        resp.raise_for_status()
        data = resp.json()
        content = data["choices"][0]["message"]["content"]

    # JSON 파싱 (코드블록 제거 + 에러 핸들링)
    content = content.strip()
    if not content:
        raise ValueError("AI 응답이 비어있습니다. 다시 시도해 주세요.")

    # 코드블록 제거
    if content.startswith("```"):
        lines = content.split("\n")
        start = 1
        end = len(lines)
        for i in range(len(lines) - 1, 0, -1):
            if lines[i].strip() == "```":
                end = i
                break
        content = "\n".join(lines[start:end]).strip()

    # JSON 블록만 추출 (앞뒤 설명 텍스트 제거)
    brace_start = content.find("{")
    brace_end = content.rfind("}")
    if brace_start != -1 and brace_end != -1:
        content = content[brace_start:brace_end + 1]

    try:
        return json.loads(content)
    except json.JSONDecodeError as e:
        raise ValueError(f"AI 응답 JSON 파싱 실패: {str(e)[:100]}. 다시 시도해 주세요.")
