"""PPTX 교안에서 프롬프트를 뽑아 prompts.js를 만든다.

사용법: python3 tools/build_data.py 교안.pptx prompts.js
"""
import json
import re
import sys

from pptx import Presentation

# 기준·요청·결과 줄에 있는 대괄호 중 사용자가 채우는 칸 (슬라이드 번호, 대괄호 원문)
# 나머지([확인 필요], [추정] 등)는 AI가 답에 붙이는 표시이므로 그대로 둔다.
FILL_IN = {
    (24, "[시험·행사]"), (27, "[사내 홍보 / 예산 심의 / 외부 제안]"), (28, "[N쪽]"),
    (38, "[의심형 / Risk 관리형 / 진취형 / 보수형]"), (46, "[의견수렴/의사결정]"),
    (47, "[ ]"), (49, "[요청 / 보고 / 공지 / 협의]"),
    (61, "[글로벌 선도기업 / 정부 정책 / 컨설팅 보고서]"), (75, "[익숙함 / 익숙하지 않음]"),
    (80, "[독자 직무]"), (81, "[서술형 개조식 / 명사형 종결]"), (87, "[주색 / 보조색]"),
    (94, "[타 부서 임원]"), (96, "[재무/현업/고객]"),
}
# 이 칸의 대괄호는 모두 사용자가 채우는 칸
FILL_ROWS = {"자료", "나", "독자", "주제", "결정", "방식"}


def walk(shapes):
    for s in shapes:
        if s.shape_type == 6:
            yield from walk(s.shapes)
        else:
            yield s


def join_wrapped(lines):
    """슬라이드에서 줄 맞춤용으로 끊은 줄을 이어 붙인다. '예:' 줄은 들여쓴 채 둔다."""
    out = []
    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue
        if out and line[:1].isspace():
            body = line.strip()
            if body.startswith("예:"):
                out.append("  " + body)
            else:
                out[-1] = out[-1].rstrip() + " " + body
        else:
            out.append(line)
    return out


def segments(slide, label, line):
    segs, pos = [], 0
    for m in re.finditer(r"\[[^\]]*\]", line):
        if label in FILL_ROWS or (slide, m.group()) in FILL_IN:
            if m.start() > pos:
                segs.append(line[pos:m.start()])
            segs.append({"f": m.group()[1:-1].strip()})
            pos = m.end()
    if pos < len(line):
        segs.append(line[pos:])
    return segs


def main(src, dst):
    prs = Presentation(src)
    chapters, cur = [], None
    for no, slide in enumerate(prs.slides, 1):
        ts = [s for s in walk(slide.shapes) if s.has_text_frame and s.text_frame.text.strip()]
        texts = [s.text_frame.text.strip() for s in ts]
        if no == 1:
            continue
        if len(ts) == 1:  # 챕터 표지
            cur = {"title": texts[0], "items": []}
            chapters.append(cur)
            continue
        k = texts.index("핵심")
        lab = next(j for j, s in enumerate(ts) if abs(s.top - 1879601) < 5000 and s.left > 3000000)
        item = {"id": no, "label": texts[lab], "core": texts[k + 1], "points": [], "rows": []}
        for j in range(k + 2, lab, 2):
            desc = " ".join(t.strip() for t in re.split(r"[\n\v]", texts[j + 1]))
            item["points"].append([re.sub(r"\s{2,}", " ", texts[j]), desc])
        for para in ts[lab + 1].text_frame.paragraphs:
            if not para.text.strip():
                continue
            name, _, body = para.text.partition("\t")
            lines = join_wrapped(body.split("\v"))
            item["rows"].append({"k": name.strip(), "v": [segments(no, name.strip(), l) for l in lines]})
        used = {k, k + 1, lab, lab + 1, *range(k + 2, lab)}
        for j, s in enumerate(ts):
            t = texts[j]
            if j in used or t == "!" or s.top == 0:
                continue
            if s.top == 215900:
                item["title"] = t
            elif 1000000 < s.top < 1100000 and s.left < 10000:
                item["sub"] = t
            elif s.top > 6000000:
                item["tip"] = t
            else:
                item.setdefault("notes", []).append(t.replace("\v", "\n"))
        cur["items"].append(item)

    with open(dst, "w", encoding="utf-8") as f:
        f.write("// tools/build_data.py로 생성한 파일입니다.\n")
        f.write("window.PROMPT_DATA = [\n")
        for ci, ch in enumerate(chapters):
            f.write('{"title": %s, "items": [\n' % json.dumps(ch["title"], ensure_ascii=False))
            f.write(",\n".join(json.dumps(it, ensure_ascii=False) for it in ch["items"]))
            f.write("\n]}" + ("," if ci < len(chapters) - 1 else "") + "\n")
        f.write("];\n")
    print(f"{len(chapters)}개 챕터, {sum(len(c['items']) for c in chapters)}개 항목")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
