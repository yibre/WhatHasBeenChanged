"""
diff_engine_word.py — 두 워드 파일의 변경점을 git처럼 추적하는 엔진.

브라우저(Pyodide)와 일반 CPython에서 동일하게 동작한다.
외부 패키지 없이 표준 라이브러리(zipfile, xml.etree)만으로 docx를 읽는다.
(python-docx가 쓰는 lxml은 C 확장이라 이 프로젝트가 쓰는 경량 Pyodide 런타임에는 없다.)

알고리즘
  1. word/document.xml 을 열어 문단(w:p)·표 행(w:tbl 안의 w:tr)을 문서 순서대로
     하나의 "블록" 목록으로 만든다.
  2. 블록 텍스트를 키로 삼아 difflib(LCS)으로 정렬 → 블록 단위 추가/삭제/수정 판정.
  3. 수정으로 판정된 블록은 단어 단위로 다시 difflib을 돌려 "어떤 단어가
     지워지고 추가됐는지"까지 보여준다 (워드의 '변경 내용 추적'과 비슷한 느낌).

CLI 사용:
    python diff_engine_word.py 이전.docx 이후.docx
    python diff_engine_word.py 이전.docx 이후.docx --json out.json
"""

from __future__ import annotations

import io
import json
import re
import zipfile
from difflib import SequenceMatcher
from xml.etree import ElementTree as ET

__all__ = ["compare", "DEFAULTS"]

# 블록/단어 상태 코드 — 엑셀 엔진과 동일한 코드를 쓴다 (SAME=0, MOD=1, ADD=2, DEL=3)
SAME, MOD, ADD, DEL = 0, 1, 2, 3

DEFAULTS = {
    "ignore_case": False,   # 대소문자 무시
    "ignore_space": False,  # 공백 무시
    "trim": True,           # 앞뒤 공백 무시
}

MAX_BLOCKS = 20000          # 이보다 큰 문서는 잘라서 비교 (브라우저 보호)
MAX_CHANGE_LIST = 20000     # 변경 목록 최대 건수
VIEW_THRESHOLD = 400        # 블록이 이보다 많으면 변경 주변만 보여준다
VIEW_CONTEXT = 3            # 변경 주변 몇 블록까지 같이 보여줄지

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _qn(tag: str) -> str:
    return f"{{{_W_NS}}}{tag}"


# 암호/DRM(IRM)이 걸린 오피스 파일은 docx(zip)가 아니라 OLE2 복합 문서 컨테이너에 담겨 온다.
_OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _check_not_encrypted(data: bytes, name: str) -> None:
    if data[:8] == _OLE_SIGNATURE:
        raise ValueError(f"ENCRYPTED_FILE:{name}")


def _to_bytes(data) -> bytes:
    """JS Uint8Array / memoryview / bytes 무엇이 와도 bytes로."""
    if isinstance(data, (bytes, bytearray)):
        return bytes(data)
    to_py = getattr(data, "to_py", None)
    if to_py is not None:
        data = to_py()
    return bytes(data)


_CACHE: dict = {"sig": None, "a": None, "b": None}


# ---------------------------------------------------------------- 파일 읽기

def _paragraph_text(p_elem) -> str:
    """문단(w:p) 하나에서 보이는 글자만 순서대로 뽑아 이어 붙인다."""
    parts: list[str] = []
    for node in p_elem.iter():
        tag = node.tag
        if tag == _qn("t"):
            parts.append(node.text or "")
        elif tag == _qn("tab"):
            parts.append("\t")
        elif tag in (_qn("br"), _qn("cr")):
            parts.append("\n")
    return "".join(parts)


def _parse_body(xml_bytes: bytes) -> list[dict]:
    """word/document.xml 을 문서 순서 그대로 '블록' 목록으로 바꾼다."""
    root = ET.fromstring(xml_bytes)
    body = root.find(_qn("body"))
    blocks: list[dict] = []
    if body is None:
        return blocks

    for child in body:
        if child.tag == _qn("p"):
            blocks.append({"kind": "p", "text": _paragraph_text(child)})
        elif child.tag == _qn("tbl"):
            for tr in child.iter(_qn("tr")):
                cells = [
                    " ".join(_paragraph_text(p) for p in tc.findall(_qn("p")))
                    for tc in tr.findall(_qn("tc"))
                ]
                blocks.append({"kind": "row", "text": " | ".join(cells)})
    return blocks


def _read_docx(data: bytes) -> list[dict]:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            xml_bytes = z.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError) as e:
        raise ValueError("NOT_A_WORD_FILE") from e
    return _parse_body(xml_bytes)


# ---------------------------------------------------------------- 텍스트 정규화

def _mkkey(opts):
    trim, low, nospace = opts["trim"], opts["ignore_case"], opts["ignore_space"]

    def key(t: str) -> str:
        if nospace:
            t = "".join(t.split())
        elif trim:
            t = t.strip()
        if low:
            t = t.casefold()
        return t

    return key


_TOKEN_RE = re.compile(r"\s+|\S+")


def _tokenize(s: str) -> list[str]:
    return _TOKEN_RE.findall(s)


def _word_diff(old: str, new: str):
    """한 문단 안에서 어떤 단어가 지워지고 추가됐는지 difflib로 잘라낸다."""
    ta, tb = _tokenize(old), _tokenize(new)
    sm = SequenceMatcher(None, ta, tb, autojunk=False)
    spans, add_n, del_n = [], 0, 0
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            spans.append({"t": "same", "s": "".join(ta[i1:i2])})
        elif tag == "delete":
            spans.append({"t": "del", "s": "".join(ta[i1:i2])})
            del_n += sum(1 for w in ta[i1:i2] if w.strip())
        elif tag == "insert":
            spans.append({"t": "add", "s": "".join(tb[j1:j2])})
            add_n += sum(1 for w in tb[j1:j2] if w.strip())
        else:
            spans.append({"t": "del", "s": "".join(ta[i1:i2])})
            spans.append({"t": "add", "s": "".join(tb[j1:j2])})
            del_n += sum(1 for w in ta[i1:i2] if w.strip())
            add_n += sum(1 for w in tb[j1:j2] if w.strip())
    return spans, add_n, del_n


# ---------------------------------------------------------------- 블록 정렬

def _diff_blocks(blocks_a, blocks_b, opts, state):
    keyfn = _mkkey(opts)
    keys_a = [keyfn(b["text"]) for b in blocks_a]
    keys_b = [keyfn(b["text"]) for b in blocks_b]

    out = []
    for tag, i1, i2, j1, j2 in SequenceMatcher(None, keys_a, keys_b, autojunk=False).get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                ia, ib = i1 + k, j1 + k
                out.append({"s": SAME, "kind": blocks_b[ib]["kind"], "a": ia + 1, "b": ib + 1, "text": blocks_b[ib]["text"]})
        elif tag == "delete":
            for ia in range(i1, i2):
                out.append({"s": DEL, "kind": blocks_a[ia]["kind"], "a": ia + 1, "b": None, "text": blocks_a[ia]["text"]})
        elif tag == "insert":
            for ib in range(j1, j2):
                out.append({"s": ADD, "kind": blocks_b[ib]["kind"], "a": None, "b": ib + 1, "text": blocks_b[ib]["text"]})
        else:
            ai, bj = list(range(i1, i2)), list(range(j1, j2))
            n = min(len(ai), len(bj))
            for k in range(n):
                ia, ib = ai[k], bj[k]
                o, nv = blocks_a[ia]["text"], blocks_b[ib]["text"]
                if keys_a[ia] == keys_b[ib]:
                    out.append({"s": SAME, "kind": blocks_b[ib]["kind"], "a": ia + 1, "b": ib + 1, "text": nv})
                    continue
                spans, add_n, del_n = _word_diff(o, nv)
                out.append({"s": MOD, "kind": blocks_b[ib]["kind"], "a": ia + 1, "b": ib + 1, "o": o, "n": nv, "d": spans})
                state["word_add"] += add_n
                state["word_del"] += del_n
                if len(state["changes"]) < MAX_CHANGE_LIST:
                    state["changes"].append({"a": ia + 1, "b": ib + 1, "o": o, "n": nv, "kind": blocks_b[ib]["kind"]})
                else:
                    state["truncated"] = True
            for ia in ai[n:]:
                out.append({"s": DEL, "kind": blocks_a[ia]["kind"], "a": ia + 1, "b": None, "text": blocks_a[ia]["text"]})
            for ib in bj[n:]:
                out.append({"s": ADD, "kind": blocks_b[ib]["kind"], "a": None, "b": ib + 1, "text": blocks_b[ib]["text"]})
    return out


def _apply_view(blocks):
    """문단이 많은 문서는 바뀐 곳 주변만 보여주고 나머지는 접는다."""
    if len(blocks) <= VIEW_THRESHOLD:
        return blocks
    keep = [False] * len(blocks)
    for i, b in enumerate(blocks):
        if b["s"] != SAME:
            for k in range(max(0, i - VIEW_CONTEXT), min(len(blocks), i + VIEW_CONTEXT + 1)):
                keep[k] = True
    if not any(keep):
        return blocks[: VIEW_CONTEXT * 2]

    out, skipped = [], 0
    for i, b in enumerate(blocks):
        if keep[i]:
            if skipped:
                out.append({"gap": skipped})
                skipped = 0
            out.append(b)
        else:
            skipped += 1
    if skipped:
        out.append({"gap": skipped})
    return out


# ---------------------------------------------------------------- 공개 API

def compare(a_bytes, b_bytes, options=None, name_a="이전", name_b="이후") -> str:
    """두 워드 파일을 비교하고 결과를 JSON 문자열로 돌려준다."""
    opts = dict(DEFAULTS)
    if options:
        opts.update(json.loads(options) if isinstance(options, str) else dict(options))

    a_bytes, b_bytes = _to_bytes(a_bytes), _to_bytes(b_bytes)
    _check_not_encrypted(a_bytes, name_a)
    _check_not_encrypted(b_bytes, name_b)

    sig = (len(a_bytes), len(b_bytes), name_a, name_b)
    if _CACHE["sig"] != sig:
        _CACHE["a"] = _read_docx(a_bytes)
        _CACHE["b"] = _read_docx(b_bytes)
        _CACHE["sig"] = sig
    blocks_a, blocks_b = _CACHE["a"], _CACHE["b"]

    truncated_doc = len(blocks_a) > MAX_BLOCKS or len(blocks_b) > MAX_BLOCKS
    blocks_a, blocks_b = blocks_a[:MAX_BLOCKS], blocks_b[:MAX_BLOCKS]

    state = {"changes": [], "word_add": 0, "word_del": 0, "truncated": False}
    blocks = _diff_blocks(blocks_a, blocks_b, opts, state)

    summary = {"para_add": 0, "para_del": 0, "para_mod": 0, "word_add": state["word_add"], "word_del": state["word_del"]}
    for blk in blocks:
        if blk["s"] == ADD:
            summary["para_add"] += 1
        elif blk["s"] == DEL:
            summary["para_del"] += 1
        elif blk["s"] == MOD:
            summary["para_mod"] += 1

    result = {
        "ok": True,
        "files": {"a": name_a, "b": name_b},
        "options": opts,
        "summary": summary,
        "blocks": _apply_view(blocks),
        "changes": state["changes"],
        "para_limit": MAX_BLOCKS,
        "para_limit_hit": truncated_doc,
        "truncated": state["truncated"],
    }
    return json.dumps(result, ensure_ascii=False)


def reset() -> None:
    """새 파일을 올릴 때 읽어둔 문서 캐시를 비운다."""
    _CACHE["sig"] = _CACHE["a"] = _CACHE["b"] = None


# ---------------------------------------------------------------- CLI

def _cli():
    import argparse
    import os

    p = argparse.ArgumentParser(description="두 워드 파일의 변경점을 비교합니다.")
    p.add_argument("old")
    p.add_argument("new")
    p.add_argument("--ignore-case", action="store_true")
    p.add_argument("--ignore-space", action="store_true")
    p.add_argument("--json", help="결과 JSON 저장 경로")
    a = p.parse_args()

    opts = {"ignore_case": a.ignore_case, "ignore_space": a.ignore_space}
    with open(a.old, "rb") as f:
        ab = f.read()
    with open(a.new, "rb") as f:
        bb = f.read()

    out = compare(ab, bb, json.dumps(opts), os.path.basename(a.old), os.path.basename(a.new))
    res = json.loads(out)
    s = res["summary"]

    print(f"{res['files']['a']}  →  {res['files']['b']}")
    print(f"  문단 추가 {s['para_add']} · 삭제 {s['para_del']} · 수정 {s['para_mod']}")
    print(f"  단어 추가 {s['word_add']} · 삭제 {s['word_del']}")
    for ch in res["changes"][:40]:
        print(f"  [{ch['a']}→{ch['b']}] {ch['o']!r} → {ch['n']!r}")
    if len(res["changes"]) > 40:
        print(f"  ... 외 {len(res['changes']) - 40}건")

    if a.json:
        with open(a.json, "w", encoding="utf-8") as f:
            f.write(out)
        print(f"JSON 저장: {a.json}")


if __name__ == "__main__":
    _cli()
