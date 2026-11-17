import json
import re
from pathlib import Path

from docx import Document


SOURCE = Path("2026年下半年中钞信达反假考试题库.docx")
OUTPUT = Path("questions.js")


CHAPTER_RE = re.compile(r"^(第[一二三四五六七八九十]+章)\s*(.*?)（\d+）$")
TYPE_RE = re.compile(r"^(单选题|多选题|判断题)(?:（\d+）)?$")
NUMBER_RE = re.compile(r"^([0-9]+)[、.．]\s*(.*)$")
OPTION_RE = re.compile(r"^([A-H])[、.．]\s*(.*)$")
ANSWER_RE = re.compile(r"^【答案】\s*(.*)$")


def clean(value):
    return re.sub(r"\s+", " ", value).strip()


def finish_question(question):
    if not question:
        return None
    question["question"] = clean(question["question"])
    question["options"] = [{"key": key, "text": clean(text)} for key, text in question["options"]]
    answer = clean(question.get("answer", ""))
    if not answer:
        # Two source paragraphs contain the answer directly after the question.
        match = re.search(r"(?<=[）。.!！？?])([A-H]{1,8})$", question["question"])
        if match:
            answer = match.group(1)
            question["question"] = question["question"][: match.start()].rstrip()
    if answer in {"正确", "错误"}:
        question["answerKeys"] = []
        question["answerText"] = answer
    else:
        keys = re.findall(r"[A-H]", answer.upper())
        question["answerKeys"] = keys
        question["answerText"] = "".join(keys)
    question["searchText"] = " ".join(
        [
            question["chapter"],
            question["section"],
            question["type"],
            question["question"],
            " ".join(text for _, text in question["options"]),
            question["answerText"],
        ]
    )
    return question


def main():
    paragraphs = [clean(paragraph.text) for paragraph in Document(SOURCE).paragraphs]
    questions = []
    chapter = ""
    section = ""
    question_type = ""
    current = None

    def push_current():
        nonlocal current
        item = finish_question(current)
        if item:
            item["id"] = len(questions) + 1
            questions.append(item)
        current = None

    for text in paragraphs:
        if not text:
            continue
        chapter_match = CHAPTER_RE.match(text)
        if chapter_match:
            push_current()
            chapter = chapter_match.group(1) + " " + chapter_match.group(2)
            section = ""
            question_type = ""
            continue
        type_match = TYPE_RE.match(text)
        if type_match:
            push_current()
            question_type = type_match.group(1)
            continue
        # Subsections include both numbered headings and labels such as “三号令”.
        if not NUMBER_RE.match(text) and not OPTION_RE.match(text) and not ANSWER_RE.match(text):
            if re.search(r"（\d+）$", text):
                push_current()
                section = re.sub(r"（\d+）$", "", text).strip()
                continue
        number_match = NUMBER_RE.match(text)
        if number_match:
            push_current()
            current = {
                "number": int(number_match.group(1)),
                "chapter": chapter,
                "section": section,
                "type": question_type or "案例题",
                "question": number_match.group(2),
                "options": [],
                "answer": "",
            }
            continue
        if current is None:
            continue
        answer_match = ANSWER_RE.match(text)
        if answer_match:
            current["answer"] = answer_match.group(1)
            continue
        option_match = OPTION_RE.match(text)
        if option_match:
            current["options"].append((option_match.group(1), option_match.group(2)))
            continue
        if text.startswith("注：") or text.startswith("说明："):
            current["question"] += " " + text
        elif not current["options"]:
            current["question"] += text

    push_current()
    payload = "window.QUESTION_BANK = " + json.dumps(questions, ensure_ascii=False, separators=(",", ":")) + ";\n"
    OUTPUT.write_text(payload, encoding="utf-8")
    print(f"Wrote {len(questions)} questions to {OUTPUT}")


if __name__ == "__main__":
    main()
