"""학교생활기록부 문장 동일성 검토 엔진.

검토 기준
1. 학생과 학생 간 전체 내용의 모든 문장이 완전 동일
2. 학생과 학생 간 내용에서 붙어 있는 연속 3문장 이상이 완전 동일
3. 학생과 학생 간 전체 내용의 문자 단위 SequenceMatcher 유사도가 95% 이상

각 결과는 Streamlit 표시와 색상 포함 Excel 내보내기에 공통으로 사용된다.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from difflib import SequenceMatcher
import hashlib
from io import BytesIO
import html
import re
from typing import Iterable, Literal

import pandas as pd

MatchKind = Literal["exact", "sequence", "similar"]

KIND_LABELS: dict[MatchKind, str] = {
    "exact": "전체 내용 완전 동일",
    "sequence": "연속 3문장 이상 동일",
    "similar": "전체 내용 유사도 95% 이상",
}

KIND_COLORS: dict[MatchKind, str] = {
    "exact": "#C00000",      # 빨간색 계열
    "sequence": "#ED7D31",   # 오렌지색 계열
    "similar": "#1F4E78",    # 파란색 계열
}

# 전체 내용 완전 동일은 빨간색, 연속 3문장은 오렌지색, 전체 내용 유사도는 파란색을 우선한다.
KIND_PRIORITY: dict[MatchKind, int] = {"similar": 1, "sequence": 2, "exact": 3}

# 마침표·물음표·느낌표 뒤의 공백 또는 문자열 끝을 문장 경계로 사용한다.
SENTENCE_PATTERN = re.compile(r".+?(?:[.!?]+(?=\s|$)|$)", re.DOTALL)
MULTI_SPACE_PATTERN = re.compile(r"\s+")


@dataclass(frozen=True)
class StudentRecord:
    """검토 단위인 학생별·영역별 한 행의 표준 정보."""

    record_id: str
    source_file: str
    sheet_name: str
    display_type: str
    grade_class: str
    number: str
    category: str
    content: str
    order: int

    @property
    def comparison_key(self) -> tuple[str, str]:
        """서로 다른 영역·과목의 문장을 불필요하게 교차하지 않기 위한 범위 키."""
        return self.display_type, self.category

    @property
    def student_key(self) -> tuple[str, str]:
        return self.grade_class, self.number

    @property
    def student_label(self) -> str:
        return f"{self.grade_class} {self.number}번".strip()


@dataclass(frozen=True)
class SentenceMarker:
    """한 문장에 부착된 동일성 검출 결과."""

    kind: MatchKind
    sentence_index: int
    peer_record_id: str
    peer_source_file: str
    peer_grade_class: str
    peer_number: str
    similarity: float | None = None
    sequence_length: int | None = None


@dataclass
class ReviewResult:
    """한 기준 학생 행에 대한 누적 검토 결과."""

    record: StudentRecord
    sentences: list[str]
    markers: list[SentenceMarker] = field(default_factory=list)
    mode: Literal["file", "cross"] = "file"

    def has_match(self) -> bool:
        return bool(self.markers)

    def kinds(self) -> list[MatchKind]:
        return [kind for kind in ("exact", "sequence", "similar") if any(marker.kind == kind for marker in self.markers)]

    def marker_kind_for_sentence(self, sentence_index: int) -> MatchKind | None:
        candidates = [marker.kind for marker in self.markers if marker.sentence_index == sentence_index]
        if not candidates:
            return None
        return max(candidates, key=lambda kind: KIND_PRIORITY[kind])

    def marked_sentence_lines(self, kind: MatchKind) -> list[str]:
        indexes = sorted({marker.sentence_index for marker in self.markers if marker.kind == kind})
        lines: list[str] = []
        for sentence_index in indexes:
            sentence = self.sentences[sentence_index]
            if kind == "similar":
                maximum = max(
                    marker.similarity or 0
                    for marker in self.markers
                    if marker.kind == "similar" and marker.sentence_index == sentence_index
                )
                lines.append(f"{sentence} (유사도 {maximum:.1%})")
            else:
                lines.append(sentence)
        return lines

    def local_student_numbers(self) -> str:
        """같은 파일 검토용: 본인까지 포함한 동일 문장 사용 학생 번호 목록."""
        if not self.has_match():
            return ""
        peers = {
            (marker.peer_grade_class, marker.peer_number)
            for marker in self.markers
        }
        peers.add((self.record.grade_class, self.record.number))
        grade_classes = {grade_class for grade_class, _ in peers}
        ordered = sorted(peers, key=lambda value: (value[0], numeric_sort_key(value[1])))
        if len(grade_classes) == 1:
            numbers = ", ".join(number for _, number in ordered)
            return f"{numbers}번"
        return "; ".join(f"{grade_class} {number}번" for grade_class, number in ordered)

    def cross_student_numbers(self) -> str:
        """교차 검토용: 기준 파일과 같은 문장을 사용한 후속 파일 학생의 학년·반·번호 목록."""
        if not self.has_match():
            return ""
        grouped: dict[str, set[str]] = {}
        for marker in self.markers:
            grouped.setdefault(marker.peer_grade_class, set()).add(marker.peer_number)

        groups: list[str] = []
        for grade_class, numbers in sorted(grouped.items()):
            ordered_numbers = sorted(numbers, key=numeric_sort_key)
            groups.append(f"{grade_class} {', '.join(ordered_numbers)}번")
        return "; ".join(groups)


def numeric_sort_key(value: str) -> tuple[int, str]:
    match = re.fullmatch(r"\d+", str(value).strip())
    return (int(value), "") if match else (10**9, str(value))


def normalize_sentence(sentence: str) -> str:
    """완전 동일 판단용으로 공백을 하나로 통일한다. 문장 부호와 글자는 유지한다."""
    return MULTI_SPACE_PATTERN.sub(" ", sentence).strip()


def split_sentences(content: str) -> list[str]:
    """내용을 표시용 문장 목록으로 분할한다. 마침표 없는 마지막 문장도 보존한다."""
    cleaned = str(content or "").strip()
    if not cleaned:
        return []
    sentences = [match.group().strip() for match in SENTENCE_PATTERN.finditer(cleaned)]
    return [sentence for sentence in sentences if sentence]


def content_column_for(dataframe: pd.DataFrame) -> str | None:
    for column in ("세부능력 및 특기사항", "특기사항", "행동특성 및 종합의견"):
        if column in dataframe.columns:
            return column
    return None


def category_for_row(row: pd.Series) -> str:
    if "과목" in row.index:
        return str(row.get("과목") or "")
    if "영역" in row.index:
        return str(row.get("영역") or "")
    return "행동특성 및 종합의견"


def records_from_processed_items(processed_items: list[dict[str, object]]) -> dict[str, list[StudentRecord]]:
    """정제기 출력 항목을 원본 파일별 검토 레코드로 변환한다."""
    by_file: dict[str, list[StudentRecord]] = {}
    global_order = 0

    for item in processed_items:
        dataframe = item["dataframe"]
        if not isinstance(dataframe, pd.DataFrame):
            continue
        content_column = content_column_for(dataframe)
        if content_column is None:
            continue

        source_file = str(item["source_file"])
        sheet_name = str(item["sheet_name"])
        display_type = str(item["display_type"])
        target = by_file.setdefault(source_file, [])

        for row_index, row in dataframe.iterrows():
            content = str(row.get(content_column) or "").strip()
            number = str(row.get("번호") or "").strip()
            if not content or not number:
                continue
            grade_class = str(row.get("학년 반") or "").strip()
            category = category_for_row(row)
            record_id = f"{source_file}|{sheet_name}|{display_type}|{row_index}|{global_order}"
            target.append(
                StudentRecord(
                    record_id=record_id,
                    source_file=source_file,
                    sheet_name=sheet_name,
                    display_type=display_type,
                    grade_class=grade_class,
                    number=number,
                    category=category,
                    content=content,
                    order=global_order,
                )
            )
            global_order += 1
    return by_file


def add_pair_markers(
    left: StudentRecord,
    right: StudentRecord,
    left_markers: dict[str, list[SentenceMarker]],
    right_markers: dict[str, list[SentenceMarker]],
) -> None:
    """두 학생의 전체 내용을 세 가지 기준으로 비교하고 해당 문장에 마커를 부착한다."""
    left_sentences = split_sentences(left.content)
    right_sentences = split_sentences(right.content)
    if not left_sentences or not right_sentences:
        return

    left_keys = [normalize_sentence(sentence) for sentence in left_sentences]
    right_keys = [normalize_sentence(sentence) for sentence in right_sentences]
    left_content = normalize_sentence(left.content)
    right_content = normalize_sentence(right.content)

    def append_markers(
        target: StudentRecord,
        peer: StudentRecord,
        target_markers: dict[str, list[SentenceMarker]],
        kind: MatchKind,
        indexes: Iterable[int],
        similarity: float | None = None,
        sequence_length: int | None = None,
    ) -> None:
        for sentence_index in indexes:
            target_markers[target.record_id].append(
                SentenceMarker(
                    kind=kind,
                    sentence_index=sentence_index,
                    peer_record_id=peer.record_id,
                    peer_source_file=peer.source_file,
                    peer_grade_class=peer.grade_class,
                    peer_number=peer.number,
                    similarity=similarity,
                    sequence_length=sequence_length,
                )
            )

    # 1) 전체 내용의 모든 문장이 완전 동일하면 전체를 빨간색으로 표시한다.
    if left_content == right_content:
        append_markers(left, right, left_markers, "exact", range(len(left_sentences)), similarity=1.0)
        append_markers(right, left, right_markers, "exact", range(len(right_sentences)), similarity=1.0)
        return

    # 2) 내용 안에서 순서가 유지된 연속 3문장 이상의 완전 동일 구간을 오렌지색으로 표시한다.
    matcher = SequenceMatcher(None, left_keys, right_keys, autojunk=False)
    for block in matcher.get_matching_blocks():
        if block.size < 3:
            continue
        append_markers(
            left,
            right,
            left_markers,
            "sequence",
            range(block.a, block.a + block.size),
            sequence_length=block.size,
        )
        append_markers(
            right,
            left,
            right_markers,
            "sequence",
            range(block.b, block.b + block.size),
            sequence_length=block.size,
        )

    # 3) 일부 조사·부사·활동명 수정처럼 전체 내용의 유사도가 95% 이상이면 전체를 파란색으로 표시한다.
    content_similarity = SequenceMatcher(None, left_content, right_content, autojunk=False).ratio()
    if content_similarity >= 0.95:
        append_markers(
            left,
            right,
            left_markers,
            "similar",
            range(len(left_sentences)),
            similarity=content_similarity,
        )
        append_markers(
            right,
            left,
            right_markers,
            "similar",
            range(len(right_sentences)),
            similarity=content_similarity,
        )


TOKEN_PATTERN = re.compile(r"[0-9A-Za-z가-힣]{2,}")
MAX_TOKEN_BUCKET_SIZE = 60


def stable_hash(value: str) -> str:
    """정규화된 텍스트를 안정적인 SHA-256 해시로 변환한다."""
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def content_hash(record: StudentRecord) -> str:
    return stable_hash(normalize_sentence(record.content))


def sentence_window_hashes(record: StudentRecord, window_size: int = 3) -> set[str]:
    """연속 3문장 동일 후보를 찾기 위한 문장 창 해시 집합을 반환한다."""
    keys = [normalize_sentence(sentence) for sentence in split_sentences(record.content)]
    if len(keys) < window_size:
        return set()
    return {
        stable_hash("\x1f".join(keys[start:start + window_size]))
        for start in range(len(keys) - window_size + 1)
    }


def content_tokens(record: StudentRecord) -> set[str]:
    """유사도 후보 필터에 쓰는 2글자 이상 토큰 집합을 반환한다."""
    return set(TOKEN_PATTERN.findall(normalize_sentence(record.content).lower()))


def can_reach_similarity_threshold(left: StudentRecord, right: StudentRecord, threshold: float = 0.95) -> bool:
    """문자열 길이만으로도 95%에 도달할 수 없는 쌍을 즉시 제외한다."""
    left_length = len(normalize_sentence(left.content))
    right_length = len(normalize_sentence(right.content))
    if not left_length or not right_length:
        return False
    maximum_possible_ratio = (2 * min(left_length, right_length)) / (left_length + right_length)
    return maximum_possible_ratio >= threshold


def canonical_pair(left: StudentRecord, right: StudentRecord) -> tuple[StudentRecord, StudentRecord]:
    return (left, right) if left.order <= right.order else (right, left)


def append_pairs_from_buckets(
    left_buckets: dict[str, list[StudentRecord]],
    right_buckets: dict[str, list[StudentRecord]],
    same_collection: bool,
    pairs: set[tuple[StudentRecord, StudentRecord]],
    require_similarity_length: bool = False,
) -> None:
    """같은 해시·토큰 버킷에 속한 학생만 후보 쌍으로 추가한다."""
    for key, left_records in left_buckets.items():
        right_records = right_buckets.get(key, [])
        if not right_records:
            continue
        if same_collection:
            for index, left in enumerate(left_records):
                for right in left_records[index + 1:]:
                    if left.student_key == right.student_key:
                        continue
                    if require_similarity_length and not can_reach_similarity_threshold(left, right):
                        continue
                    pairs.add(canonical_pair(left, right))
        else:
            for left in left_records:
                for right in right_records:
                    if left.student_key == right.student_key:
                        continue
                    if require_similarity_length and not can_reach_similarity_threshold(left, right):
                        continue
                    pairs.add((left, right))


def build_candidate_pairs(
    left_records: list[StudentRecord],
    right_records: list[StudentRecord] | None = None,
) -> tuple[set[tuple[StudentRecord, StudentRecord]], set[tuple[StudentRecord, StudentRecord]]]:
    """해시와 토큰 인덱스로 정밀 비교가 필요한 학생 쌍만 추린다.

    첫 번째 반환값은 전체 내용 완전 동일 후보이고, 두 번째는 연속 3문장 또는
    전체 유사도 정밀 검사를 통과해야 하는 후보이다.
    """
    same_collection = right_records is None
    target_records = left_records if same_collection else right_records or []

    left_exact: dict[str, list[StudentRecord]] = defaultdict(list)
    right_exact: dict[str, list[StudentRecord]] = defaultdict(list)
    left_windows: dict[str, list[StudentRecord]] = defaultdict(list)
    right_windows: dict[str, list[StudentRecord]] = defaultdict(list)
    left_tokens: dict[str, list[StudentRecord]] = defaultdict(list)
    right_tokens: dict[str, list[StudentRecord]] = defaultdict(list)

    for record in left_records:
        left_exact[content_hash(record)].append(record)
        for window_hash in sentence_window_hashes(record):
            left_windows[window_hash].append(record)
        for token in content_tokens(record):
            left_tokens[token].append(record)

    for record in target_records:
        right_exact[content_hash(record)].append(record)
        for window_hash in sentence_window_hashes(record):
            right_windows[window_hash].append(record)
        for token in content_tokens(record):
            right_tokens[token].append(record)

    exact_pairs: set[tuple[StudentRecord, StudentRecord]] = set()
    append_pairs_from_buckets(left_exact, right_exact, same_collection, exact_pairs)

    # 연속 3문장 해시가 같으면 문장 배열 정밀 비교 후보로 등록한다.
    candidate_pairs: set[tuple[StudentRecord, StudentRecord]] = set()
    append_pairs_from_buckets(left_windows, right_windows, same_collection, candidate_pairs)

    # 토큰이 지나치게 넓게 공유되는 버킷은 무시하고, 길이 상 95% 도달 가능 후보만 남긴다.
    bounded_left_tokens = {
        token: records
        for token, records in left_tokens.items()
        if len(records) <= MAX_TOKEN_BUCKET_SIZE
    }
    bounded_right_tokens = {
        token: records
        for token, records in right_tokens.items()
        if len(records) <= MAX_TOKEN_BUCKET_SIZE
    }
    append_pairs_from_buckets(
        bounded_left_tokens,
        bounded_right_tokens,
        same_collection,
        candidate_pairs,
        require_similarity_length=True,
    )
    candidate_pairs.difference_update(exact_pairs)
    return exact_pairs, candidate_pairs


def comparison_candidate_counts(
    left_records: list[StudentRecord],
    right_records: list[StudentRecord] | None = None,
) -> dict[str, int]:
    """성능 검증용으로 전체 가능 쌍과 해시 필터 통과 후보 수를 반환한다."""
    same_collection = right_records is None
    target_records = left_records if same_collection else right_records or []
    if same_collection:
        total_pairs = len(left_records) * max(len(left_records) - 1, 0) // 2
    else:
        total_pairs = len(left_records) * len(target_records)
    exact_pairs, candidate_pairs = build_candidate_pairs(left_records, right_records)
    return {
        "전체 가능 쌍": total_pairs,
        "완전 동일 해시 후보": len(exact_pairs),
        "정밀 비교 후보": len(candidate_pairs),
        "정밀 비교 회피 쌍": max(total_pairs - len(exact_pairs) - len(candidate_pairs), 0),
    }


def make_results(
    records: Iterable[StudentRecord],
    markers: dict[str, list[SentenceMarker]],
    mode: Literal["file", "cross"],
) -> list[ReviewResult]:
    return [
        ReviewResult(
            record=record,
            sentences=split_sentences(record.content),
            markers=markers.get(record.record_id, []),
            mode=mode,
        )
        for record in sorted(records, key=lambda item: item.order)
    ]


def review_within_file(records: list[StudentRecord]) -> list[ReviewResult]:
    """같은 파일의 학생을 해시 필터링 후 필요한 후보 쌍만 정밀 검토한다."""
    markers: dict[str, list[SentenceMarker]] = {record.record_id: [] for record in records}
    grouped: dict[tuple[str, str], list[StudentRecord]] = {}
    for record in records:
        grouped.setdefault(record.comparison_key, []).append(record)

    for group_records in grouped.values():
        exact_pairs, candidate_pairs = build_candidate_pairs(group_records)
        # 해시로 같은 전체 내용을 찾은 쌍은 즉시 완전 동일 처리하고, 나머지만 정밀 비교한다.
        for left, right in exact_pairs | candidate_pairs:
            add_pair_markers(left, right, markers, markers)
    return make_results(records, markers, mode="file")


def review_cross_files(
    base_records: list[StudentRecord],
    comparison_records: list[StudentRecord],
) -> list[ReviewResult]:
    """기준 파일과 후속 파일을 해시·토큰 후보 필터를 거쳐 교차 검토한다."""
    base_markers: dict[str, list[SentenceMarker]] = {record.record_id: [] for record in base_records}
    comparison_markers: dict[str, list[SentenceMarker]] = {
        record.record_id: [] for record in comparison_records
    }
    base_by_key: dict[tuple[str, str], list[StudentRecord]] = defaultdict(list)
    comparison_by_key: dict[tuple[str, str], list[StudentRecord]] = defaultdict(list)
    for record in base_records:
        base_by_key[record.comparison_key].append(record)
    for record in comparison_records:
        comparison_by_key[record.comparison_key].append(record)

    for comparison_key, grouped_base in base_by_key.items():
        grouped_comparison = comparison_by_key.get(comparison_key, [])
        if not grouped_comparison:
            continue
        exact_pairs, candidate_pairs = build_candidate_pairs(grouped_base, grouped_comparison)
        for left, right in exact_pairs | candidate_pairs:
            add_pair_markers(left, right, base_markers, comparison_markers)
    return make_results(base_records, base_markers, mode="cross")


def review_all_files(
    records_by_file: dict[str, list[StudentRecord]],
    upload_order: list[str],
) -> tuple[dict[str, list[ReviewResult]], list[ReviewResult], str | None]:
    """파일별 검토와 기준 파일 우선 교차 검토를 함께 수행한다."""
    file_results = {
        source_file: review_within_file(records)
        for source_file, records in records_by_file.items()
    }

    available_order = [source_file for source_file in upload_order if source_file in records_by_file]
    if len(available_order) < 2:
        return file_results, [], available_order[0] if available_order else None

    base_file = available_order[0]
    subsequent_records = [
        record
        for source_file in available_order[1:]
        for record in records_by_file[source_file]
    ]
    cross_results = review_cross_files(records_by_file[base_file], subsequent_records)
    return file_results, cross_results, base_file


def review_table(results: list[ReviewResult], matches_only: bool = False, cross: bool = False) -> pd.DataFrame:
    """화면 표와 Excel 내보내기의 기본 열을 생성한다. 학생 번호 열을 항상 마지막에 둔다."""
    rows: list[dict[str, str]] = []
    for result in results:
        if matches_only and not result.has_match():
            continue
        record = result.record
        rows.append(
            {
                "검토 대상 영역": record.category,
                "학년 반": record.grade_class,
                "번호": record.number,
                "검토 내용": record.content,
                "전체 내용 완전 동일": "\n".join(result.marked_sentence_lines("exact")),
                "연속 3문장 이상 동일": "\n".join(result.marked_sentence_lines("sequence")),
                "전체 내용 유사도 95% 이상": "\n".join(result.marked_sentence_lines("similar")),
                "동일 문장 사용 학생 번호" if not cross else "교차 동일 문장 사용 학생 번호": (
                    result.cross_student_numbers() if cross else result.local_student_numbers()
                ),
            }
        )
    return pd.DataFrame(rows)


def colored_content_html(result: ReviewResult) -> str:
    """Streamlit HTML 미리보기용 문장별 색상 표시 문자열을 만든다."""
    fragments: list[str] = []
    for index, sentence in enumerate(result.sentences):
        kind = result.marker_kind_for_sentence(index)
        escaped = html.escape(sentence)
        if kind:
            fragments.append(f'<span style="color:{KIND_COLORS[kind]}; font-weight:700">{escaped}</span>')
        else:
            fragments.append(escaped)
    return " ".join(fragments)


def _write_review_sheet(
    workbook: object,
    worksheet: object,
    results: list[ReviewResult],
    matches_only: bool,
    cross: bool,
) -> None:
    """xlsxwriter 워크시트에 문장별 글자색을 포함한 검토 결과를 기록한다."""
    # xlsxwriter의 타입 스텁이 제한적이므로 object로 받아 런타임 API를 사용한다.
    header_format = workbook.add_format(
        {"bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF", "border": 1, "text_wrap": True}
    )
    cell_format = workbook.add_format({"valign": "top", "text_wrap": True, "border": 1})
    exact_format = workbook.add_format({"font_color": KIND_COLORS["exact"]})
    sequence_format = workbook.add_format({"font_color": KIND_COLORS["sequence"]})
    similar_format = workbook.add_format({"font_color": KIND_COLORS["similar"]})
    format_by_kind = {"exact": exact_format, "sequence": sequence_format, "similar": similar_format}
    result_cell_formats = {
        "exact": workbook.add_format({"valign": "top", "text_wrap": True, "border": 1, "font_color": KIND_COLORS["exact"]}),
        "sequence": workbook.add_format({"valign": "top", "text_wrap": True, "border": 1, "font_color": KIND_COLORS["sequence"]}),
        "similar": workbook.add_format({"valign": "top", "text_wrap": True, "border": 1, "font_color": KIND_COLORS["similar"]}),
    }

    number_header = "교차 동일 문장 사용 학생 번호" if cross else "동일 문장 사용 학생 번호"
    headers = [
        "검토 대상 영역",
        "학년 반",
        "번호",
        "검토 내용",
        "전체 내용 완전 동일",
        "연속 3문장 이상 동일",
        "전체 내용 유사도 95% 이상",
        number_header,
    ]
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, header_format)

    row_number = 1
    for result in results:
        if matches_only and not result.has_match():
            continue
        record = result.record
        values = [
            record.category,
            record.grade_class,
            record.number,
            record.content,
            "\n".join(result.marked_sentence_lines("exact")),
            "\n".join(result.marked_sentence_lines("sequence")),
            "\n".join(result.marked_sentence_lines("similar")),
            result.cross_student_numbers() if cross else result.local_student_numbers(),
        ]
        for col, value in enumerate(values):
            if col != 3:
                result_kind = {4: "exact", 5: "sequence", 6: "similar"}.get(col)
                worksheet.write(row_number, col, value, result_cell_formats[result_kind] if result_kind else cell_format)
                continue

            # 원문 내용을 문장 단위로 조각내고, 우선순위 색을 적용한다.
            formatted_fragments: list[object] = []
            for sentence_index, sentence in enumerate(result.sentences):
                if sentence_index:
                    formatted_fragments.extend([cell_format, " "])
                kind = result.marker_kind_for_sentence(sentence_index)
                formatted_fragments.extend([format_by_kind[kind] if kind else cell_format, sentence])
            if len(result.sentences) <= 1:
                worksheet.write(row_number, col, record.content, format_by_kind[result.marker_kind_for_sentence(0)] if result.sentences and result.marker_kind_for_sentence(0) else cell_format)
            else:
                worksheet.write_rich_string(row_number, col, *formatted_fragments, cell_format)
        row_number += 1

    worksheet.freeze_panes(1, 0)
    worksheet.autofilter(0, 0, max(row_number - 1, 0), len(headers) - 1)
    worksheet.set_column(0, 2, 18)
    worksheet.set_column(3, 3, 70)
    worksheet.set_column(4, 6, 45)
    worksheet.set_column(7, 7, 32)
    worksheet.set_default_row(18)


def build_review_excel(
    file_results: dict[str, list[ReviewResult]],
    cross_results: list[ReviewResult],
    matches_only: bool = True,
) -> bytes:
    """파일별 시트와 교차 검토 시트를 포함하는 색상 Excel 통합문서를 생성한다."""
    try:
        import xlsxwriter  # noqa: F401
    except ImportError as exc:
        raise RuntimeError("Excel 색상 내보내기에는 XlsxWriter 패키지가 필요합니다.") from exc

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        used_sheet_names: set[str] = set()
        for file_index, (_, results) in enumerate(file_results.items(), start=1):
            sheet_name = f"파일별 검토 {file_index}"
            while sheet_name in used_sheet_names:
                sheet_name = f"파일별 검토 {file_index}_{len(used_sheet_names)}"[:31]
            used_sheet_names.add(sheet_name)
            worksheet = workbook.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = worksheet
            _write_review_sheet(workbook, worksheet, results, matches_only=matches_only, cross=False)

        cross_sheet = workbook.add_worksheet("교차 검토")
        writer.sheets["교차 검토"] = cross_sheet
        _write_review_sheet(workbook, cross_sheet, cross_results, matches_only=matches_only, cross=True)
    return output.getvalue()
