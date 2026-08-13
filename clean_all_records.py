"""학교생활기록부 엑셀 자동 판별·정제 도구.

지원 영역: 교과 세부능력 및 특기사항, 창의적체험활동상황, 자유학기활동상황,
행동특성 및 종합의견. 파일명에 의존하지 않고 각 시트의 제목 및 헤더를 분석한다.
"""

from __future__ import annotations

import argparse
import math
import re
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import BinaryIO, Literal

import pandas as pd

RecordType = Literal["subject", "activity_creative", "activity_free", "behavior"]

GRADE_CLASS_PATTERN = re.compile(r"(\d+)\s*학년\s*(\d+)\s*반")
FREE_SEMESTER_ITEM_PATTERN = re.compile(r"(?<!^)(\([^\n()]+\)\(\d+\s*시간\))")

DISPLAY_NAMES: dict[RecordType, str] = {
    "subject": "교과 세부능력 및 특기사항",
    "activity_creative": "창의적체험활동상황",
    "activity_free": "자유학기활동상황",
    "behavior": "행동특성 및 종합의견",
}

MASTER_FILENAMES = {
    "subject": "master_subject_records.xlsx",
    "activity": "master_activity_records.xlsx",
    "behavior": "master_behavior_records.xlsx",
}


@dataclass
class ColumnLayout:
    """원본 출력 양식에서 탐지한 주요 열 위치."""

    content_header_row: int
    number_col: int | None
    name_col: int | None
    category_col: int | None
    content_col: int


def as_text(value: object) -> str:
    """빈 셀·NaN을 빈 문자열로, 나머지는 양 끝 공백을 제거한 문자열로 변환한다."""
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def normalized(value: object) -> str:
    """헤더 탐지를 위해 공백과 줄바꿈을 제거한 문자열을 반환한다."""
    return re.sub(r"\s+", "", as_text(value))


def canonical_number(value: object) -> str:
    """엑셀의 1.0 같은 학번 표기를 1로 통일한다."""
    text = as_text(value)
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def row_values(df: pd.DataFrame, row_index: int) -> list[str]:
    return [as_text(value) for value in df.iloc[row_index].tolist()]


def row_joined_normalized(df: pd.DataFrame, row_index: int) -> str:
    return "".join(normalized(value) for value in df.iloc[row_index].tolist())


def infer_grade_class(df: pd.DataFrame) -> str:
    """시트 전체에서 첫 번째 'n학년 n반' 표기를 추출한다."""
    for value in df.to_numpy().flatten():
        match = GRADE_CLASS_PATTERN.search(as_text(value))
        if match:
            return f"{match.group(1)}학년 {match.group(2)}반"
    return ""


def detect_record_type(df: pd.DataFrame) -> RecordType | None:
    """파일명 대신 제목·반복 헤더로 영역을 자동 판별한다."""
    # 표제와 헤더는 보통 첫 페이지 상단에 있으나, 병합 셀을 고려해 전체 시트를 확인한다.
    text = "\n".join(as_text(value) for value in df.to_numpy().flatten() if as_text(value))
    compact = normalized(text)

    if "세부능력및특기사항" in compact:
        return "subject"
    if "행동특성및종합의견" in compact:
        return "behavior"
    if "자유학기활동상황" in compact:
        return "activity_free"
    if "창의적체험활동상황" in compact:
        return "activity_creative"
    return None


def _find_label_column(df: pd.DataFrame, label: str, start_row: int, end_row: int) -> int | None:
    for row_index in range(max(0, start_row), min(len(df), end_row + 1)):
        for column_index, value in enumerate(row_values(df, row_index)):
            if label in normalized(value):
                return column_index
    return None


def find_layout(df: pd.DataFrame, record_type: RecordType) -> ColumnLayout | None:
    """첫 헤더 묶음에서 번호·성명·영역·내용 열을 동적으로 찾는다.

    원본 양식은 '번호/성명'과 '특기사항'이 서로 다른 두 줄에 놓일 수 있으므로,
    내용 헤더 기준 위아래 3행을 하나의 헤더 묶음으로 검사한다.
    """
    if record_type == "subject":
        content_label = "세부능력및특기사항"
        category_label = "과목"
    elif record_type in ("activity_creative", "activity_free"):
        content_label = "특기사항"
        category_label = "영역"
    else:
        content_label = "행동특성및종합의견"
        category_label = None

    for row_index in range(len(df)):
        content_col = _find_label_column(df, content_label, row_index, row_index)
        if content_col is None:
            continue

        # 활동 특기사항은 문장 본문에도 나타날 수 있으므로, 인근에 번호 또는 성명 헤더가 있어야 한다.
        number_col = _find_label_column(df, "번호", row_index - 3, row_index + 1)
        name_col = _find_label_column(df, "성명", row_index - 3, row_index + 1)
        category_col = (
            _find_label_column(df, category_label, row_index - 1, row_index + 1)
            if category_label
            else None
        )

        if record_type in ("activity_creative", "activity_free") and category_col is None:
            continue
        if record_type in ("subject", "behavior") and number_col is None:
            continue

        return ColumnLayout(
            content_header_row=row_index,
            number_col=number_col,
            name_col=name_col,
            category_col=category_col,
            content_col=content_col,
        )
    return None


def is_repeated_header(row: list[str], record_type: RecordType) -> bool:
    """페이지마다 반복되는 열 이름 행을 데이터에서 제외한다."""
    compact = "".join(normalized(value) for value in row)
    if "번호" in compact and "성명" in compact:
        return True
    if record_type in ("activity_creative", "activity_free"):
        return "특기사항" in compact and ("영역" in compact or "시간" in compact)
    if record_type == "subject":
        # 일부 양식은 과목·번호와 내용 헤더를 서로 다른 행에 배치한다.
        return "세부능력및특기사항" in compact
    # 행동특성 양식도 번호/성명 헤더와 내용 헤더가 분리될 수 있다.
    return "행동특성및종합의견" in compact


def is_metadata_or_page_footer(row: list[str]) -> bool:
    """표제, 사용자명, 페이지 표시 등 학생 데이터가 아닌 행을 식별한다."""
    values = [value for value in row if value]
    compact = "".join(normalized(value) for value in values)
    if not values:
        return True
    if "학교생활기록부" in compact or "사용자명" in compact:
        return True
    # 출력 하단의 '1 / 30 학교명'과 같은 페이징 행을 처리한다.
    if "/" in values and len(values) <= 5:
        return True
    return False


def clean_subject_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    layout = find_layout(df, "subject")
    columns = ["학년 반", "번호", "과목", "세부능력 및 특기사항"]
    if layout is None:
        return pd.DataFrame(columns=columns)

    grade_class = infer_grade_class(df)
    records: list[dict[str, str]] = []
    current_subject = ""

    for row_index in range(layout.content_header_row + 1, len(df)):
        row = row_values(df, row_index)
        if is_metadata_or_page_footer(row) or is_repeated_header(row, "subject"):
            continue

        number = canonical_number(row[layout.number_col]) if layout.number_col is not None else ""
        name = as_text(row[layout.name_col]) if layout.name_col is not None else ""
        subject = as_text(row[layout.category_col]) if layout.category_col is not None else ""
        content = as_text(row[layout.content_col])

        if subject:
            current_subject = subject
        if not any([number, name, content]):
            continue

        same_student = bool(records and number and number == records[-1]["번호"])
        continuation = (not number and not name) or (same_student and current_subject == records[-1]["과목"])
        if continuation and records:
            if content:
                records[-1]["세부능력 및 특기사항"] = (
                    records[-1]["세부능력 및 특기사항"] + " " + content
                ).strip()
        else:
            records.append(
                {
                    "학년 반": grade_class,
                    "번호": number,
                    "과목": current_subject,
                    "세부능력 및 특기사항": content,
                }
            )

    result = pd.DataFrame(records, columns=columns)
    if result.empty:
        return result
    result = result[result["세부능력 및 특기사항"].str.strip().ne("")]
    return result.drop_duplicates(subset=columns, keep="first").reset_index(drop=True)


def clean_activity_dataframe(df: pd.DataFrame, is_free_semester: bool) -> pd.DataFrame:
    """창체·자유학기활동을 정제한다.

    정적 영역 목록을 사용하지 않고, 원본의 '영역' 열에 있는 모든 값을 그대로 인식한다.
    따라서 자율·자치, 동아리, 진로, 주제선택, 진로탐색 등 새 영역에도 대응한다.
    """
    record_type: RecordType = "activity_free" if is_free_semester else "activity_creative"
    layout = find_layout(df, record_type)
    columns = ["학년 반", "번호", "영역", "특기사항"]
    if layout is None:
        return pd.DataFrame(columns=columns)

    grade_class = infer_grade_class(df)
    records: list[dict[str, str]] = []
    current_area = ""

    for row_index in range(layout.content_header_row + 1, len(df)):
        row = row_values(df, row_index)
        if is_metadata_or_page_footer(row) or is_repeated_header(row, record_type):
            continue

        number = canonical_number(row[layout.number_col]) if layout.number_col is not None else ""
        name = as_text(row[layout.name_col]) if layout.name_col is not None else ""
        area = as_text(row[layout.category_col]) if layout.category_col is not None else ""
        content = as_text(row[layout.content_col])

        if area:
            current_area = area
        if not any([number, name, content]):
            continue

        if is_free_semester and current_area == "주제선택활동" and content:
            # (활동명)(시간)으로 시작하는 각 선택활동을 줄바꿈으로 명확히 구분한다.
            content = FREE_SEMESTER_ITEM_PATTERN.sub(r"\n\1", content)

        same_student = bool(records and number and number == records[-1]["번호"])
        continuation = (not number and not name) or (same_student and current_area == records[-1]["영역"])
        if continuation and records:
            if content:
                joiner = " "
                if is_free_semester and current_area == "주제선택활동" and re.match(
                    r"^\([^\n()]+\)\(\d+\s*시간\)", content
                ):
                    joiner = "\n"
                records[-1]["특기사항"] = (records[-1]["특기사항"] + joiner + content).strip()
        else:
            records.append(
                {
                    "학년 반": grade_class,
                    "번호": number,
                    "영역": current_area,
                    "특기사항": content,
                }
            )

    result = pd.DataFrame(records, columns=columns)
    if result.empty:
        return result
    result = result[result["특기사항"].str.strip().ne("")]
    return result.drop_duplicates(subset=columns, keep="first").reset_index(drop=True)


def clean_behavior_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    layout = find_layout(df, "behavior")
    columns = ["학년 반", "번호", "행동특성 및 종합의견"]
    if layout is None:
        return pd.DataFrame(columns=columns)

    grade_class = infer_grade_class(df)
    records: list[dict[str, str]] = []

    for row_index in range(layout.content_header_row + 1, len(df)):
        row = row_values(df, row_index)
        if is_metadata_or_page_footer(row) or is_repeated_header(row, "behavior"):
            continue

        number = canonical_number(row[layout.number_col]) if layout.number_col is not None else ""
        name = as_text(row[layout.name_col]) if layout.name_col is not None else ""
        content = as_text(row[layout.content_col])
        if not any([number, name, content]):
            continue

        same_student = bool(records and number and number == records[-1]["번호"])
        continuation = (not number and not name) or same_student
        if continuation and records:
            if content:
                records[-1]["행동특성 및 종합의견"] = (
                    records[-1]["행동특성 및 종합의견"] + " " + content
                ).strip()
        else:
            records.append(
                {"학년 반": grade_class, "번호": number, "행동특성 및 종합의견": content}
            )

    result = pd.DataFrame(records, columns=columns)
    if result.empty:
        return result
    result = result[result["행동특성 및 종합의견"].str.strip().ne("")]
    return result.drop_duplicates(subset=columns, keep="first").reset_index(drop=True)


def clean_dataframe(df: pd.DataFrame, record_type: RecordType) -> pd.DataFrame:
    """판별된 유형에 맞는 정제 함수를 호출한다."""
    if record_type == "subject":
        return clean_subject_dataframe(df)
    if record_type == "activity_creative":
        return clean_activity_dataframe(df, is_free_semester=False)
    if record_type == "activity_free":
        return clean_activity_dataframe(df, is_free_semester=True)
    return clean_behavior_dataframe(df)


def category_for(record_type: RecordType) -> str:
    return "activity" if record_type.startswith("activity_") else record_type


def process_workbook(
    source: str | Path | BinaryIO | BytesIO,
    source_name: str | None = None,
) -> tuple[list[dict[str, object]], list[dict[str, str]]]:
    """한 엑셀 통합문서의 모든 시트를 자동 판별·정제한다.

    Returns:
        (성공한 결과 목록, 상태/오류 목록). 각 결과에는 source_file, sheet_name,
        record_type, display_type, dataframe 키가 포함된다.
    """
    workbook = pd.ExcelFile(source)
    source_label = source_name or Path(str(source)).name
    processed: list[dict[str, object]] = []
    logs: list[dict[str, str]] = []

    for sheet_name in workbook.sheet_names:
        raw_df = pd.read_excel(workbook, sheet_name=sheet_name, header=None)
        record_type = detect_record_type(raw_df)
        if record_type is None:
            logs.append(
                {
                    "파일": source_label,
                    "시트": str(sheet_name),
                    "판별 결과": "지원하지 않는 양식",
                    "상태": "제목 또는 헤더에서 지원 영역을 확인하지 못했습니다.",
                }
            )
            continue

        result_df = clean_dataframe(raw_df, record_type)
        if result_df.empty:
            logs.append(
                {
                    "파일": source_label,
                    "시트": str(sheet_name),
                    "판별 결과": DISPLAY_NAMES[record_type],
                    "상태": "헤더는 감지했으나 정제 가능한 학생 데이터가 없습니다.",
                }
            )
            continue

        processed.append(
            {
                "source_file": source_label,
                "sheet_name": str(sheet_name),
                "record_type": record_type,
                "display_type": DISPLAY_NAMES[record_type],
                "dataframe": result_df,
            }
        )
        logs.append(
            {
                "파일": source_label,
                "시트": str(sheet_name),
                "판별 결과": DISPLAY_NAMES[record_type],
                "상태": f"정제 완료: {len(result_df):,}행",
            }
        )
    return processed, logs


def save_excel(dataframe: pd.DataFrame, output_path: str | Path) -> None:
    """정제 결과를 인덱스 없이 .xlsx 파일로 저장한다."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="정제결과")


def process_directory(input_dir: str | Path, output_dir: str | Path) -> list[dict[str, str]]:
    """폴더 내 모든 .xlsx 파일을 자동 판별해 개별·영역별 통합 결과로 저장한다."""
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    master_frames: dict[str, list[pd.DataFrame]] = {"subject": [], "activity": [], "behavior": []}
    all_logs: list[dict[str, str]] = []

    for excel_path in sorted(input_path.glob("*.xlsx")):
        if excel_path.name.startswith("~$"):
            continue
        try:
            processed, logs = process_workbook(excel_path)
            all_logs.extend(logs)
            for item in processed:
                record_type = item["record_type"]
                dataframe = item["dataframe"]
                category = category_for(record_type)  # type: ignore[arg-type]
                safe_sheet = re.sub(r"[^0-9A-Za-z가-힣_-]+", "_", str(item["sheet_name"]))
                output_name = f"refined_{excel_path.stem}_{safe_sheet}.xlsx"
                save_excel(dataframe, output_path / output_name)  # type: ignore[arg-type]
                master_frames[category].append(dataframe)  # type: ignore[arg-type]
        except Exception as exc:  # 한 파일 오류가 전체 처리를 중단하지 않도록 한다.
            all_logs.append(
                {
                    "파일": excel_path.name,
                    "시트": "-",
                    "판별 결과": "처리 실패",
                    "상태": f"{type(exc).__name__}: {exc}",
                }
            )

    for category, frames in master_frames.items():
        if frames:
            master = pd.concat(frames, ignore_index=True)
            save_excel(master, output_path / MASTER_FILENAMES[category])

    pd.DataFrame(all_logs, columns=["파일", "시트", "판별 결과", "상태"]).to_csv(
        output_path / "processing_report.csv", index=False, encoding="utf-8-sig"
    )
    return all_logs


def main() -> None:
    parser = argparse.ArgumentParser(description="학교생활기록부 엑셀 자동 판별·정제")
    parser.add_argument(
        "--input-dir",
        default="./input",
        help="원본 .xlsx 파일이 있는 폴더 (기본값: ./input)",
    )
    parser.add_argument(
        "--output-dir",
        default="./refined_all",
        help="정제 파일 저장 폴더 (기본값: ./refined_all)",
    )
    args = parser.parse_args()

    logs = process_directory(args.input_dir, args.output_dir)
    complete_count = sum("정제 완료" in log["상태"] for log in logs)
    print(f"처리 완료: {complete_count}개 시트. 결과 폴더: {Path(args.output_dir).resolve()}")
    for log in logs:
        print(f"[{log['파일']} / {log['시트']}] {log['판별 결과']} — {log['상태']}")


if __name__ == "__main__":
    main()
