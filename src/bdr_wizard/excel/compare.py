from __future__ import annotations

import json
from difflib import SequenceMatcher

from bdr_wizard.models import CompareReport, ImportDecision, Severity, WorkbookProfile


class CompareEngine:
    def compare(self, profiles: list[WorkbookProfile]) -> CompareReport:
        decisions: list[ImportDecision] = []
        similar_pairs: list[dict[str, object]] = []
        sheet_differences: list[dict[str, object]] = []

        for left_index, left in enumerate(profiles):
            for right in profiles[left_index + 1 :]:
                similarity = self.profile_similarity(left, right)
                formula_delta = abs(
                    left.fingerprint.total_formula_cells
                    - right.fingerprint.total_formula_cells
                )
                sheet_delta = abs(left.fingerprint.sheet_count - right.fingerprint.sheet_count)
                sheet_report = self._sheet_difference(left, right)
                if sheet_report:
                    sheet_differences.append(sheet_report)
                if similarity >= 0.72:
                    pair = {
                        "left_file_id": left.file_id,
                        "right_file_id": right.file_id,
                        "left_file_name": left.original_name,
                        "right_file_name": right.original_name,
                        "similarity": similarity,
                        "sheet_delta": sheet_delta,
                        "formula_delta": formula_delta,
                    }
                    similar_pairs.append(pair)
                    decisions.append(
                        ImportDecision(
                            severity=Severity.INFO,
                            code="похожие_файлы",
                            message=(
                                f"Файлы «{left.original_name}» и «{right.original_name}» похожи "
                                f"по структуре на {round(similarity * 100)}%."
                            ),
                            context=pair,
                        )
                    )

        return CompareReport(
            decisions=decisions,
            similar_pairs=similar_pairs,
            sheet_differences=sheet_differences,
        )

    def profile_similarity(self, left: WorkbookProfile, right: WorkbookProfile) -> float:
        left_headers = " ".join(left.fingerprint.header_tokens)
        right_headers = " ".join(right.fingerprint.header_tokens)
        header_similarity = SequenceMatcher(None, left_headers, right_headers).ratio()
        formula_similarity = SequenceMatcher(
            None,
            " ".join(left.fingerprint.formula_tokens),
            " ".join(right.fingerprint.formula_tokens),
        ).ratio()
        left_ranges = json.dumps(
            [(sheet.max_row, sheet.max_column, sheet.hidden, sheet.empty) for sheet in left.sheets],
            ensure_ascii=False,
        )
        right_ranges = json.dumps(
            [(sheet.max_row, sheet.max_column, sheet.hidden, sheet.empty) for sheet in right.sheets],
            ensure_ascii=False,
        )
        range_similarity = SequenceMatcher(None, left_ranges, right_ranges).ratio()
        sheet_delta = abs(left.fingerprint.sheet_count - right.fingerprint.sheet_count)
        sheet_similarity = 1 / (1 + sheet_delta)
        return round(
            header_similarity * 0.4
            + formula_similarity * 0.25
            + range_similarity * 0.2
            + sheet_similarity * 0.15,
            3,
        )

    def _sheet_difference(
        self,
        left: WorkbookProfile,
        right: WorkbookProfile,
    ) -> dict[str, object] | None:
        left_sheets = {sheet.name: sheet for sheet in left.sheets}
        right_sheets = {sheet.name: sheet for sheet in right.sheets}
        missing_left = sorted(set(right_sheets) - set(left_sheets))
        missing_right = sorted(set(left_sheets) - set(right_sheets))
        common = sorted(set(left_sheets) & set(right_sheets))
        changed = []
        for sheet_name in common:
            left_sheet = left_sheets[sheet_name]
            right_sheet = right_sheets[sheet_name]
            if (
                left_sheet.max_row,
                left_sheet.max_column,
                left_sheet.formula_patterns,
                left_sheet.headers,
            ) != (
                right_sheet.max_row,
                right_sheet.max_column,
                right_sheet.formula_patterns,
                right_sheet.headers,
            ):
                changed.append(
                    {
                        "sheet": sheet_name,
                        "left_size": [left_sheet.max_row, left_sheet.max_column],
                        "right_size": [right_sheet.max_row, right_sheet.max_column],
                        "header_similarity": SequenceMatcher(
                            None,
                            " ".join(left_sheet.headers),
                            " ".join(right_sheet.headers),
                        ).ratio(),
                        "formula_similarity": SequenceMatcher(
                            None,
                            " ".join(left_sheet.formula_patterns),
                            " ".join(right_sheet.formula_patterns),
                        ).ratio(),
                    }
                )
        if not missing_left and not missing_right and not changed:
            return None
        return {
            "left_file_id": left.file_id,
            "right_file_id": right.file_id,
            "missing_in_left": missing_left,
            "missing_in_right": missing_right,
            "changed_sheets": changed,
        }


def compare_workbooks(profiles: list[WorkbookProfile]) -> list[ImportDecision]:
    return CompareEngine().compare(profiles).decisions
