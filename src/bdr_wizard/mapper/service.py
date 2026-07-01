from __future__ import annotations

import json
from difflib import SequenceMatcher
from pathlib import Path
from uuid import uuid4

from pydantic import TypeAdapter

from bdr_wizard.analyzer.service import normalize_text
from bdr_wizard.localization import TARGET_FIELDS
from bdr_wizard.models import ColumnCandidate, FileRole, MappingRule, WorkbookProfile


MAX_STORED_RULES = 500


class MappingRuleStore:
    def __init__(self, path: Path = Path("data/mapping_rules.json")) -> None:
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def load(self) -> list[MappingRule]:
        if not self.path.exists():
            return []
        return TypeAdapter(list[MappingRule]).validate_python(json.loads(self.path.read_text("utf-8")))

    def save(self, rules: list[MappingRule]) -> None:
        payload = [
            rule.model_dump(mode="json")
            for rule in rules
            if rule.confirmed and rule.source_signature
        ][:MAX_STORED_RULES]
        temporary_path = self.path.with_suffix(f".{uuid4().hex}.tmp")
        temporary_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        temporary_path.replace(self.path)


class MappingEngine:
    def __init__(self, rule_store: MappingRuleStore | None = None) -> None:
        self.rule_store = rule_store or MappingRuleStore()

    def suggest(self, profiles: list[WorkbookProfile]) -> tuple[list[MappingRule], list[ColumnCandidate]]:
        candidates = self._build_candidates(profiles)
        saved_rules = self.rule_store.load()
        suggested_rules: list[MappingRule] = []

        for target_field in TARGET_FIELDS:
            saved = self._find_saved_rule(target_field, saved_rules, profiles)
            if saved is not None:
                suggested_rules.append(saved.model_copy(update={"confirmed": False}))
                continue

            field_candidates = [candidate for candidate in candidates if candidate.target_field == target_field]
            best = max(field_candidates, key=lambda item: item.score, default=None)
            if best is None or best.score < 0.45:
                suggested_rules.append(MappingRule(target_field=target_field))
            else:
                suggested_rules.append(
                    MappingRule(
                        target_field=target_field,
                        source_file_id=best.source_file_id,
                        source_sheet=best.source_sheet,
                        source_header=best.source_header,
                        source_signature=self.source_signature(
                            profiles,
                            best.source_file_id,
                            best.source_sheet,
                            best.source_header,
                        ),
                    )
                )

        return suggested_rules, candidates

    def _build_candidates(self, profiles: list[WorkbookProfile]) -> list[ColumnCandidate]:
        candidates: list[ColumnCandidate] = []
        for profile in profiles:
            role_weight = self._role_weight(profile.detected_role)
            for sheet in profile.sheets:
                sheet_weight = min(0.2, sheet.data_density)
                for header in sheet.headers:
                    normalized_header = normalize_text(header)
                    for target_field, synonyms in TARGET_FIELDS.items():
                        score, reason = self._score_header(normalized_header, synonyms)
                        if score <= 0:
                            continue
                        candidates.append(
                            ColumnCandidate(
                                source_file_id=profile.file_id,
                                source_sheet=sheet.name,
                                source_header=header,
                                target_field=target_field,
                                score=round(min(1.0, score + role_weight + sheet_weight), 3),
                                reason=reason,
                            )
                        )
        return candidates

    def _score_header(self, header: str, synonyms: list[str]) -> tuple[float, str]:
        best_score = 0.0
        best_synonym = ""
        for synonym in synonyms:
            normalized_synonym = normalize_text(synonym)
            if not normalized_synonym:
                continue
            if normalized_synonym == header:
                return 0.82, f"точное совпадение с «{synonym}»"
            if normalized_synonym in header or header in normalized_synonym:
                score = 0.7
            else:
                score = SequenceMatcher(None, header, normalized_synonym).ratio() * 0.68
            if score > best_score:
                best_score = score
                best_synonym = synonym
        if best_score < 0.36:
            return 0.0, ""
        return best_score, f"похоже на «{best_synonym}»"

    def _find_saved_rule(
        self,
        target_field: str,
        saved_rules: list[MappingRule],
        profiles: list[WorkbookProfile],
    ) -> MappingRule | None:
        available_by_signature = {
            self._source_signature_from_parts(profile.fingerprint.structure_signature, sheet.name, header): (
                profile.file_id,
                sheet.name,
                header,
            )
            for profile in profiles
            for sheet in profile.sheets
            for header in sheet.headers
        }
        for rule in saved_rules:
            if rule.target_field != target_field or not rule.source_signature:
                continue
            match = available_by_signature.get(rule.source_signature)
            if match is not None:
                source_file_id, source_sheet, source_header = match
                return rule.model_copy(
                    update={
                        "source_file_id": source_file_id,
                        "source_sheet": source_sheet,
                        "source_header": source_header,
                    }
                )
        return None

    def source_signature(
        self,
        profiles: list[WorkbookProfile],
        source_file_id: str,
        source_sheet: str,
        source_header: str,
    ) -> str:
        profile = next(item for item in profiles if item.file_id == source_file_id)
        return self._source_signature_from_parts(
            profile.fingerprint.structure_signature,
            source_sheet,
            source_header,
        )

    def _source_signature_from_parts(
        self,
        workbook_signature: str,
        source_sheet: str,
        source_header: str,
    ) -> str:
        return "|".join(
            [
                workbook_signature,
                normalize_text(source_sheet),
                normalize_text(source_header),
            ]
        )

    def _role_weight(self, role: FileRole) -> float:
        if role == FileRole.MAIN:
            return 0.12
        if role == FileRole.MAPPING:
            return 0.08
        if role == FileRole.DIRECTORY:
            return 0.04
        return 0.0
