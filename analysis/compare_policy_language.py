#!/usr/bin/env python3
"""
Compare district policy PDFs with matching templates.

Behavior:
- Each folder under --policy-root represents one unique policy.
- Policy identity is determined from the folder name or PDF filenames.
- BP and AR are treated as different policy types.
- Templates are matched by policy type and number in the filename.
- If one or more matching templates exist, each district policy is compared
  against them and its closest template is reported.
- If no matching template exists, the district policies are clustered by
  language similarity and the number of clusters is selected automatically.

Example policy keys:
- BP 3510
- BP 3514
- AR 3514.1
"""

from __future__ import annotations

import argparse
import csv
import difflib
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

import numpy as np
from pypdf import PdfReader
from sklearn.cluster import AgglomerativeClustering
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics import silhouette_score
from sklearn.metrics.pairwise import cosine_similarity


START_MARKERS = (
    "the governing board",
    "the board of education",
    "the board recognizes",
    "the board believes",
    "the superintendent or designee",
    "the district shall",
    "it is the policy",
)

STOP_MARKERS = (
    "policy reference disclaimer",
    "board policy references",
    "legal reference:",
    "legal references:",
    "management resources:",
    "cross references",
    "cross reference",
)

WEB_NOISE_PATTERNS = (
    r"https?://\S+",
    r"\bpolicyprintgenerator\b",
    r"\bviewpolicy\.aspx\b",
    r"\bpolicylisting\.aspx\b",
    r"\bprivacy notice\b",
    r"\baccessibility notice\b",
    r"\bcopyright\b",
    r"\bpowered by\b",
    r"\bsearch translate login\b",
)


@dataclass
class PolicyDocument:
    path: Path
    clean_text: str
    comparison_text: str


@dataclass
class ClusterChoice:
    labels: np.ndarray
    cluster_count: int
    silhouette: float | None
    method_note: str


def normalize_policy_type(value: str) -> str:
    upper = value.upper()
    if upper in {"BP", "BOARD POLICY"}:
        return "BP"
    if upper in {"AR", "ADMINISTRATIVE REGULATION"}:
        return "AR"
    raise ValueError(f"Unsupported policy type: {value}")


def parse_policy_key(name: str) -> str | None:
    """
    Extract a policy key from a filename or folder name.

    Supports examples such as:
    - BP 3510
    - BP_3510_Green_Schools.pdf
    - 3510 CSBA BP - Green Schools.pdf
    - Administrative Regulation 3514.1
    """
    text = Path(name).stem.upper()

    type_matches: list[tuple[str, int, int]] = []
    type_pattern = re.compile(
        r"\b(BOARD\s+POLICY|ADMINISTRATIVE\s+REGULATION|BP|AR)\b",
        flags=re.IGNORECASE,
    )
    for match in type_pattern.finditer(text):
        type_matches.append(
            (normalize_policy_type(match.group(1)), match.start(), match.end())
        )

    number_matches: list[tuple[str, int, int]] = []
    for match in re.finditer(r"\b(\d{3,4}(?:\.\d+)?)\b", text):
        number_matches.append((match.group(1), match.start(), match.end()))

    if not type_matches or not number_matches:
        return None

    best_pair: tuple[int, str, str] | None = None
    for policy_type, type_start, type_end in type_matches:
        type_center = (type_start + type_end) // 2
        for number, number_start, number_end in number_matches:
            number_center = (number_start + number_end) // 2
            distance = abs(type_center - number_center)
            candidate = (distance, policy_type, number)
            if best_pair is None or candidate[0] < best_pair[0]:
                best_pair = candidate

    if best_pair is None:
        return None

    _, policy_type, number = best_pair
    return f"{policy_type} {number}"


def safe_slug(value: str) -> str:
    value = value.strip().replace(" ", "_")
    value = re.sub(r"[^A-Za-z0-9_.-]+", "_", value)
    return value.strip("_") or "policy"


def extract_pdf_text(pdf_path: Path) -> str:
    """Extract text from a text-based PDF, preferring layout-aware extraction."""
    try:
        reader = PdfReader(str(pdf_path))
    except Exception as exc:
        raise RuntimeError(f"Could not open PDF: {exc}") from exc

    pages: list[str] = []
    for page_number, page in enumerate(reader.pages, start=1):
        try:
            try:
                text = page.extract_text(extraction_mode="layout") or ""
            except (TypeError, ValueError):
                text = page.extract_text() or ""
        except Exception as exc:
            print(
                f"Warning: failed to extract page {page_number} from "
                f"{pdf_path.name}: {exc}",
                file=sys.stderr,
            )
            text = ""
        pages.append(text)

    combined = "\n".join(pages)
    if len(re.sub(r"\s+", "", combined)) < 100:
        raise RuntimeError(
            "Very little text was extracted. The PDF may be scanned and require OCR."
        )
    return combined


def remove_editorial_notes(text: str) -> str:
    """Remove CSBA editorial Note blocks enclosed by triple asterisks."""
    return re.sub(
        r"\*{3}\s*note\s*:.*?\*{3}",
        " ",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )


def normalize_layout_spacing(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\uf0b7", " ")

    # Layout extraction sometimes separates digits and punctuation.
    text = re.sub(r"(?<=\d)\s+(?=\d)", "", text)
    text = re.sub(r"(?<=\d)\s+(?=[.\-])", "", text)
    text = re.sub(r"(?<=[.\-])\s+(?=\d)", "", text)

    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def remove_web_noise(text: str) -> str:
    cleaned_lines: list[str] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        lower = line.lower()

        if not line:
            cleaned_lines.append("")
            continue

        if any(
            re.search(pattern, lower, flags=re.IGNORECASE)
            for pattern in WEB_NOISE_PATTERNS
        ):
            continue

        # Query-string and navigation fragments are common in downloaded PDFs.
        if line.count("=") >= 3 or line.count("/") >= 5:
            continue

        if lower in {"home", "search", "translate", "login"}:
            continue

        cleaned_lines.append(line)

    return "\n".join(cleaned_lines)


def fuzzy_phrase_pattern(phrase: str) -> re.Pattern[str]:
    """Match a heading even when PDF layout extraction inserts spaces inside words."""
    words = re.findall(r"[A-Za-z]+", phrase)
    word_patterns = [r"\s*".join(re.escape(letter) for letter in word) for word in words]
    return re.compile(r"\s+".join(word_patterns), flags=re.IGNORECASE)


def isolate_policy_body(text: str) -> str:
    """Keep substantive policy language and exclude later reference sections."""
    lower = text.lower()

    start_index = 0
    starts = [lower.find(marker) for marker in START_MARKERS]
    starts = [index for index in starts if index >= 0]
    if starts:
        start_index = min(starts)

    end_index = len(text)
    for marker in STOP_MARKERS:
        match = fuzzy_phrase_pattern(marker).search(text, pos=start_index)
        if match is not None:
            end_index = min(end_index, match.start())

    return text[start_index:end_index].strip()


def clean_policy_text(raw_text: str) -> str:
    text = normalize_layout_spacing(raw_text)
    text = remove_editorial_notes(text)
    text = remove_web_noise(text)
    text = isolate_policy_body(text)
    text = normalize_layout_spacing(text)

    # Remove common standalone page numbers.
    text = re.sub(r"^\s*page\s+\d+(?:\s+of\s+\d+)?\s*$", "", text, flags=re.I | re.M)
    text = re.sub(r"^\s*\d{1,3}\s*$", "", text, flags=re.M)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def normalize_for_comparison(text: str) -> str:
    """Normalize formatting while retaining meaningful wording differences."""
    text = text.lower()
    text = re.sub(r"\bpage\s+\d+(?:\s+of\s+\d+)?\b", " ", text)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def sentence_lines(text: str) -> list[str]:
    """Create moderately readable lines for text-difference files."""
    compact = re.sub(r"\s+", " ", text).strip()
    if not compact:
        return []

    pieces = re.split(
        r"(?<=[.!?])\s+(?=(?:[A-Z]|\d+[.)]|[a-z][.)]))",
        compact,
    )
    return [piece.strip() for piece in pieces if piece.strip()]


def make_diff(template_text: str, district_text: str) -> str:
    diff = difflib.unified_diff(
        sentence_lines(template_text),
        sentence_lines(district_text),
        fromfile="TEMPLATE",
        tofile="DISTRICT_POLICY",
        lineterm="",
        n=1,
    )
    return "\n".join(diff)


def discover_policy_folders(policy_root: Path) -> list[Path]:
    """Find every directory that directly contains at least one PDF."""
    folders: list[Path] = []

    if any(policy_root.glob("*.pdf")):
        folders.append(policy_root)

    for folder in sorted(path for path in policy_root.rglob("*") if path.is_dir()):
        if any(folder.glob("*.pdf")):
            folders.append(folder)

    # Preserve order while removing duplicates.
    seen: set[Path] = set()
    unique: list[Path] = []
    for folder in folders:
        resolved = folder.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(folder)
    return unique


def determine_folder_policy_key(folder: Path, pdfs: Sequence[Path]) -> str:
    folder_key = parse_policy_key(folder.name)
    file_keys = {key for pdf in pdfs if (key := parse_policy_key(pdf.name))}

    if folder_key:
        conflicts = sorted(key for key in file_keys if key != folder_key)
        if conflicts:
            raise RuntimeError(
                f"Folder appears to be {folder_key}, but filenames also contain: "
                + ", ".join(conflicts)
            )
        return folder_key

    if len(file_keys) == 1:
        return next(iter(file_keys))

    if not file_keys:
        raise RuntimeError(
            "Could not identify a policy type and number from the folder name "
            "or PDF filenames. Include a name such as 'BP 3510' or 'AR 3514.1'."
        )

    raise RuntimeError(
        "Multiple policy keys were found in one folder: " + ", ".join(sorted(file_keys))
    )


def index_templates(template_folder: Path) -> tuple[dict[str, list[Path]], list[str]]:
    index: dict[str, list[Path]] = {}
    warnings: list[str] = []

    for pdf in sorted(template_folder.rglob("*.pdf")):
        key = parse_policy_key(pdf.name)
        if key is None:
            warnings.append(
                f"Ignored template with no recognizable BP/AR policy key: {pdf.name}"
            )
            continue
        index.setdefault(key, []).append(pdf)

    return index, warnings


def load_documents(paths: Sequence[Path]) -> tuple[list[PolicyDocument], list[dict[str, str]]]:
    documents: list[PolicyDocument] = []
    errors: list[dict[str, str]] = []

    for path in paths:
        try:
            raw = extract_pdf_text(path)
            clean = clean_policy_text(raw)
            comparison = normalize_for_comparison(clean)
            if len(comparison.split()) < 20:
                raise RuntimeError("Too little usable policy language remained after cleaning.")
            documents.append(
                PolicyDocument(path=path, clean_text=clean, comparison_text=comparison)
            )
        except Exception as exc:
            errors.append({"file": path.name, "error": str(exc)})

    return documents, errors


def build_tfidf(texts: Sequence[str]):
    vectorizer = TfidfVectorizer(
        ngram_range=(1, 2),
        lowercase=False,
        sublinear_tf=True,
        max_features=75000,
    )
    return vectorizer.fit_transform(texts)


def write_csv(path: Path, fieldnames: Sequence[str], rows: Iterable[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def analyze_with_templates(
    policy_key: str,
    district_docs: Sequence[PolicyDocument],
    template_docs: Sequence[PolicyDocument],
    output_dir: Path,
) -> dict[str, object]:
    all_docs = list(district_docs) + list(template_docs)
    matrix = build_tfidf([doc.comparison_text for doc in all_docs])

    district_matrix = matrix[: len(district_docs)]
    template_matrix = matrix[len(district_docs) :]
    similarities = cosine_similarity(district_matrix, template_matrix)

    comparison_rows: list[dict[str, object]] = []
    best_rows: list[dict[str, object]] = []
    diff_dir = output_dir / "diffs"
    diff_dir.mkdir(parents=True, exist_ok=True)

    for district_index, district_doc in enumerate(district_docs):
        district_scores = similarities[district_index]
        best_template_index = int(np.argmax(district_scores))
        best_template = template_docs[best_template_index]
        best_score = float(district_scores[best_template_index])

        exact_match = (
            district_doc.comparison_text == best_template.comparison_text
        )
        adoption_type = "Verbatim" if exact_match else "Modified"

        for template_index, template_doc in enumerate(template_docs):
            comparison_rows.append(
                {
                    "policy_key": policy_key,
                    "district_policy_file": district_doc.path.name,
                    "template_file": template_doc.path.name,
                    "similarity_percent": round(
                        float(district_scores[template_index]) * 100, 2
                    ),
                    "exact_normalized_match": (
                        district_doc.comparison_text == template_doc.comparison_text
                    ),
                }
            )

        diff_name = (
            f"{safe_slug(district_doc.path.stem)}__vs__"
            f"{safe_slug(best_template.path.stem)}.txt"
        )
        diff_path = diff_dir / diff_name
        diff_path.write_text(
            make_diff(best_template.clean_text, district_doc.clean_text),
            encoding="utf-8",
        )

        best_rows.append(
            {
                "policy_key": policy_key,
                "district_policy_file": district_doc.path.name,
                "closest_template_file": best_template.path.name,
                "similarity_percent": round(best_score * 100, 2),
                "adoption_type": adoption_type,
                "district_word_count": len(district_doc.comparison_text.split()),
                "template_word_count": len(best_template.comparison_text.split()),
                "diff_file": str(diff_path.relative_to(output_dir)),
            }
        )

    write_csv(
        output_dir / "template_comparisons.csv",
        (
            "policy_key",
            "district_policy_file",
            "template_file",
            "similarity_percent",
            "exact_normalized_match",
        ),
        comparison_rows,
    )
    write_csv(
        output_dir / "best_template_matches.csv",
        (
            "policy_key",
            "district_policy_file",
            "closest_template_file",
            "similarity_percent",
            "adoption_type",
            "district_word_count",
            "template_word_count",
            "diff_file",
        ),
        best_rows,
    )

    return {
        "mode": "template comparison",
        "policy_count": len(district_docs),
        "template_count": len(template_docs),
        "cluster_count": "",
    }


def agglomerative_labels(distance_matrix: np.ndarray, cluster_count: int) -> np.ndarray:
    try:
        model = AgglomerativeClustering(
            n_clusters=cluster_count,
            metric="precomputed",
            linkage="average",
        )
    except TypeError:  # Compatibility with older scikit-learn versions.
        model = AgglomerativeClustering(
            n_clusters=cluster_count,
            affinity="precomputed",
            linkage="average",
        )
    return model.fit_predict(distance_matrix)


def choose_clusters(similarity_matrix: np.ndarray) -> ClusterChoice:
    """
    Select a cluster count from the similarity structure.

    - One document always forms one cluster.
    - Two documents form one cluster when strongly similar, otherwise two.
    - For three or more documents, uniformly similar collections remain one
      cluster. Otherwise, candidate counts are evaluated with silhouette score.
    """
    count = similarity_matrix.shape[0]

    if count == 1:
        return ClusterChoice(
            labels=np.array([0]),
            cluster_count=1,
            silhouette=None,
            method_note="Only one policy PDF was available.",
        )

    upper = similarity_matrix[np.triu_indices(count, k=1)]
    mean_similarity = float(np.mean(upper)) if upper.size else 1.0
    tenth_percentile = float(np.percentile(upper, 10)) if upper.size else 1.0

    if count == 2:
        if float(upper[0]) >= 0.80:
            labels = np.array([0, 0])
            return ClusterChoice(
                labels=labels,
                cluster_count=1,
                silhouette=None,
                method_note="The two policies had at least 80% cosine similarity.",
            )
        labels = np.array([0, 1])
        return ClusterChoice(
            labels=labels,
            cluster_count=2,
            silhouette=None,
            method_note="The two policies had less than 80% cosine similarity.",
        )

    # Avoid inventing clusters when the whole collection is already highly similar.
    if mean_similarity >= 0.90 and tenth_percentile >= 0.80:
        return ClusterChoice(
            labels=np.zeros(count, dtype=int),
            cluster_count=1,
            silhouette=None,
            method_note=(
                "All policies were treated as one cluster because the collection "
                "was uniformly highly similar."
            ),
        )

    distance_matrix = np.clip(1.0 - similarity_matrix, 0.0, 1.0)
    np.fill_diagonal(distance_matrix, 0.0)

    best_labels: np.ndarray | None = None
    best_count: int | None = None
    best_score = float("-inf")

    max_clusters = min(10, count - 1)
    for cluster_count in range(2, max_clusters + 1):
        labels = agglomerative_labels(distance_matrix, cluster_count)
        if len(set(labels)) < 2:
            continue
        score = float(
            silhouette_score(distance_matrix, labels, metric="precomputed")
        )
        if score > best_score:
            best_score = score
            best_labels = labels
            best_count = cluster_count

    if best_labels is None or best_count is None or best_score < 0.05:
        return ClusterChoice(
            labels=np.zeros(count, dtype=int),
            cluster_count=1,
            silhouette=None if best_score == float("-inf") else best_score,
            method_note=(
                "No multi-cluster solution had a sufficiently clear silhouette, "
                "so the policies were kept in one cluster."
            ),
        )

    return ClusterChoice(
        labels=best_labels,
        cluster_count=best_count,
        silhouette=best_score,
        method_note=(
            "The cluster count was selected automatically by maximizing the "
            "silhouette score for average-linkage hierarchical clustering."
        ),
    )


def remap_cluster_labels(labels: np.ndarray, filenames: Sequence[str]) -> np.ndarray:
    """Renumber clusters from 1, ordered by size and then representative name."""
    unique = sorted(set(int(value) for value in labels))
    sortable: list[tuple[int, str, int]] = []

    for old_label in unique:
        members = [
            filenames[index]
            for index, value in enumerate(labels)
            if int(value) == old_label
        ]
        sortable.append((-len(members), min(members).lower(), old_label))

    mapping = {
        old_label: new_label
        for new_label, (_, _, old_label) in enumerate(sorted(sortable), start=1)
    }
    return np.array([mapping[int(value)] for value in labels], dtype=int)


def analyze_without_template(
    policy_key: str,
    district_docs: Sequence[PolicyDocument],
    output_dir: Path,
) -> dict[str, object]:
    matrix = build_tfidf([doc.comparison_text for doc in district_docs])
    similarity = cosine_similarity(matrix)

    choice = choose_clusters(similarity)
    filenames = [doc.path.name for doc in district_docs]
    labels = remap_cluster_labels(choice.labels, filenames)

    assignment_rows: list[dict[str, object]] = []
    for index, doc in enumerate(district_docs):
        cluster_id = int(labels[index])
        member_indices = np.where(labels == cluster_id)[0]

        if len(member_indices) == 1:
            average_in_cluster = 100.0
        else:
            other_indices = [value for value in member_indices if value != index]
            average_in_cluster = round(
                float(np.mean(similarity[index, other_indices])) * 100,
                2,
            )

        cluster_means: dict[int, float] = {}
        for member_index in member_indices:
            others = [value for value in member_indices if value != member_index]
            if not others:
                cluster_means[int(member_index)] = 1.0
            else:
                cluster_means[int(member_index)] = float(
                    np.mean(similarity[member_index, others])
                )
        best_representative_score = max(cluster_means.values())
        representative_candidates = [
            item
            for item, score in cluster_means.items()
            if abs(score - best_representative_score) < 1e-12
        ]
        representative_index = min(
            representative_candidates,
            key=lambda item: filenames[item].lower(),
        )

        assignment_rows.append(
            {
                "policy_key": policy_key,
                "policy_file": doc.path.name,
                "cluster_id": cluster_id,
                "cluster_size": len(member_indices),
                "representative_policy": index == representative_index,
                "average_similarity_within_cluster_percent": average_in_cluster,
                "word_count": len(doc.comparison_text.split()),
            }
        )

    write_csv(
        output_dir / "cluster_assignments.csv",
        (
            "policy_key",
            "policy_file",
            "cluster_id",
            "cluster_size",
            "representative_policy",
            "average_similarity_within_cluster_percent",
            "word_count",
        ),
        assignment_rows,
    )

    matrix_path = output_dir / "similarity_matrix.csv"
    with matrix_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(["policy_file", *filenames])
        for row_index, filename in enumerate(filenames):
            writer.writerow(
                [filename]
                + [round(float(value) * 100, 2) for value in similarity[row_index]]
            )

    upper = similarity[np.triu_indices(len(district_docs), k=1)]
    average_pairwise = (
        round(float(np.mean(upper)) * 100, 2) if upper.size else 100.0
    )

    summary_lines = [
        f"Policy: {policy_key}",
        "Mode: automatic clustering because no matching template was found",
        f"Policy PDFs analyzed: {len(district_docs)}",
        f"Clusters selected: {choice.cluster_count}",
        f"Average pairwise similarity: {average_pairwise}%",
        (
            "Silhouette score: not applicable"
            if choice.silhouette is None
            else f"Silhouette score: {choice.silhouette:.4f}"
        ),
        f"Method: {choice.method_note}",
    ]
    (output_dir / "cluster_summary.txt").write_text(
        "\n".join(summary_lines) + "\n",
        encoding="utf-8",
    )

    return {
        "mode": "automatic clustering",
        "policy_count": len(district_docs),
        "template_count": 0,
        "cluster_count": choice.cluster_count,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Compare each policy folder with a matching BP/AR template, or "
            "automatically cluster the policies when no template exists."
        )
    )
    parser.add_argument(
        "--policy-root",
        type=Path,
        required=True,
        help=(
            "Root folder containing one folder per policy. PDFs placed directly "
            "in this root are also treated as one policy folder."
        ),
    )
    parser.add_argument(
        "--template-folder",
        type=Path,
        required=True,
        help="Folder containing available BP and AR template PDFs.",
    )
    parser.add_argument(
        "--output-folder",
        type=Path,
        default=Path("output"),
        help="Folder where results will be written.",
    )
    args = parser.parse_args()

    if not args.policy_root.exists():
        print(f"Error: policy root does not exist: {args.policy_root}", file=sys.stderr)
        return 1
    if not args.template_folder.exists():
        print(
            f"Error: template folder does not exist: {args.template_folder}",
            file=sys.stderr,
        )
        return 1

    args.output_folder.mkdir(parents=True, exist_ok=True)

    template_index, template_warnings = index_templates(args.template_folder)
    for warning in template_warnings:
        print(f"Warning: {warning}", file=sys.stderr)

    folders = discover_policy_folders(args.policy_root)
    if not folders:
        print(
            f"Error: no folders containing PDFs were found under {args.policy_root}",
            file=sys.stderr,
        )
        return 1

    summary_rows: list[dict[str, object]] = []

    for folder in folders:
        pdfs = sorted(folder.glob("*.pdf"))
        relative_folder = (
            "." if folder == args.policy_root else str(folder.relative_to(args.policy_root))
        )

        try:
            policy_key = determine_folder_policy_key(folder, pdfs)
        except Exception as exc:
            print(f"Skipping {relative_folder}: {exc}", file=sys.stderr)
            summary_rows.append(
                {
                    "policy_folder": relative_folder,
                    "policy_key": "",
                    "mode": "error",
                    "district_pdfs_analyzed": 0,
                    "matching_templates": 0,
                    "clusters": "",
                    "errors": str(exc),
                }
            )
            continue

        output_dir = args.output_folder / safe_slug(policy_key)
        output_dir.mkdir(parents=True, exist_ok=True)

        print(f"Processing {policy_key}: {len(pdfs)} district PDF(s)")
        district_docs, district_errors = load_documents(pdfs)

        if district_errors:
            write_csv(
                output_dir / "extraction_errors.csv",
                ("file", "error"),
                district_errors,
            )

        if not district_docs:
            error_message = "No district PDFs could be processed."
            summary_rows.append(
                {
                    "policy_folder": relative_folder,
                    "policy_key": policy_key,
                    "mode": "error",
                    "district_pdfs_analyzed": 0,
                    "matching_templates": len(template_index.get(policy_key, [])),
                    "clusters": "",
                    "errors": error_message,
                }
            )
            continue

        template_paths = template_index.get(policy_key, [])
        template_docs: list[PolicyDocument] = []
        template_errors: list[dict[str, str]] = []
        if template_paths:
            template_docs, template_errors = load_documents(template_paths)
            if template_errors:
                write_csv(
                    output_dir / "template_extraction_errors.csv",
                    ("file", "error"),
                    template_errors,
                )

        try:
            if template_docs:
                result = analyze_with_templates(
                    policy_key,
                    district_docs,
                    template_docs,
                    output_dir,
                )
            else:
                result = analyze_without_template(
                    policy_key,
                    district_docs,
                    output_dir,
                )

            summary_rows.append(
                {
                    "policy_folder": relative_folder,
                    "policy_key": policy_key,
                    "mode": result["mode"],
                    "district_pdfs_analyzed": result["policy_count"],
                    "matching_templates": result["template_count"],
                    "clusters": result["cluster_count"],
                    "errors": len(district_errors) + len(template_errors),
                }
            )
        except Exception as exc:
            print(f"Error processing {policy_key}: {exc}", file=sys.stderr)
            summary_rows.append(
                {
                    "policy_folder": relative_folder,
                    "policy_key": policy_key,
                    "mode": "error",
                    "district_pdfs_analyzed": len(district_docs),
                    "matching_templates": len(template_docs),
                    "clusters": "",
                    "errors": str(exc),
                }
            )

    write_csv(
        args.output_folder / "run_summary.csv",
        (
            "policy_folder",
            "policy_key",
            "mode",
            "district_pdfs_analyzed",
            "matching_templates",
            "clusters",
            "errors",
        ),
        summary_rows,
    )

    print(f"Done. Results saved to: {args.output_folder}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
