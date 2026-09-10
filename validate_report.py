#!/usr/bin/env python3
"""
Sanity-checks a generated report .pptx before the monthly routine is allowed
to upload it to Drive. Exits 0 (pass) or 1 (fail).

Checks:
  - Every hero stat card (the label/value/delta 3-paragraph rounded-rectangle
    shapes built by add_hero_cards() in generate_report.py) has a non-empty
    value. TikTok's "Current Followers" card is an intentional manual
    placeholder ("[Count]") and is explicitly exempted, not a failure.
  - Every chart has at least one non-zero data point across all its series.

On failure, appends a timestamped entry to run-log.txt (in the current
working directory) describing exactly what failed, and does NOT raise --
the caller decides what to do with the exit code.
"""
import sys
from datetime import datetime, timezone

from pptx import Presentation

MANUAL_PLACEHOLDER_CARDS = {("Current Followers", "[Count]")}


def validate(pptx_path):
    problems = []
    prs = Presentation(pptx_path)

    for i, slide in enumerate(prs.slides, 1):
        title = next(
            (s.text_frame.text for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()),
            f"Slide {i}",
        )

        for shape in slide.shapes:
            # Hero cards: rounded-rectangle autoshapes with the label/value/delta
            # 3-paragraph structure used by add_hero_cards(). Anything else with
            # 3 paragraphs of text is not a hero card and is skipped.
            if shape.has_text_frame and len(shape.text_frame.paragraphs) == 3:
                paras = shape.text_frame.paragraphs
                if not (paras[0].runs and paras[1].runs):
                    continue
                label = paras[0].runs[0].text
                value = paras[1].runs[0].text
                if (label, value) in MANUAL_PLACEHOLDER_CARDS:
                    continue
                if not value or not value.strip():
                    problems.append(f"Slide {i} ({title!r}): hero card {label!r} has an empty value")

            if shape.has_chart:
                series_list = shape.chart.series
                any_nonzero = any(v for series in series_list for v in series.values if v)
                if not any_nonzero:
                    problems.append(f"Slide {i} ({title!r}): chart has no non-zero data points in any series")

    return problems


def main():
    if len(sys.argv) != 2:
        print("Usage: validate_report.py <path-to-pptx>", file=sys.stderr)
        sys.exit(2)

    pptx_path = sys.argv[1]
    problems = validate(pptx_path)

    if problems:
        with open("run-log.txt", "a") as f:
            f.write(f"\n[{datetime.now(timezone.utc).isoformat()}] VALIDATION FAILED for {pptx_path}\n")
            for p in problems:
                f.write(f"  - {p}\n")
        print("VALIDATION FAILED:")
        for p in problems:
            print(f"  - {p}")
        sys.exit(1)

    print(f"VALIDATION PASSED: {pptx_path}")
    sys.exit(0)


if __name__ == "__main__":
    main()
