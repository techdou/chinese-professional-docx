# Chinese Professional DOCX Skill · v1.7.0

General-purpose Chinese academic and professional Word authoring and repair. The default is a conservative black-on-white scholarly style: two-character first-line indentation, native multilevel headings, three-line tables, lightly shaded code blocks, and editable cross-references. Supplied school, publisher, or institutional templates always take precedence.

**New in v1.7:** GB/T 7714—2025 core bibliography formatting for J/M/C/D/R/S/EB/DS/PP, numeric superscripts (including numeric spans and source page locators), author-year citations, and edit-safe OOXML `SEQ`/`REF` links. All bibliographic metadata must be source-verified. Preformatted legacy references remain allowed but are flagged as unverified; full standard and publisher compliance requires human review.

Run `python scripts/docx_ops.py --help` in the standalone skill package and `python -m unittest discover -s tests -v` for automated tests. See [Chinese documentation](README.md), [SKILL.md](SKILL.md) and [reference rules](references/gbt7714-2025.md).

**Original font files are not included in the source package**. Prepare them locally (e.g. from a licensed Windows install) and keep font binaries out of the repository; `tools/merge_full_original.py` merges a local full archive with upgraded sources. Windows Word and WPS native renderer verification remains environment-dependent; do not claim it has passed based only on LibreOffice.
