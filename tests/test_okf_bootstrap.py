"""
Python twin of OKFBootstrap.bas -- create-if-absent bootstrap policy.
"""


def bootstrap_plan(seed_paths: list[str], existing_paths: set[str]) -> list[str]:
    """Return the seed paths that should be written (those not already present)."""
    return [p for p in seed_paths if p not in existing_paths]


SEED_PATHS = [
    "_foundation/00-operating-profile.md",
    "builds/example-build.md",
    "skills/skill-md-authoring.md",
]


def test_empty_bundle_returns_all_seeds() -> None:
    result = bootstrap_plan(SEED_PATHS, set())
    assert result == SEED_PATHS


def test_partial_bundle_returns_missing_only() -> None:
    existing = {"skills/skill-md-authoring.md"}
    result = bootstrap_plan(SEED_PATHS, existing)
    assert result == [
        "_foundation/00-operating-profile.md",
        "builds/example-build.md",
    ]


def test_full_bundle_returns_empty() -> None:
    existing = set(SEED_PATHS)
    result = bootstrap_plan(SEED_PATHS, existing)
    assert result == []
