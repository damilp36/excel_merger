import pandas as pd

from excel_merger.merge import LookupSpec, MatchOptions, canonicalize_key, merge_lookups


def make_inputs() -> tuple[pd.DataFrame, LookupSpec]:
    base = pd.DataFrame(
        {
            "Customer ID": [1, " 2 ", "0042", None],
            "Customer": ["Ada", "Ben", "Cy", "Dee"],
        }
    )
    lookup = pd.DataFrame(
        {
            "Account": ["1", 2, 42, "2"],
            "Region": ["West", "East", "Central", "Replacement"],
            "Score": [10, 20, 30, 21],
        }
    )
    spec = LookupSpec(
        file_name="accounts.xlsx",
        sheet_name="Data",
        dataframe=lookup,
        key_column="Account",
        return_columns=["Region", "Score"],
        prefix="Accounts",
    )
    return base, spec


def test_first_match_behaves_like_vlookup_and_audits_duplicates() -> None:
    base, spec = make_inputs()
    result = merge_lookups(base, "Customer ID", [spec], MatchOptions())

    assert len(result.dataframe) == 4
    assert result.dataframe["Accounts.Region"].tolist()[:2] == ["West", "East"]
    assert pd.isna(result.dataframe.loc[2, "Accounts.Region"])
    assert result.dataframe["Accounts.Match status"].tolist() == [
        "Matched",
        "Matched",
        "Not matched",
        "Missing key",
    ]
    audit = result.audits[0]
    assert audit.duplicate_key_rows == 2
    assert audit.matched_base_rows == 2
    assert audit.unmatched_base_rows == 1
    assert audit.missing_base_keys == 1


def test_expand_strategy_returns_all_matches() -> None:
    base, spec = make_inputs()
    result = merge_lookups(
        base,
        "Customer ID",
        [spec],
        MatchOptions(duplicate_strategy="expand"),
    )

    assert len(result.dataframe) == 5
    ben_rows = result.dataframe[result.dataframe["Customer"] == "Ben"]
    assert ben_rows["Accounts.Region"].tolist() == ["East", "Replacement"]


def test_key_normalization_preserves_zero_padded_ids() -> None:
    options = MatchOptions()
    assert canonicalize_key(42, options) == canonicalize_key("42", options)
    assert canonicalize_key("  Acct-7 ", options) == canonicalize_key("acct-7", options)
    assert canonicalize_key("0042", options) != canonicalize_key(42, options)
    assert canonicalize_key("  ", options) is None


def test_base_return_columns_keep_the_key_and_selected_fields_only() -> None:
    base = pd.DataFrame(
        {
            "MATRIC NUMBER": ["A-1", "A-2"],
            "CA SCORE": [18, 16],
            "Private note": ["review", "ok"],
        }
    )
    lookup = LookupSpec(
        file_name="exam.xlsx",
        sheet_name="Sheet1",
        dataframe=pd.DataFrame(
            {"MATRIC NUMBER": ["a-1", "a-2"], "EXAM SCORE": [62, 55]}
        ),
        key_column="MATRIC NUMBER",
        return_columns=["EXAM SCORE"],
        prefix="Exam",
    )

    result = merge_lookups(
        base,
        "MATRIC NUMBER",
        [lookup],
        base_return_columns=["CA SCORE"],
    )

    assert list(result.dataframe.columns) == [
        "MATRIC NUMBER",
        "CA SCORE",
        "Exam.EXAM SCORE",
        "Exam.Match status",
    ]
    assert result.dataframe["Exam.EXAM SCORE"].tolist() == [62, 55]
