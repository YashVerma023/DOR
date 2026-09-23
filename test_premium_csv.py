"""The automated premium must be byte-identical to a hand-uploaded file.

    python test_premium_csv.py

That is the whole point of the Automated / Manual switch: one format, one
reader, so switching modes cannot change a number in the report. This checks
the bytes the fetcher hands the app against the bytes it writes to disk, and
that the object wrapping them still looks like an upload to the code that
reads it. No Grafana needed — the rows are fabricated.
"""

import io
import pathlib
import tempfile
from datetime import datetime

import premium_fetcher as pf

ROWS = [
    [datetime(2026, 9, 22, 9, 15, 0), 210],        # integer stays "210"
    [datetime(2026, 9, 22, 9, 15, 5), 210.0],      # float that IS an integer
    [datetime(2026, 9, 22, 15, 39, 55), 187.25],   # genuine decimal
]


def test_csv_text_matches_write_csv():
    """The in-memory bytes and the on-disk file must not drift apart."""
    with tempfile.TemporaryDirectory() as tmp:
        path = pf.write_csv(ROWS, "NIFTY", ROWS[0][0].date(), tmp)
        on_disk = pathlib.Path(path).read_bytes()
    in_memory = pf.csv_text(ROWS).encode("utf-8-sig")
    assert on_disk == in_memory, "fetched bytes differ from the written file"


def test_format_is_the_uploaders():
    text = pf.csv_text(ROWS)
    assert text.startswith("sep=,\r\n"), "the sep= line the exports carry is gone"
    assert '"time","Premium"\r\n' in text, "header no longer matches the exports"
    # "210", never "210.0" — a diff against a hand-made file must come back empty
    assert ",210\r\n" in text and "210.0" not in text
    assert ",187.25\r\n" in text, "decimals must survive"
    assert text.count("\r\n") == len(ROWS) + 2


def test_filename_matches_the_uploads():
    assert (pf.filename("NIFTY", ROWS[0][0].date())
            == "Synthetic Premium NF 22SEP26.csv")
    assert pf.filename("SENSEX", ROWS[0][0].date()).startswith(
        "Synthetic Premium SX ")


def test_fetched_wears_the_uploader_interface():
    """app._Fetched stands in for an upload, so it needs .name + .getvalue()."""
    class _Fetched(io.BytesIO):            # the definition in app.py, verbatim
        def __init__(self, data, name):
            super().__init__(data)
            self.name = name

    data = pf.csv_text(ROWS).encode("utf-8-sig")
    fetched = _Fetched(data, pf.filename("NIFTY", ROWS[0][0].date()))
    assert fetched.getvalue() == data
    assert fetched.name.endswith(".csv")
    # this is exactly what app.py does with an upload
    assert io.BytesIO(fetched.getvalue()).read(6)


def test_window_defaults_to_the_session():
    day = ROWS[0][0].date()
    start, end = pf.window(day)
    assert (start.hour, start.minute) == pf.SESSION_START
    assert (end.hour, end.minute) == pf.SESSION_END


def test_window_narrows_for_testing():
    day = ROWS[0][0].date()
    start, end = pf.window(day, "09:15", "10:15")
    assert (start.hour, start.minute) == (9, 15)
    assert (end.hour, end.minute) == (10, 15)
    # seconds are accepted and ignored, since that is what people paste
    assert pf.window(day, "09:15:00", "10:15:00") == (start, end)
    # one side only still works
    assert pf.window(day, None, "10:15")[0] == pf.window(day)[0]


def test_bad_window_is_refused():
    """A typo must stop the run, not silently fetch the whole session and be
    mistaken for the narrow window that was asked for."""
    for bad in ("10.15", "lunchtime", "10:xx"):
        try:
            pf.window(ROWS[0][0].date(), bad)
        except SystemExit:
            pass
        else:
            raise AssertionError(f"{bad!r} should be refused")


def test_index_bands_do_not_overlap():
    """The bands name the instrument, so an overlap would silently mislabel a
    whole day's premium file — NF holding SENSEX data."""
    import grafana_probe as gp

    spans = sorted(gp.INDEX_BANDS.values())
    for (_, hi), (lo, _) in zip(spans, spans[1:]):
        assert hi <= lo, f"bands overlap at {hi}..{lo}"


def test_bands_cover_the_real_levels():
    """The levels seen live on 23-09-2026, and levels far either side of them,
    so a band that drifts out of date fails here rather than in a report."""
    import grafana_probe as gp

    def band_for(level):
        return next((n for n, (lo, hi) in gp.INDEX_BANDS.items()
                     if lo <= level < hi), None)

    assert band_for(23_400) == "NIFTY"          # live, 23-09-2026
    assert band_for(74_700) == "SENSEX"         # live, 23-09-2026
    assert band_for(51_000) == "BANKNIFTY"
    assert band_for(12_000) == "NIFTY"          # a long way down
    assert band_for(150_000) == "SENSEX"        # a long way up
    assert band_for(5) is None                  # nonsense stays unnamed


def test_missing_index_is_refused():
    """No instance running the asked-for index must fail loudly, not return an
    empty series the report would draw as a missing premium line."""
    try:
        pf.fetch("BANKNIFTY", ROWS[0][0].date(), config=[])
    except SystemExit as exc:
        assert "no Grafana instance is running BANKNIFTY" in str(exc)
    else:
        raise AssertionError("an unavailable index should refuse to fetch")


if __name__ == "__main__":
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn()
            print(f"  ok  {name}")
    print("\nall passed")
