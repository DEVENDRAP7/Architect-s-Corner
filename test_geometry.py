"""Correctness check for the tag-free geometry engine (_detect_geometry).
Runs on test_building.dxf, whose column layout is known. No framework."""
import ezdxf
import ArchTools as A


def test_detect():
    doc = ezdxf.readfile("test_building.dxf")
    d = A._detect_geometry(doc.modelspace())
    # test_building draws 2 floor plans, 9 columns each (450x450/300x450/300x300)
    assert d["floor_count"] == 2, d["floor_count"]
    assert d["columns_per_floor"] == 9, d["columns_per_floor"]
    assert 90 <= d["footprint"] <= 160, d["footprint"]          # ~120 m2 hull
    assert 4.0 <= d["grid"] <= 6.5, d["grid"]                   # ~5 m grid
    assert (450, 450) in d["schedule"], dict(d["schedule"])
    # clustering must isolate real sheets, not smear everything into one blob
    assert d["total_columns"] == 18, d["total_columns"]
    print("geometry engine OK:", {k: d[k] for k in
          ("floor_count", "columns_per_floor", "footprint", "grid")})


if __name__ == "__main__":
    test_detect()
