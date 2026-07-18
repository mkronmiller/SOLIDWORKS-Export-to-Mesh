"""
reorient_3mf.py - bake a SolidWorks coordinate system into a 3MF export.

Called by the Export to Mesh macro (v3) as:
    python reorient_3mf.py <file.3mf> r00 r01 r02 r10 r11 r12 r20 r21 r22 tx ty tz

The 12 numbers are the first 12 elements of the SolidWorks MathTransform
ArrayData for the print-orientation coordinate system (rotation 3x3 row-major,
then translation in METERS), as returned by GetCoordinateSystemTransformByName.

That transform maps coordinate-system space -> model space. Exported vertices
are in model space, so this script applies the INVERSE, re-expressing every
vertex in the coordinate system's frame (Z = print up). The rotation is baked
directly into the vertex coordinates, so the result is orientation-correct in
any slicer with no reliance on 3MF transform support.

Uses only the Python standard library. On error, writes reorient_error.txt
next to the 3MF and exits nonzero (the macro logs a WARN with the exit code).
"""

import os
import re
import sys
import shutil
import tempfile
import zipfile

MODEL_PATH_IN_ZIP = "3D/3dmodel.model"

VERTEX_RE = re.compile(
    r'(<vertex\s+x=")([-0-9.eE+]+)("\s+y=")([-0-9.eE+]+)("\s+z=")([-0-9.eE+]+)(")'
)


def main():
    if len(sys.argv) != 14:
        raise ValueError(
            "Expected 13 arguments (file + 12 transform values), got %d"
            % (len(sys.argv) - 1)
        )

    path = sys.argv[1]
    vals = [float(x) for x in sys.argv[2:14]]

    # SolidWorks MathTransform: rotation rows, row-vector convention
    # (model_point = cs_point * R + t). Translation arrives in meters.
    R = [vals[0:3], vals[3:6], vals[6:9]]
    t = [v * 1000.0 for v in vals[9:12]]  # meters -> millimeters

    def to_cs(p):
        """Inverse transform: cs_point = (model_point - t) * R^T."""
        q = (p[0] - t[0], p[1] - t[1], p[2] - t[2])
        return [
            q[0] * R[j][0] + q[1] * R[j][1] + q[2] * R[j][2]
            for j in range(3)
        ]

    def replace_vertex(m):
        p = (float(m.group(2)), float(m.group(4)), float(m.group(6)))
        c = to_cs(p)
        return "%s%.6f%s%.6f%s%.6f%s" % (
            m.group(1), c[0], m.group(3), c[1], m.group(5), c[2], m.group(7)
        )

    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        if MODEL_PATH_IN_ZIP not in names:
            raise ValueError("%s not found in archive" % MODEL_PATH_IN_ZIP)
        model_xml = zin.read(MODEL_PATH_IN_ZIP).decode("utf-8")
        others = {n: zin.read(n) for n in names if n != MODEL_PATH_IN_ZIP}

    new_xml, n_verts = VERTEX_RE.subn(replace_vertex, model_xml)
    if n_verts == 0:
        raise ValueError("No vertices matched - unexpected 3MF structure")

    # Write to a temp file, then atomically replace the original
    fd, tmp_path = tempfile.mkstemp(suffix=".3mf", dir=os.path.dirname(path) or ".")
    os.close(fd)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in others.items():
                zout.writestr(name, data)
            zout.writestr(MODEL_PATH_IN_ZIP, new_xml)
        shutil.move(tmp_path, path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # write a breadcrumb next to the file, exit nonzero
        try:
            target_dir = (
                os.path.dirname(os.path.abspath(sys.argv[1]))
                if len(sys.argv) > 1
                else "."
            )
            with open(os.path.join(target_dir, "reorient_error.txt"), "a") as f:
                f.write("reorient_3mf failed: %r\nargs: %r\n\n" % (exc, sys.argv))
        except Exception:
            pass
        sys.exit(1)
