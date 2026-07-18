# SOLIDWORKS-Export-to-Mesh (fork)

> **AI disclosure:** The modifications in this fork were developed with AI assistance
> (Anthropic's Claude Fable), then human-tested against a real SolidWorks + NAS + PrusaSlicer
> workflow. Review before use, as with any macro that writes files.

This is a fork of [SigmaRelief/SOLIDWORKS-Export-to-Mesh](https://github.com/SigmaRelief/SOLIDWORKS-Export-to-Mesh)
with fixes and additions for automated, orientation-correct batch export. All credit
for the original macro, GUI, and body-renaming concept goes to SigmaRelief.

## What this fork adds

- **Print-orientation-correct exports.** Exports are re-expressed in a named
  coordinate system (e.g. `print_orientation`) by baking its transform directly
  into the 3MF vertices via a Python post-processing step (`reorient_3mf.py`).
  Parts land in the slicer lying correctly on the bed, every time. See
  [Why orientation is handled in Python](#why-orientation-is-handled-in-python).
- **Network/UNC path support** (`\\server\share\...`). The original folder
  creation crashed with run-time error 52 on UNC paths.
- **Run logging.** Every run appends to `ExportLog.txt` in the export folder:
  settings used, per-file OK/FAIL/WARN status, and a summary.
- **Robustness fixes:** export success verified on disk instead of trusting the
  SaveAs API return value (which returns False-with-no-error on success for some
  formats); overwrite guard when exporting all configurations without
  `[ConfigName]` in the name template; guarded Python steps (a bad Python path
  logs a warning instead of aborting the batch); safe INI parsing on first run;
  larger INI buffer for long paths. Full list in the module header changelog.

### New INI setting

| Key | Purpose |
| --- | --- |
| `CoordinateSystemName` | Name of the coordinate system feature used to orient exports. Empty = model space (original behavior). |

### Print orientation setup (once per part)

![SW_Coor_Example](doc/sw_coordinate_plane_example.jpg)

1. In each part, create a coordinate system (**Insert → Reference Geometry →
   Coordinate System**). For the **Z axis**, select the flat face that should
   sit on the print bed — then use the flip arrow so **Z points up out of the
   part**, away from the bed. Optionally set X along the part's length. Face
   selections survive remodeling better than edge selections.
2. Name it to match the INI setting (default in this fork: `print_orientation`).
   Use the same name in every part so one INI serves the whole project.

If a model has no coordinate system by that name, the macro logs a WARN and
exports in model space.

### Why orientation is handled in Python

Two SolidWorks export quirks make the "obvious" approaches fail, documented here
so nobody has to rediscover them:

1. **STEP exports carry a hidden inverse placement.** When exporting STEP with a
   custom output coordinate system, SolidWorks bakes the transform into the
   geometry *and* writes an `AXIS2_PLACEMENT_3D` encoding the inverse. Importers
   that honor the placement (PrusaSlicer does) undo the orientation — the output
   coordinate system option is effectively a no-op for STEP → PrusaSlicer. Use
   3MF/STL for slicing; keep STEP for CAD source distribution where orientation
   is cosmetic.
2. **The API rejects the output-coordinate-system preference.** Setting
   `swUserPreferenceStringValue_e.swFileSaveAsCoordinateSystem` via the API
   returns False and reads back empty on the SolidWorks version this fork was
   developed on, even with a verified-correct coordinate system name — so
   macro-driven exports silently come out in model space. The manual Save As
   dialog works; the API route does not.

This fork therefore reads the coordinate system's transform with
`GetCoordinateSystemTransformByName` (a reliable read-only API) and passes it to
`reorient_3mf.py`, which rewrites the 3MF's vertex coordinates in the print
frame. The rotation is baked into the geometry itself, so the result is
orientation-correct in any slicer regardless of 3MF transform support.

### Python path note

Set **PythonPath** to the **full path** to `python.exe` (find it with `py -0p`
in a Command Prompt), e.g.
`C:\Users\you\AppData\Local\Programs\Python\Python314\python.exe`.
Bare commands like `python` or `py` resolve against the environment captured
when SolidWorks launched and fail unpredictably across sessions; full paths are
deterministic. Only the Python standard library is used — no packages needed.

---

# Original project description

This is a simple SOLIDWORKS macro to export files to the .3MF format and update
the body names within the file. By default, SOLIDWORKS assigns random names to
the bodies saved within these files and it becomes confusing when imported into
some FDM slicers like Slic3r, PrusaSlicer, or
[SuperSlicer](https://github.com/supermerill/SuperSlicer/releases) (highly
recommended) that display body names as opposed to file names. This macro
automatically sets body names to the file name and applies sequential numbers
for multi-body files.

Additional functionality from my other SOLIDWORKS macros has crept into this and
there are provisions to to export all configurations in a single click, change
the file name, export location, file format, and ability to save preferences
from within the GUI.

![Screenshot](Doc/Export%20Options%20Screenshot.png)

# Installation

* Place all the macro files (including this fork's `reorient_3mf.py`) together
  in your preferred macro folder — the macro locates its companion files
  relative to itself. A local path is recommended over a network location.
* In SOLIDWORKS, click Tools/Customize/Commands/Macro, Select "New Macro
  Button", and place it in the GUI.
* Point the new button to "Export to Mesh.swp"
* The icon should default to "Export to Mesh.bmp", set it manually if it does
  not.
* Add a tool tip or prompt if desired.

![SOLIDWORKS Macro Setup](Doc/Macro%20Setup.png)

# Usage

Set the Python path (see [Python path note](#python-path-note) above — this
fork recommends a full path over relying on the Windows PATH variable), then
use is as easy as clicking the "Export" button.

![Screenshot with all options](Doc/Export%20Options%20All%20Screenshot.png)

Changes to the default export location can also be saved with the "Save
Settings..." box. The export location can optionally be saved as relative to
the user's profile folder by including the [UserDir] variable

Tolerance on mesh files can be optionally adjusted at time of export. If save
settings is selected, these values are set as the new SOLIDWORKS defaults.
SOLIDWORKS uses coarse values by default so it is probably useful to
significantly reduce this value if part accuracy and finish are important.

The "Export all configurations" option iterates through all configurations and
exports them as individual files. The file name (and thus internal body names)
can be configured via a dropdown to place the SOLIDWORKS configuration name at
the beginning or end of the file name, or use the configuration name only.
Entries within this list can be adapted with any desired delimiter or order by
typing directly in the dropdown box and saved with the "Save settings..." box.
*(Fork note: if the template lacks `[ConfigName]`, it is appended automatically
to prevent configurations overwriting each other.)*

The file format dropdown only includes .3MF, .STL, and .STEP by default (in
both cases to suite your preference). The body re-naming and orientation
correction are only executed with .3MF formats, but most valid file formats are
supported should you have a need to do something like exporting all
configurations to another format (not limited to mesh types).

The underlying body re-name happens entirely within Python and this portion of
the code may be applicable to other CAD tools with minor updates to search for
their specific body naming convention.

Variables including SOLIDWORKS Custom Properties, Configuration Names, and
current date/time in multiple formats can be included in the File Name and File
path fields for single use, or saved as part of the defaults by clicking the
Use Variables button.

![Screenshot of Variable Inputs](Doc/Variable%20Inputs%20Screenshot.png)

# Requirements

The rename and orientation functionality requires Python 3.X to be installed
(standard library only).
The [Date ####] variable requires Microsoft Excel to be installed.

# Project status

**Fork:** tested with SolidWorks 2026 Maker Edition and PrusaSlicer 2.9.6 against network (UNC) storage. The original disclaimer applies doubly: this is a hobby project and should be treated accordingly. No guarantees or warranty are provided.

**Upstream:** tested only with SOLIDWORKS 2019, PrusaSlicer 2.2.0, and
SuperSlicer 2.2.53.

# License

GPL-3.0, inherited from the upstream project.