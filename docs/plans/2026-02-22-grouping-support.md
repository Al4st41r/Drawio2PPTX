# Grouping Support Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Preserve Draw.io group structure in the generated PPTX so grouped shapes remain grouped and editable in PowerPoint.

**Architecture:** The parser detects which vertex cells are group parents by checking which cell IDs appear as `parent_id` of other vertex cells. The engine then processes vertices in a tree walk, creating `python-pptx` `GroupShape` objects for group cells and adding child shapes into the correct container. Invisible group cells (style `"group"`) produce a PPTX group with no visual shape; visible containers also get a shape rendered inside the group.

**Tech Stack:** Python 3.13, `python-pptx>=1.0.2`, `pytest` (dev dependency), `uv`.

---

## Task 1: Add pytest and create test fixture

**Files:**
- Modify: `pyproject.toml`
- Create: `tests/fixtures/grouped_diagram.drawio`
- Create: `tests/test_grouping.py`

**Step 1: Add pytest to dev dependencies**

```bash
uv add --dev pytest
```

Expected: `pyproject.toml` updated, `uv.lock` updated.

**Step 2: Create the grouped diagram fixture**

Create `tests/fixtures/grouped_diagram.drawio`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<mxfile>
  <diagram name="Groups">
    <mxGraphModel>
      <root>
        <mxCell id="0"/>
        <mxCell id="1" parent="0"/>
        <mxCell id="g1" value="" style="group" vertex="1" parent="1">
          <mxGeometry x="100" y="100" width="200" height="100" as="geometry"/>
        </mxCell>
        <mxCell id="a" value="A" style="rounded=1;whiteSpace=wrap;" vertex="1" parent="g1">
          <mxGeometry x="10" y="20" width="80" height="40" as="geometry"/>
        </mxCell>
        <mxCell id="b" value="B" style="rounded=1;whiteSpace=wrap;" vertex="1" parent="g1">
          <mxGeometry x="110" y="20" width="80" height="40" as="geometry"/>
        </mxCell>
        <mxCell id="c" value="C" style="rounded=1;whiteSpace=wrap;" vertex="1" parent="1">
          <mxGeometry x="50" y="300" width="100" height="50" as="geometry"/>
        </mxCell>
        <mxCell id="e1" value="" edge="1" source="a" target="b" parent="1">
          <mxGeometry relative="1" as="geometry"/>
        </mxCell>
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
```

- `g1`: invisible group container at absolute (100, 100)
- `a`: inside group, relative (10, 20) → absolute (110, 120)
- `b`: inside group, relative (110, 20) → absolute (210, 120)
- `c`: top-level shape at (50, 300)
- `e1`: connector from `a` to `b`

**Step 3: Create the test file with failing tests**

Create `tests/test_grouping.py`:

```python
import os
import tempfile

import pytest
from pptx import Presentation
from pptx.shapes.group import GroupShape

from converter import convert
from converter.parser import DrawioParser

FIXTURE = os.path.join(os.path.dirname(__file__), "fixtures", "grouped_diagram.drawio")


# --- Parser tests ---

def test_parser_marks_group_cell():
    pages = DrawioParser(FIXTURE).parse()
    vertices, _ = pages[0]["data"]
    groups = [v for v in vertices if v["is_group"]]
    assert len(groups) == 1
    assert groups[0]["id"] == "g1"


def test_parser_non_group_cells_not_marked():
    pages = DrawioParser(FIXTURE).parse()
    vertices, _ = pages[0]["data"]
    non_groups = [v for v in vertices if not v["is_group"]]
    assert len(non_groups) == 3  # a, b, c


def test_parser_group_child_absolute_positions():
    pages = DrawioParser(FIXTURE).parse()
    vertices, _ = pages[0]["data"]
    a = next(v for v in vertices if v["id"] == "a")
    b = next(v for v in vertices if v["id"] == "b")
    assert a["x"] == 110
    assert a["y"] == 120
    assert b["x"] == 210
    assert b["y"] == 120


# --- Engine/integration tests ---

@pytest.fixture
def output_pptx(tmp_path):
    out = str(tmp_path / "out.pptx")
    convert(FIXTURE, out)
    return Presentation(out)


def test_slide_contains_one_group(output_pptx):
    slide = output_pptx.slides[0]
    groups = [s for s in slide.shapes if isinstance(s, GroupShape)]
    assert len(groups) == 1


def test_group_contains_two_child_shapes(output_pptx):
    slide = output_pptx.slides[0]
    group = next(s for s in slide.shapes if isinstance(s, GroupShape))
    assert len(group.shapes) == 2


def test_top_level_shape_exists_outside_group(output_pptx):
    slide = output_pptx.slides[0]
    non_groups = [s for s in slide.shapes if not isinstance(s, GroupShape)]
    # Should contain: shape C + connector e1 (edge labels excluded in this diagram)
    shape_names = {s.name for s in non_groups}
    assert len(non_groups) >= 2  # at minimum C and the connector
```

**Step 4: Run tests to confirm they fail**

```bash
uv run pytest tests/test_grouping.py -v
```

Expected: All tests FAIL — `is_group` key not present in vertex dicts.

---

## Task 2: Implement `is_group` in the parser

**Files:**
- Modify: `converter/parser.py` (lines 88–118, the `_extract_elements` method)

**Step 1: Add group detection after the main loop**

In `converter/parser.py`, immediately before `return vertices, edges` at the end of `_extract_elements`, add:

```python
        # Detect group cells: vertices whose ID is referenced as parent_id by other vertices
        vertex_ids = {v['id'] for v in vertices}
        group_cell_ids = {v['parent_id'] for v in vertices if v['parent_id'] in vertex_ids}
        for v in vertices:
            v['is_group'] = v['id'] in group_cell_ids
```

The full tail of `_extract_elements` should now read:

```python
            elif cell_data['edge']:
                edges.append(cell_data)

        # Detect group cells: vertices whose ID is referenced as parent_id by other vertices
        vertex_ids = {v['id'] for v in vertices}
        group_cell_ids = {v['parent_id'] for v in vertices if v['parent_id'] in vertex_ids}
        for v in vertices:
            v['is_group'] = v['id'] in group_cell_ids

        return vertices, edges
```

**Step 2: Run the parser tests**

```bash
uv run pytest tests/test_grouping.py::test_parser_marks_group_cell tests/test_grouping.py::test_parser_non_group_cells_not_marked tests/test_grouping.py::test_parser_group_child_absolute_positions -v
```

Expected: All 3 PASS. Engine tests still FAIL.

**Step 3: Commit**

```bash
git add converter/parser.py tests/test_grouping.py tests/fixtures/grouped_diagram.drawio pyproject.toml uv.lock
git commit -m "feat: detect group cells in parser, add grouping tests"
```

---

## Task 3: Refactor engine — extract helpers

**Files:**
- Modify: `converter/engine.py`

This task extracts the per-vertex rendering logic into reusable methods, and updates `_add_image` to accept a `shapes_container` parameter. No behaviour changes — the integration test for the flat (non-group) case should still pass after this step.

**Step 1: Update `_add_image` signature**

Change `_add_image(self, vertex)` to `_add_image(self, shapes_container, vertex)` and replace `self.slide.shapes.add_picture(...)` with `shapes_container.add_picture(...)`.

The updated method:

```python
    def _add_image(self, shapes_container, vertex):
        """Extract embedded image data and add it to shapes_container as a picture."""
        style = vertex['style']
        image_uri = style.get('image', '')

        match = re.match(r'data:image/[^,]+,(.+)', image_uri)
        if not match:
            return None

        try:
            image_bytes = base64.b64decode(match.group(1))
        except Exception:
            return None

        image_stream = io.BytesIO(image_bytes)

        pic = shapes_container.add_picture(
            image_stream,
            px_to_emu(vertex['x']),
            px_to_emu(vertex['y']),
            px_to_emu(vertex['width']),
            px_to_emu(vertex['height']),
        )
        return pic
```

**Step 2: Add `_is_pure_group_container`**

Add this method to `PptxGenerator`:

```python
    def _is_pure_group_container(self, vertex):
        """Return True if this vertex is an invisible Draw.io group container (style='group')."""
        s = vertex['style_str'].strip()
        return s == 'group' or s.startswith('group;') or s.startswith('group ')
```

**Step 3: Add `_add_vertex_to_shapes`**

Extract the per-vertex body from the current `add_vertices` loop into a new private method:

```python
    def _add_vertex_to_shapes(self, shapes_container, v):
        """Render a single vertex into shapes_container. Returns the created shape or None."""
        x, y, w, h = v["x"], v["y"], v["width"], v["height"]
        style = v["style"]

        if style.get('shape') == 'image' and 'image' in style:
            pic = self._add_image(shapes_container, v)
            if pic:
                self.id_to_shape[v["id"]] = pic
            return pic

        style['_width'] = w
        style['_height'] = h

        shape_type = get_shape_type(style)

        shape = shapes_container.add_shape(
            shape_type, px_to_emu(x), px_to_emu(y), px_to_emu(w), px_to_emu(h)
        )

        self._apply_shape_style(shape, style)

        if v["value"]:
            self._apply_text(shape, v["value"], style)

        self.id_to_shape[v["id"]] = shape
        return shape
```

**Step 4: Rewrite `add_vertices` to use `_add_vertex_to_shapes` (flat, no grouping yet)**

Replace the body of `add_vertices` with:

```python
    def add_vertices(self, vertices):
        if not self.slide:
            self.create_slide()

        for v in vertices:
            self._add_vertex_to_shapes(self.slide.shapes, v)
```

This is a pure refactor — behaviour unchanged.

**Step 5: Verify no regressions**

```bash
uv run pytest tests/test_grouping.py -v
```

Expected: parser tests PASS, engine tests still FAIL (groups not created yet).

Also do a quick smoke test with the existing diagram:

```bash
uv run main.py tests/TestDiagram.drawio /tmp/smoke_test.pptx && echo "OK"
```

Expected: `OK` with no errors.

**Step 6: Commit**

```bash
git add converter/engine.py
git commit -m "refactor: extract _add_vertex_to_shapes and _add_image helpers in engine"
```

---

## Task 4: Implement `_add_vertices_recursive` and wire up grouping

**Files:**
- Modify: `converter/engine.py`

**Step 1: Add `_add_vertices_recursive`**

Add this method to `PptxGenerator`:

```python
    def _add_vertices_recursive(self, container, vertex_list, children_map, group_ids):
        """
        Add vertex_list into container (a Slide or GroupShape), recursing into groups.

        container   — slide or GroupShape whose .shapes is the target shapetree
        vertex_list — vertices to add at this level
        children_map — {parent_id: [child_vertex, ...]}
        group_ids   — set of vertex IDs that are group cells
        """
        for v in vertex_list:
            if v['id'] in group_ids:
                grp = container.shapes.add_group_shape()
                self.id_to_shape[v['id']] = grp

                # If the group container itself has visible styling, render it inside the group
                if not self._is_pure_group_container(v):
                    sh = self._add_vertex_to_shapes(grp.shapes, v)
                    if sh:
                        self.id_to_shape[v['id']] = sh

                children = children_map.get(v['id'], [])
                self._add_vertices_recursive(grp, children, children_map, group_ids)
            else:
                self._add_vertex_to_shapes(container.shapes, v)
```

**Step 2: Rewrite `add_vertices` to use the recursive helper**

Replace the current `add_vertices` body:

```python
    def add_vertices(self, vertices):
        if not self.slide:
            self.create_slide()

        vertex_ids = {v['id'] for v in vertices}
        group_ids = {v['id'] for v in vertices if v.get('is_group')}

        # Build parent → children map
        children_map = {}
        for v in vertices:
            children_map.setdefault(v['parent_id'], []).append(v)

        # Top-level: vertices whose parent_id is not another vertex (i.e. parent is a layer cell)
        top_level = [v for v in vertices if v['parent_id'] not in vertex_ids]

        self._add_vertices_recursive(self.slide, top_level, children_map, group_ids)
```

**Step 3: Run the full test suite**

```bash
uv run pytest tests/test_grouping.py -v
```

Expected: All 6 tests PASS.

**Step 4: Smoke test with existing diagram**

```bash
uv run main.py tests/TestDiagram.drawio /tmp/smoke_test.pptx && echo "OK"
```

Expected: `OK` with no errors.

**Step 5: Commit**

```bash
git add converter/engine.py
git commit -m "feat: add group shape support — group cells produce native PPTX groups"
```

---

## Task 5: Bump version and final verification

**Files:**
- Modify: `converter/__init__.py`

**Step 1: Bump version**

In `converter/__init__.py`, change `__version__ = "1.0.8"` to `__version__ = "1.0.9"`.

Also update `pyproject.toml` version from `1.0.8` to `1.0.9`.

**Step 2: Run all tests**

```bash
uv run pytest tests/test_grouping.py -v
```

Expected: All 6 PASS.

**Step 3: Final commit**

```bash
git add converter/__init__.py pyproject.toml
git commit -m "chore: bump version to 1.0.9"
```

---

## Notes

- Connectors (`add_edges`) are unaffected. They are always added to `self.slide.shapes` and use coordinate-based positioning. A connector source/target that is a group cell will use the `GroupShape` stored in `id_to_shape`; `GroupShape` exposes `.left/.top/.width/.height` so centre-point calculation works correctly.
- Edge labels (`_add_edge_label`) are also unaffected — always added to `self.slide.shapes`.
- The `is_group` flag set by the parser is `False` when there are no groups in a diagram, so existing non-grouped diagrams pass through the new `add_vertices` path without change.
- Nested groups (groups within groups) are handled automatically by the recursion.
