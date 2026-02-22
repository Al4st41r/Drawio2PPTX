# Grouping Support Design

**Date:** 2026-02-22
**Feature:** Preserve Draw.io group structure in PowerPoint output
**Status:** Approved

---

## Background

Draw.io groups are represented in XML as vertex cells (`vertex="1"`) whose `id` is referenced in the `parent` attribute of child vertex cells. Currently the parser produces a flat vertex list with absolute positions, and the engine adds all shapes directly to the slide. Group structure is lost.

PowerPoint supports native shape groups (`<p:grpSp>`). `python-pptx` exposes this via `slide.shapes.add_group_shape()`, which returns a `GroupShape` whose `.shapes` property accepts the same `add_shape`, `add_picture`, `add_textbox`, and `add_group_shape` calls as `SlideShapes`. Positions are passed in absolute slide coordinates; python-pptx handles the internal coordinate transformation. Group bounding boxes are auto-calculated.

---

## Approach

Full recursive grouping. Detect group cells in the parser. In the engine, build a children map keyed by `parent_id`, then walk the tree depth-first, creating PPTX `GroupShape` objects for Draw.io group cells and adding child shapes into the correct container. Nested groups are handled naturally by the recursion.

Connectors and edge labels remain at slide level — connector positioning is coordinate-based and is unaffected by group membership.

---

## Changes

### `converter/parser.py`

After the main vertex/edge collection loop in `_extract_elements`, detect which vertices are group cells:

```python
vertex_ids = {v['id'] for v in vertices}
group_cell_ids = {v['parent_id'] for v in vertices if v['parent_id'] in vertex_ids}
for v in vertices:
    v['is_group'] = v['id'] in group_cell_ids
```

This adds an `is_group` boolean to each vertex's data dict. No other parser changes are needed; `parent_id` is already recorded.

### `converter/engine.py`

Three changes:

**1. Refactor `add_vertices`** to build a children map and invoke a recursive helper:

```python
def add_vertices(self, vertices):
    if not self.slide:
        self.create_slide()

    vertex_ids = {v['id'] for v in vertices}
    group_ids = {v['id'] for v in vertices if v.get('is_group')}

    # children_map: parent_id → [vertex, ...]
    children_map = {}
    for v in vertices:
        pid = v['parent_id']
        children_map.setdefault(pid, []).append(v)

    # Top-level: vertices whose parent is not another vertex (i.e., parent is a layer cell)
    top_level = [v for v in vertices if v['parent_id'] not in vertex_ids]

    self._add_vertices_recursive(self.slide, top_level, children_map, group_ids)
```

**2. Add `_add_vertices_recursive`:**

```python
def _add_vertices_recursive(self, container, vertex_list, children_map, group_ids):
    for v in vertex_list:
        if v['id'] in group_ids:
            grp = container.shapes.add_group_shape()
            self.id_to_shape[v['id']] = grp

            # Add a visual shape inside the group if the container is not invisible
            if not self._is_pure_group_container(v):
                sh = self._add_vertex_to_shapes(grp.shapes, v)
                if sh:
                    self.id_to_shape[v['id']] = sh

            children = children_map.get(v['id'], [])
            self._add_vertices_recursive(grp, children, children_map, group_ids)
        else:
            self._add_vertex_to_shapes(container.shapes, v)
```

**3. Extract `_add_vertex_to_shapes` and `_is_pure_group_container`:**

`_add_vertex_to_shapes(shapes_container, vertex)` — the existing per-vertex logic from `add_vertices`, now accepting a target `shapes_container`. The `_add_image` helper is similarly adjusted to accept a `shapes_container`.

`_is_pure_group_container(vertex)` — returns `True` if the style string begins with `"group"`, indicating a Draw.io invisible group cell with no visual representation.

---

## Invisible vs Visible Group Containers

- `style="group"` or `style="group;..."` → invisible. PPTX `GroupShape` created, no shape added inside it.
- Any other style (e.g. `style="swimlane;..."`) → visible. PPTX `GroupShape` created, and the container cell is also added as a visual shape inside the group, alongside its children.

---

## Edge Cases

| Case | Handling |
|---|---|
| Nested groups | Handled recursively — outer group created before inner |
| Connector source/target is a group cell | `id_to_shape` stores the `GroupShape`; `.left/.top/.width/.height` work on `GroupShape`, so connector positioning is unaffected |
| Image inside a group | `_add_image` receives `shapes_container`; `GroupShapes.add_picture` is available |
| Edge labels | Added to `self.slide.shapes` directly — no change |
| Draw.io files with no groups | Top-level only; `children_map` lookups return empty lists; no groups created |

---

## Files Modified

- `converter/parser.py` — ~6 lines added
- `converter/engine.py` — `add_vertices` refactored; 3 private methods added/modified

No changes to `ppt_map.py`, `utils.py`, `__init__.py`, or `webapp/`.
