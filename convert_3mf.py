#!/usr/bin/env python3
"""
Convert 3MF files to the current Orca Slicer format.

Handles:
1. Old BambuStudio format (inline mesh + assembly, no p:path)
2. Bare 3MF files (mesh objects, no BambuStudio wrapper)
3. PrusaSlicer 3MF files (production ext but no Bambu metadata)
4. Multi-object/multi-part files
5. Color/material data (basematerials, colorgroup, pid/p1, paint_color)
6. Sliced .gcode.3mf detection

Uses streaming XML for large files — never loads full mesh data into a DOM.
"""

import argparse
import re
import shutil
import sys
import tempfile
import uuid
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

NS_CORE = "http://schemas.microsoft.com/3dmanufacturing/core/2015/02"
NS_PRODUCTION = "http://schemas.microsoft.com/3dmanufacturing/production/2015/06"
NS_BAMBU = "http://schemas.bambulab.com/package/2021"
NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_3DMODEL_TYPE = "http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel"
NS_THUMBNAIL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"

ET.register_namespace("", NS_CORE)
ET.register_namespace("p", NS_PRODUCTION)
ET.register_namespace("BambuStudio", NS_BAMBU)


def make_uuid():
    return str(uuid.uuid4())


def classify_3mf(model_data, zip_names=None):
    """
    Classify a 3MF model's format.
    Reads the first 50KB and last 50KB to catch structure tags
    that may appear after large mesh blocks.

    Args:
        model_data: bytes or string of the model XML (or head+tail)
        zip_names: optional list of filenames in the ZIP for extra signals

    Returns one of:
        "orca_ready"  - has p:path component refs (Orca/modern BambuStudio)
        "bambu_old"   - BambuStudio assembly + inline mesh, no p:path
        "prusa"       - PrusaSlicer format (has slic3rpe metadata)
        "bare"        - plain 3MF from modelers/CAD tools
        "sliced"      - sliced .gcode.3mf with no geometry
        "unknown"     - can't determine
    """
    if isinstance(model_data, bytes):
        if len(model_data) <= 100000:
            combined = model_data.decode("utf-8", errors="replace")
        else:
            header = model_data[:50000].decode("utf-8", errors="replace")
            tail = model_data[-50000:].decode("utf-8", errors="replace")
            combined = header + tail
    else:
        if len(model_data) <= 100000:
            combined = model_data
        else:
            combined = model_data[:50000] + model_data[-50000:]

    # Check ZIP contents for extra signals
    if zip_names:
        has_gcode = any(n.endswith(".gcode") for n in zip_names)
        has_objects_dir = any(n.startswith("3D/Objects/") for n in zip_names)
        has_slic3r = any("Slic3r" in n or "slic3r" in n for n in zip_names)
    else:
        has_gcode = False
        has_objects_dir = False
        has_slic3r = False

    has_mesh = bool(re.search(r'<mesh', combined))
    has_component = bool(re.search(r'<component\b', combined))
    has_p_path = bool(re.search(r'p:path\s*=', combined))
    has_object = bool(re.search(r'<object\b', combined))
    has_bambu = bool(re.search(r'BambuStudio|bambulab', combined))
    has_prusa = bool(re.search(r'PrusaSlicer|Slic3r', combined)) or has_slic3r

    # Sliced file with no geometry
    if has_gcode and not has_mesh:
        return "sliced"

    # Already Orca-compatible (has p:path AND has Objects dir)
    if has_p_path and has_objects_dir:
        return "orca_ready"
    if has_p_path:
        return "orca_ready"

    # PrusaSlicer (may or may not need conversion)
    if has_prusa and has_object:
        if has_component:
            return "prusa"  # has assembly but may lack Bambu metadata
        return "bare"  # PrusaSlicer bare export

    # Old BambuStudio
    if has_component:
        return "bambu_old"

    # Bare 3MF
    if has_object and has_mesh:
        return "bare"
    if has_object:
        return "bare"

    return "unknown"


def needs_conversion(model_xml):
    """Check if this 3MF needs conversion."""
    return classify_3mf(model_xml) in ("bambu_old", "bare", "prusa")


# ---------------------------------------------------------------------------
# Content_Types.xml generation
# ---------------------------------------------------------------------------

CONTENT_TYPES_XML = '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="model" ContentType="application/vnd.ms-package.3dmanufacturing-3dmodel+xml"/>
 <Default Extension="png" ContentType="image/png"/>
 <Default Extension="gcode" ContentType="text/x.gcode"/>
</Types>'''


def _generate_root_rels():
    """Generate _rels/.rels pointing to the main 3D model."""
    return '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Target="/3D/3dmodel.model" Id="rel-1" Type="http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel"/>
</Relationships>'''


def _generate_model_rels(object_filenames):
    """Generate 3D/_rels/3dmodel.model.rels."""
    entries = []
    for i, fname in enumerate(object_filenames, 1):
        entries.append(f' <Relationship Target="/3D/Objects/{fname}" Id="rel-{i}" Type="{NS_3DMODEL_TYPE}"/>')
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            + "\n".join(entries) + '\n</Relationships>')


# ---------------------------------------------------------------------------
# Color/material extraction
# ---------------------------------------------------------------------------

def _extract_materials(raw_text):
    """
    Extract material/color definitions from the model XML.
    Handles:
    - <basematerials> with <base> elements (core spec)
    - <colorgroup> with <color> elements (materials extension)
    Returns the raw XML blocks to preserve in sub-models.
    """
    materials_blocks = []

    # Find <basematerials ...>...</basematerials>
    for m in re.finditer(r'(<basematerials\b[^>]*>.*?</basematerials>)', raw_text, re.DOTALL):
        materials_blocks.append(m.group(1))

    # Find <m:colorgroup ...>...</m:colorgroup> or <colorgroup ...>
    for m in re.finditer(r'(<(?:m:)?colorgroup\b[^>]*>.*?</(?:m:)?colorgroup>)', raw_text, re.DOTALL):
        materials_blocks.append(m.group(1))

    # Find <m:color ...> standalone (not already inside a colorgroup)
    all_existing = "\n".join(materials_blocks)
    for m in re.finditer(r'(<(?:m:)?color\b[^>]*/>)', raw_text):
        if m.group(1) not in all_existing:
            materials_blocks.append(m.group(1))

    return materials_blocks


def _extract_namespaces(xml_header):
    """Extract all xmlns declarations from the model header."""
    return re.findall(r'xmlns(?::\w+)?="[^"]*"', xml_header[:2000])


# ---------------------------------------------------------------------------
# Multi-object bare 3MF conversion (streaming)
# ---------------------------------------------------------------------------

def _stream_convert_bare(zin, zout, source_name, input_name, strip_settings=False):
    """
    Convert a bare 3MF using streaming.
    Handles multiple objects, preserves color/material data.
    """
    with zin.open("3D/3dmodel.model") as f:
        raw_bytes = f.read()
    raw_text = raw_bytes.decode("utf-8", errors="replace")

    # Extract namespace declarations from model header
    ns_decls = _extract_namespaces(raw_text)
    extra_ns = " ".join(d for d in ns_decls
                        if "3dmanufacturing/core" not in d
                        and "production" not in d
                        and "bambulab" not in d)

    # Find ALL object IDs (scan full text for headers since they may be scattered)
    obj_pattern = re.compile(r'<object\s+[^>]*?\bid="(\d+)"[^>]*>')
    obj_ids = obj_pattern.findall(raw_text)
    if not obj_ids:
        print("No objects found in model")
        return False

    # Extract material/color blocks from resources
    material_blocks = _extract_materials(raw_text)
    materials_xml = "\n   ".join(material_blocks) if material_blocks else ""

    # Find each <object>...</object> block with its mesh
    obj_blocks = {}  # id -> full block text
    # Use a more robust pattern that handles nested elements
    for oid in obj_ids:
        # Find the start of this object
        obj_start_pattern = re.compile(rf'<object\s+[^>]*?\bid="{oid}"[^>]*>')
        start_match = obj_start_pattern.search(raw_text)
        if not start_match:
            continue

        # Find the matching </object>
        start_pos = start_match.start()
        # Count nesting
        depth = 0
        pos = start_pos
        end_pos = None
        while pos < len(raw_text):
            open_m = re.search(r'<object\b', raw_text[pos:])
            close_m = re.search(r'</object>', raw_text[pos:])

            if open_m and (not close_m or open_m.start() < close_m.start()):
                depth += 1
                pos = pos + open_m.start() + 1
            elif close_m:
                depth -= 1
                if depth == 0:
                    end_pos = pos + close_m.end()
                    break
                pos = pos + close_m.start() + 1
            else:
                break

        if end_pos:
            block = raw_text[start_pos:end_pos]
            # Only include if it has mesh data
            if "<mesh" in block:
                obj_blocks[oid] = block

    if not obj_blocks:
        print("No mesh objects found")
        return False

    # Create assembly structure
    max_id = max(int(x) for x in obj_ids)
    asm_id = str(max_id + 1)
    asm_uuid = make_uuid()
    build_uuid = make_uuid()

    object_filenames = []
    component_lines = []
    part_configs = []

    for idx, (oid, block) in enumerate(obj_blocks.items()):
        obj_filename = f"{source_name}_{oid}.model"
        part_uuid = make_uuid()
        object_filenames.append(obj_filename)

        # Extract the mesh from the block
        mesh_start = block.find("<mesh")
        mesh_end = block.rfind("</mesh>") + len("</mesh>")
        mesh_block = block[mesh_start:mesh_end] if mesh_start != -1 and mesh_end > len("</mesh>") - 1 else ""

        if not mesh_block:
            continue

        # Check for pid/p1 attributes on triangles (per-triangle material)
        has_pid = 'pid="' in mesh_block or "pid='" in mesh_block

        # Build sub-model with materials if needed
        resources_extra = ""
        if materials_xml and has_pid:
            resources_extra = f"\n   {materials_xml}"

        sub_model = f'''<?xml version="1.0" encoding="UTF-8"?>
<model xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02" xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06" {extra_ns} unit="millimeter" xml:lang="en-US" requiredextensions="p">
 <metadata name="BambuStudio:3mfVersion">1</metadata>
 <resources>{resources_extra}
  <object id="{oid}" p:UUID="{part_uuid}" type="model">
   {mesh_block}
  </object>
 </resources>
</model>'''

        zout.writestr(f"3D/Objects/{obj_filename}", sub_model.encode("utf-8"))

        component_lines.append(
            f'    <component p:path="/3D/Objects/{obj_filename}" objectid="{oid}" '
            f'p:UUID="{part_uuid}" transform="1 0 0 0 1 0 0 0 1 0 0 0"/>')

        # Determine extruder from pid if available
        extruder = str(idx + 1) if len(obj_blocks) > 1 else "1"
        part_configs.append(f'''    <part id="{oid}" subtype="normal_part">
      <metadata key="name" value="{source_name}_part{oid}"/>
      <metadata key="matrix" value="1 0 0 0 0 1 0 0 0 0 1 0 0 0 0 1"/>
      <metadata key="source_file" value="{input_name}"/>
      <metadata key="source_object_id" value="{idx}"/>
      <metadata key="source_volume_id" value="0"/>
      <metadata key="source_offset_x" value="0"/>
      <metadata key="source_offset_y" value="0"/>
      <metadata key="source_offset_z" value="0"/>
      <mesh_stat edges_fixed="0" degenerate_facets="0" facets_removed="0" facets_reversed="0" backwards_edges="0"/>
    </part>''')

    if not component_lines:
        print("No valid mesh blocks extracted")
        return False

    # --- Write main model ---
    item_uuid = make_uuid()
    main_model = f'''<?xml version="1.0" encoding="UTF-8"?>
<model xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02" xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06" xmlns:BambuStudio="http://schemas.bambulab.com/package/2021" {extra_ns} unit="millimeter" xml:lang="en-US" requiredextensions="p">
 <metadata name="Application">BambuStudio-2.3.2-rc2</metadata>
 <metadata name="BambuStudio:3mfVersion">1</metadata>
 <resources>
  <object id="{asm_id}" p:UUID="{asm_uuid}" type="model">
   <components>
{chr(10).join(component_lines)}
   </components>
  </object>
 </resources>
 <build p:UUID="{build_uuid}">
  <item objectid="{asm_id}" p:UUID="{item_uuid}" printable="1"/>
 </build>
</model>'''

    zout.writestr("3D/3dmodel.model", main_model.encode("utf-8"))

    # --- Write rels ---
    zout.writestr("3D/_rels/3dmodel.model.rels",
                  _generate_model_rels(object_filenames).encode("utf-8"))

    # --- Write model_settings.config ---
    settings = f'''<?xml version="1.0" encoding="UTF-8"?>
<config>
  <object id="{asm_id}">
    <metadata key="name" value="{source_name}"/>
    <metadata key="extruder" value="1"/>
{chr(10).join(part_configs)}
  </object>
  <plate>
    <metadata key="plater_id" value="1"/>
    <metadata key="plater_name" value=""/>
    <metadata key="locked" value="false"/>
    <model_instance>
      <metadata key="object_id" value="{asm_id}"/>
      <metadata key="instance_id" value="0"/>
      <metadata key="identify_id" value="1"/>
    </model_instance>
  </plate>
</config>'''
    zout.writestr("Metadata/model_settings.config", settings.encode("utf-8"))

    if not strip_settings:
        zout.writestr("Metadata/project_settings.config", b"")

    # --- Write Content_Types and root rels ---
    zout.writestr("[Content_Types].xml", CONTENT_TYPES_XML.encode("utf-8"))
    zout.writestr("_rels/.rels", _generate_root_rels().encode("utf-8"))

    # Copy other files (thumbnails, etc.) but skip what we generated
    skip = {"3D/3dmodel.model", "3D/_rels/3dmodel.model.rels",
            "Metadata/model_settings.config", "Metadata/project_settings.config",
            "[Content_Types].xml", "_rels/.rels"}
    for name in zin.namelist():
        if name not in skip and not name.startswith("3D/Objects/"):
            if strip_settings and name.startswith("Metadata/") and name.endswith(".config"):
                continue
            zout.writestr(name, zin.read(name))

    return True


# ---------------------------------------------------------------------------
# Old BambuStudio conversion (streaming)
# ---------------------------------------------------------------------------

def _stream_convert_bambu_old(zin, zout, source_name, strip_settings=False):
    """
    Convert old BambuStudio format using streaming.
    The model has assembly + inline mesh objects.
    Preserves color/material data.
    """
    raw_bytes = zin.read("3D/3dmodel.model")
    raw_text = raw_bytes.decode("utf-8", errors="replace")

    # Extract namespace declarations
    ns_decls = _extract_namespaces(raw_text)

    # Extract material blocks
    material_blocks = _extract_materials(raw_text)
    materials_xml = "\n   ".join(material_blocks) if material_blocks else ""

    # Find all <object>...</object> blocks
    obj_pattern = re.compile(r'<object\s+([^>]*?)>(.*?)</object>', re.DOTALL)
    mesh_objects = {}
    assembly_info = {}

    for m in obj_pattern.finditer(raw_text):
        attrs = m.group(1)
        body = m.group(2)
        obj_id = re.search(r'\bid="(\d+)"', attrs)
        if not obj_id:
            continue
        oid = obj_id.group(1)
        if "<mesh" in body:
            mesh_start = body.find("<mesh")
            mesh_end = body.rfind("</mesh>") + len("</mesh>")
            mesh_objects[oid] = {
                "match": m,
                "mesh_block": body[mesh_start:mesh_end],
                "attrs": attrs,
            }
        if "<component" in body:
            assembly_info[oid] = {
                "match": m,
                "body": body,
                "attrs": attrs,
            }

    if not mesh_objects:
        return False

    object_files = {}
    object_filenames = []

    # Process assembly objects
    new_main = raw_text
    for asm_id, asm in assembly_info.items():
        asm_uuid = make_uuid()

        comp_pattern = re.compile(r'<component\s+([^/]*?)/>', re.DOTALL)
        new_body = asm["body"]

        for cm in comp_pattern.finditer(asm["body"]):
            comp_attrs = cm.group(1)
            ref_match = re.search(r'objectid="(\d+)"', comp_attrs)
            if not ref_match:
                continue
            ref_id = ref_match.group(1)
            if ref_id not in mesh_objects:
                continue

            obj_filename = f"{source_name}_{ref_id}.model"
            part_uuid = make_uuid()
            object_filenames.append(obj_filename)

            mesh_block = mesh_objects[ref_id]["mesh_block"]

            # Check if mesh has pid references
            has_pid = 'pid="' in mesh_block
            resources_extra = ""
            if materials_xml and has_pid:
                resources_extra = f"\n   {materials_xml}"

            # Preserve extra namespace declarations for materials
            extra_ns = " ".join(d for d in ns_decls
                                if "3dmanufacturing/core" not in d
                                and "production" not in d
                                and "bambulab" not in d)

            sub_model = f'''<?xml version="1.0" encoding="UTF-8"?>
<model xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02" xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06" {extra_ns} unit="millimeter" xml:lang="en-US" requiredextensions="p">
 <metadata name="BambuStudio:3mfVersion">1</metadata>
 <resources>{resources_extra}
  <object id="{ref_id}" p:UUID="{part_uuid}" type="model">
   {mesh_block}
  </object>
 </resources>
</model>'''
            object_files[obj_filename] = sub_model.encode("utf-8")

            old_comp = cm.group(0)
            new_comp = old_comp.replace("/>", f' p:path="/3D/Objects/{obj_filename}" p:UUID="{part_uuid}"/>')
            new_body = new_body.replace(cm.group(0), new_comp, 1)

        old_asm_obj = asm["match"].group(0)
        new_attrs = asm["attrs"]
        if "p:UUID" not in new_attrs:
            new_attrs = new_attrs + f' p:UUID="{asm_uuid}"'
        new_asm_obj = f'<object {new_attrs}>{new_body}</object>'
        new_main = new_main.replace(old_asm_obj, new_asm_obj, 1)

    # Remove mesh objects from main model
    for mesh_id, mesh_info in mesh_objects.items():
        is_referenced = any(
            re.search(rf'objectid="{mesh_id}"', asm["body"])
            for asm in assembly_info.values()
        )
        if is_referenced:
            new_main = new_main.replace(mesh_info["match"].group(0), "", 1)

    # Only remove material definitions from main model if ALL mesh objects
    # were extracted to sub-models (otherwise some objects still need them)
    all_extracted = all(
        any(re.search(rf'objectid="{mid}"', asm["body"])
            for asm in assembly_info.values())
        for mid in mesh_objects.keys()
    )
    if all_extracted and material_blocks:
        for block in material_blocks:
            new_main = new_main.replace(block, "", 1)

    # Add UUIDs to build items
    def add_build_uuids(m):
        item_str = m.group(0)
        if "p:UUID" not in item_str:
            item_str = item_str.replace("/>", f' p:UUID="{make_uuid()}"/>')
        return item_str
    new_main = re.sub(r'<item\s+[^>]*/>', add_build_uuids, new_main)

    # Add UUID to build element
    def add_build_uuid(m):
        build_tag = m.group(0)
        if "p:UUID" not in build_tag:
            build_tag = build_tag.replace(">", f' p:UUID="{make_uuid()}">', 1)
        return build_tag
    new_main = re.sub(r'<build[^>]*>', add_build_uuid, new_main)

    # Ensure required namespaces and extensions
    if 'requiredextensions=' not in new_main[:500]:
        new_main = new_main.replace('<model ', '<model requiredextensions="p" ', 1)
    if 'xmlns:p=' not in new_main[:500]:
        new_main = new_main.replace(
            'xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02"',
            'xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02" '
            'xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06"',
            1)
    if 'xmlns:BambuStudio=' not in new_main[:500]:
        new_main = new_main.replace(
            'xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06"',
            'xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06" '
            'xmlns:BambuStudio="http://schemas.bambulab.com/package/2021"',
            1)

    # Update application version
    new_main = re.sub(
        r'(<metadata\s+name="Application">)[^<]*(</metadata>)',
        r'\1BambuStudio-2.3.2-rc2\2',
        new_main)

    # Write outputs
    zout.writestr("3D/3dmodel.model", new_main.encode("utf-8"))

    for filename, content in object_files.items():
        zout.writestr(f"3D/Objects/{filename}", content)

    zout.writestr("3D/_rels/3dmodel.model.rels",
                  _generate_model_rels(object_filenames).encode("utf-8"))

    # Always write correct Content_Types and root rels
    zout.writestr("[Content_Types].xml", CONTENT_TYPES_XML.encode("utf-8"))
    zout.writestr("_rels/.rels", _generate_root_rels().encode("utf-8"))

    # Copy other files
    skip = {"3D/3dmodel.model", "3D/_rels/3dmodel.model.rels",
            "[Content_Types].xml", "_rels/.rels"}
    for name in zin.namelist():
        if name not in skip and not name.startswith("3D/Objects/"):
            if strip_settings and name.startswith("Metadata/") and name.endswith(".config"):
                continue
            zout.writestr(name, zin.read(name))

    return True


# ---------------------------------------------------------------------------
# PrusaSlicer conversion
# ---------------------------------------------------------------------------

def _stream_convert_prusa(zin, zout, source_name, strip_settings=False):
    """
    Convert PrusaSlicer 3MF to Orca format.
    PrusaSlicer uses the production extension but with different metadata.
    We need to add BambuStudio namespace/metadata and generate model_settings.config.
    """
    raw_bytes = zin.read("3D/3dmodel.model")
    raw_text = raw_bytes.decode("utf-8", errors="replace")

    # Check if it already has p:path (newer PrusaSlicer)
    has_p_path = bool(re.search(r'p:path\s*=', raw_text))

    if has_p_path:
        # Already has production extension — just add Bambu metadata
        new_main = raw_text

        # Add BambuStudio namespace if missing
        if 'xmlns:BambuStudio=' not in new_main[:500]:
            new_main = new_main.replace(
                'xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06"',
                'xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06" '
                'xmlns:BambuStudio="http://schemas.bambulab.com/package/2021"',
                1)

        # Add BambuStudio version metadata if missing
        if 'BambuStudio:3mfVersion' not in new_main:
            new_main = new_main.replace(
                '<resources>',
                '<metadata name="BambuStudio:3mfVersion">1</metadata>\n <resources>',
                1)

        # Update/add application metadata
        if 'name="Application"' in new_main:
            new_main = re.sub(
                r'(<metadata\s+name="Application">)[^<]*(</metadata>)',
                r'\1BambuStudio-2.3.2-rc2\2',
                new_main)
        else:
            new_main = new_main.replace(
                '<resources>',
                '<metadata name="Application">BambuStudio-2.3.2-rc2</metadata>\n <resources>',
                1)

        zout.writestr("3D/3dmodel.model", new_main.encode("utf-8"))

        # Copy all existing files, regenerate Content_Types and root rels
        skip = {"3D/3dmodel.model", "[Content_Types].xml", "_rels/.rels"}
        for name in zin.namelist():
            if name not in skip:
                if strip_settings and "Slic3r" in name:
                    continue
                zout.writestr(name, zin.read(name))

        zout.writestr("[Content_Types].xml", CONTENT_TYPES_XML.encode("utf-8"))
        zout.writestr("_rels/.rels", _generate_root_rels().encode("utf-8"))

        # Generate model_settings.config if missing
        if "Metadata/model_settings.config" not in zin.namelist():
            # Find assembly object ID and parts
            asm_ids = re.findall(r'<object\s+[^>]*?\bid="(\d+)"[^>]*>.*?<components>', raw_text, re.DOTALL)
            if asm_ids:
                asm_id = asm_ids[0]
                settings = f'''<?xml version="1.0" encoding="UTF-8"?>
<config>
  <object id="{asm_id}">
    <metadata key="name" value="{source_name}"/>
    <metadata key="extruder" value="1"/>
  </object>
  <plate>
    <metadata key="plater_id" value="1"/>
    <metadata key="plater_name" value=""/>
    <metadata key="locked" value="false"/>
    <model_instance>
      <metadata key="object_id" value="{asm_id}"/>
      <metadata key="instance_id" value="0"/>
      <metadata key="identify_id" value="1"/>
    </model_instance>
  </plate>
</config>'''
                zout.writestr("Metadata/model_settings.config", settings.encode("utf-8"))

        return True

    else:
        # Older PrusaSlicer without production extension — treat like bare
        return _stream_convert_bare(zin, zout, source_name,
                                    source_name + ".3mf", strip_settings)


# ---------------------------------------------------------------------------
# Main conversion entry point
# ---------------------------------------------------------------------------

def convert_3mf(input_path, output_path=None, force=False, strip_settings=False):
    """
    Convert a 3MF file to Orca-compatible format.

    Args:
        input_path: Path to input 3MF file
        output_path: Path for output (default: input_stem_orca.3mf)
        force: Overwrite existing output
        strip_settings: Remove slicer-specific settings (geometry only)
    """
    input_path = Path(input_path)
    if output_path is None:
        output_path = input_path.parent / f"{input_path.stem}_orca.3mf"
    else:
        output_path = Path(output_path)

    if output_path.exists() and not force:
        print(f"Output file already exists: {output_path}")
        print("Use --force to overwrite.")
        return False

    with zipfile.ZipFile(input_path, "r") as zin:
        names = zin.namelist()

        if "3D/3dmodel.model" not in names:
            # Check if it's a sliced file
            has_gcode = any(n.endswith(".gcode") for n in names)
            if has_gcode:
                print(f"This is a sliced .gcode.3mf file with no model geometry: {input_path}")
                print("This file is meant for direct printing, not for opening in a slicer.")
                return False
            print(f"Not a valid 3MF (no 3D/3dmodel.model): {input_path}")
            return False

        # Read header + tail for classification
        model_size = zin.getinfo("3D/3dmodel.model").file_size
        with zin.open("3D/3dmodel.model") as f:
            if model_size <= 10 * 1024 * 1024:
                classify_bytes = f.read()
            else:
                head = f.read(50000)
                tail = b""
                while True:
                    chunk = f.read(1024 * 1024)
                    if not chunk:
                        break
                    tail = chunk[-50000:] if len(chunk) >= 50000 else (tail + chunk)[-50000:]
                classify_bytes = head + tail

        fmt = classify_3mf(classify_bytes, zip_names=names)
        print(f"Detected format: {fmt}")

        if fmt == "orca_ready":
            if strip_settings:
                # Even if already Orca-ready, strip settings if requested
                pass  # fall through to conversion
            else:
                print(f"Already in Orca format: {input_path}")
                if input_path != output_path:
                    shutil.copy2(input_path, output_path)
                return True

        if fmt == "sliced":
            print(f"This is a sliced .gcode.3mf file: {input_path}")
            print("It contains G-code for printing but no model geometry to convert.")
            print("You need the original (unsliced) 3MF or STL file instead.")
            return False

        if fmt == "unknown":
            print(f"Unrecognized 3MF structure: {input_path}")
            return False

        source_name = re.sub(r'[^\w]', '_', input_path.stem) or "model"

        try:
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
                if fmt == "bare":
                    success = _stream_convert_bare(zin, zout, source_name,
                                                   input_path.name, strip_settings)
                elif fmt == "prusa":
                    success = _stream_convert_prusa(zin, zout, source_name, strip_settings)
                elif fmt == "orca_ready" and strip_settings:
                    success = _strip_settings_only(zin, zout)
                else:
                    success = _stream_convert_bambu_old(zin, zout, source_name, strip_settings)
        except Exception as e:
            print(f"Error during conversion: {e}")
            output_path.unlink(missing_ok=True)
            raise

        if success:
            print(f"Converted: {input_path}")
            print(f"Output:    {output_path}")
        else:
            output_path.unlink(missing_ok=True)
            print(f"Conversion failed: {input_path}")

        return success


def _strip_settings_only(zin, zout):
    """Strip slicer settings from an already Orca-ready file."""
    for name in zin.namelist():
        if name == "Metadata/project_settings.config":
            zout.writestr(name, b"")
        elif name.endswith(".gcode"):
            continue  # skip embedded gcode
        elif "Slic3r" in name or "slic3r" in name:
            continue
        else:
            zout.writestr(name, zin.read(name))
    return True


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Convert 3MF files to current Orca Slicer format"
    )
    parser.add_argument("input", nargs="+", help="Input 3MF file(s)")
    parser.add_argument("-o", "--output", help="Output file (only valid with single input)")
    parser.add_argument("-f", "--force", action="store_true", help="Overwrite existing output files")
    parser.add_argument("--in-place", action="store_true",
                        help="Convert files in place (overwrites originals)")
    parser.add_argument("--suffix", default="_orca",
                        help="Suffix for output files (default: _orca)")
    parser.add_argument("--strip-settings", action="store_true",
                        help="Remove slicer-specific settings (geometry + colors only)")

    args = parser.parse_args()

    if args.output and len(args.input) > 1:
        print("Error: --output can only be used with a single input file", file=sys.stderr)
        sys.exit(1)

    success_count = 0
    for input_file in args.input:
        input_path = Path(input_file)
        if not input_path.exists():
            print(f"File not found: {input_file}", file=sys.stderr)
            continue
        if input_path.suffix.lower() != ".3mf":
            print(f"Not a 3MF file: {input_file}", file=sys.stderr)
            continue

        if args.output:
            output_path = Path(args.output)
        elif args.in_place:
            output_path = input_path
            args.force = True
        else:
            output_path = input_path.parent / f"{input_path.stem}{args.suffix}.3mf"

        if output_path == input_path:
            with tempfile.NamedTemporaryFile(suffix=".3mf", delete=False,
                                             dir=input_path.parent) as tmp:
                tmp_path = Path(tmp.name)
            if convert_3mf(input_path, tmp_path, force=True,
                           strip_settings=args.strip_settings):
                tmp_path.replace(input_path)
                success_count += 1
            else:
                tmp_path.unlink(missing_ok=True)
        else:
            if convert_3mf(input_path, output_path, force=args.force,
                           strip_settings=args.strip_settings):
                success_count += 1

    if len(args.input) > 1:
        print(f"\nConverted {success_count}/{len(args.input)} files.")

    sys.exit(0 if success_count > 0 else 1)


if __name__ == "__main__":
    main()
