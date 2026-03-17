#!/usr/bin/env python3
"""
Convert 3MF files to the current Orca Slicer format.

Handles:
1. Old BambuStudio format (inline mesh + assembly, no p:path)
2. Bare 3MF files (mesh objects, no BambuStudio wrapper)
3. Already-Orca files (detected and skipped)

Uses streaming XML for large files — never loads full mesh data into a DOM.
"""

import argparse
import io
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
NS_3DMODEL_TYPE = "http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel"

ET.register_namespace("", NS_CORE)
ET.register_namespace("p", NS_PRODUCTION)
ET.register_namespace("BambuStudio", NS_BAMBU)


def make_uuid():
    return str(uuid.uuid4())


def serialize_xml(root):
    """Serialize an Element to clean UTF-8 XML bytes."""
    ET.indent(ET.ElementTree(root), space=" ")
    raw = ET.tostring(root, encoding="unicode", xml_declaration=False)
    seen = set()
    def dedup(m):
        d = m.group(0)
        if d in seen: return ""
        seen.add(d)
        return d
    raw = re.sub(r'xmlns(?::\w+)?="[^"]*"', dedup, raw)
    raw = re.sub(r'  +', ' ', raw)
    raw = re.sub(r' >', '>', raw)
    raw = re.sub(r' />', ' />', raw)
    return ('<?xml version="1.0" encoding="UTF-8"?>\n' + raw).encode("utf-8")


def classify_3mf(model_data):
    """
    Classify a 3MF model's format.
    Reads the first 50KB and last 50KB to catch structure tags
    that may appear after large mesh blocks.

    Accepts bytes or string.
    """
    if isinstance(model_data, bytes):
        header = model_data[:50000].decode("utf-8", errors="replace")
        tail = model_data[-50000:].decode("utf-8", errors="replace") if len(model_data) > 50000 else ""
    else:
        header = model_data[:50000]
        tail = model_data[-50000:] if len(model_data) > 50000 else ""

    combined = header + tail

    has_mesh = bool(re.search(r'<mesh', combined))
    has_component = bool(re.search(r'<component\b', combined))
    has_p_path = bool(re.search(r'p:path\s*=', combined))
    has_object = bool(re.search(r'<object\b', combined))

    if has_p_path:
        return "orca_ready"
    if has_component:
        return "bambu_old"
    if has_object and has_mesh:
        return "bare"
    if has_object:
        return "bare"
    return "unknown"


def needs_conversion(model_xml):
    """Check if this 3MF needs conversion."""
    return classify_3mf(model_xml) in ("bambu_old", "bare")


def _extract_objects_from_header(header_text):
    """
    Parse object IDs, types, and whether they have mesh/components
    from the XML header text (first ~50KB).
    """
    objects = []
    for m in re.finditer(r'<object\s+([^>]+?)/?>', header_text, re.DOTALL):
        attrs_str = m.group(1)
        obj_id = re.search(r'\bid="(\d+)"', attrs_str)
        if obj_id:
            objects.append({
                "id": obj_id.group(1),
                "start": m.start(),
                "attrs": attrs_str,
            })
    return objects


def _stream_convert_bare(zin, zout, source_name, input_name):
    """
    Convert a bare 3MF using streaming — reads the raw model bytes line-by-line,
    wraps the mesh in a sub-model file, and creates the Orca assembly structure.
    """
    # Read the model XML as raw bytes from the zip
    with zin.open("3D/3dmodel.model") as f:
        raw_bytes = f.read()

    raw_text = raw_bytes.decode("utf-8", errors="replace")

    # Find object IDs from the XML
    # For bare 3MF, typically just one object with a mesh
    obj_ids = re.findall(r'<object\s+[^>]*\bid="(\d+)"', raw_text[:50000])
    if not obj_ids:
        return False

    # We'll use the first object as our mesh object
    mesh_id = obj_ids[0]
    asm_id = str(max(int(x) for x in obj_ids) + 1)

    obj_filename = f"{source_name}_{mesh_id}.model"
    part_uuid = make_uuid()
    asm_uuid = make_uuid()
    build_uuid = make_uuid()
    item_uuid = make_uuid()

    # --- Write the sub-model file ---
    # We need to wrap the existing mesh in a new model envelope
    # Extract the <mesh>...</mesh> block from the raw text using byte offsets for speed

    # Find <mesh> and </mesh> positions
    mesh_start = raw_text.find("<mesh>")
    if mesh_start == -1:
        mesh_start = raw_text.find("<mesh ")
    mesh_end = raw_text.rfind("</mesh>")

    if mesh_start == -1 or mesh_end == -1:
        print("Could not find <mesh> block in model")
        return False

    mesh_end += len("</mesh>")
    mesh_block = raw_text[mesh_start:mesh_end]

    sub_model = f'''<?xml version="1.0" encoding="UTF-8"?>
<model xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02" xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06" unit="millimeter" xml:lang="en-US" requiredextensions="p">
 <metadata name="BambuStudio:3mfVersion">1</metadata>
 <resources>
  <object id="{mesh_id}" p:UUID="{part_uuid}" type="model">
   {mesh_block}
  </object>
 </resources>
</model>'''

    # Write sub-model to zip (encode as bytes)
    zout.writestr(f"3D/Objects/{obj_filename}", sub_model.encode("utf-8"))

    # --- Write the main model ---
    main_model = f'''<?xml version="1.0" encoding="UTF-8"?>
<model xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02" xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06" xmlns:BambuStudio="http://schemas.bambulab.com/package/2021" unit="millimeter" xml:lang="en-US" requiredextensions="p">
 <metadata name="Application">BambuStudio-2.3.2-rc2</metadata>
 <metadata name="BambuStudio:3mfVersion">1</metadata>
 <resources>
  <object id="{asm_id}" p:UUID="{asm_uuid}" type="model">
   <components>
    <component p:path="/3D/Objects/{obj_filename}" objectid="{mesh_id}" p:UUID="{part_uuid}" transform="1 0 0 0 1 0 0 0 1 0 0 0"/>
   </components>
  </object>
 </resources>
 <build p:UUID="{build_uuid}">
  <item objectid="{asm_id}" p:UUID="{item_uuid}" printable="1"/>
 </build>
</model>'''

    zout.writestr("3D/3dmodel.model", main_model.encode("utf-8"))

    # --- Write model rels ---
    rels = f'''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Target="/3D/Objects/{obj_filename}" Id="rel-1" Type="http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel"/>
</Relationships>'''
    zout.writestr("3D/_rels/3dmodel.model.rels", rels.encode("utf-8"))

    # --- Write model_settings.config ---
    settings = f'''<?xml version="1.0" encoding="UTF-8"?>
<config>
  <object id="{asm_id}">
    <metadata key="name" value="{source_name}"/>
    <metadata key="extruder" value="1"/>
    <part id="{mesh_id}" subtype="normal_part">
      <metadata key="name" value="{source_name}"/>
      <metadata key="matrix" value="1 0 0 0 0 1 0 0 0 0 1 0 0 0 0 1"/>
      <metadata key="source_file" value="{input_name}"/>
      <metadata key="source_object_id" value="0"/>
      <metadata key="source_volume_id" value="0"/>
      <metadata key="source_offset_x" value="0"/>
      <metadata key="source_offset_y" value="0"/>
      <metadata key="source_offset_z" value="0"/>
      <mesh_stat edges_fixed="0" degenerate_facets="0" facets_removed="0" facets_reversed="0" backwards_edges="0"/>
    </part>
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
    zout.writestr("Metadata/project_settings.config", b"")

    # Copy other files from input (Content_Types, rels, etc.) except what we replaced
    skip = {"3D/3dmodel.model", "3D/_rels/3dmodel.model.rels",
            "Metadata/model_settings.config", "Metadata/project_settings.config"}
    for name in zin.namelist():
        if name not in skip and not name.startswith("3D/Objects/"):
            zout.writestr(name, zin.read(name))

    return True


def _stream_convert_bambu_old(zin, zout, source_name):
    """
    Convert old BambuStudio format using streaming.
    The model has assembly + inline mesh objects.
    """
    raw_bytes = zin.read("3D/3dmodel.model")
    raw_text = raw_bytes.decode("utf-8", errors="replace")

    # Find object blocks
    # We need to:
    # 1. Find mesh objects (have <mesh>)
    # 2. Find assembly objects (have <components>)
    # 3. Extract mesh objects to sub-files
    # 4. Update assembly to reference sub-files

    # Find all <object ...> positions and their content type
    obj_pattern = re.compile(r'<object\s+([^>]*?)>(.*?)</object>', re.DOTALL)
    mesh_objects = {}  # id -> (full_match, mesh_block)
    assembly_info = {}  # id -> (full_match, components)

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

    # Process assembly objects
    new_main = raw_text
    for asm_id, asm in assembly_info.items():
        asm_uuid = make_uuid()

        # Find components in this assembly
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

            # Build sub-model
            mesh_block = mesh_objects[ref_id]["mesh_block"]
            sub_model = f'''<?xml version="1.0" encoding="UTF-8"?>
<model xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02" xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06" unit="millimeter" xml:lang="en-US" requiredextensions="p">
 <metadata name="BambuStudio:3mfVersion">1</metadata>
 <resources>
  <object id="{ref_id}" p:UUID="{part_uuid}" type="model">
   {mesh_block}
  </object>
 </resources>
</model>'''
            object_files[obj_filename] = sub_model.encode("utf-8")

            # Update component to add p:path and p:UUID
            old_comp = cm.group(0)
            new_comp = old_comp.replace("/>", f' p:path="/3D/Objects/{obj_filename}" p:UUID="{part_uuid}"/>')
            new_body = new_body.replace(cm.group(0), new_comp, 1)

        # Update assembly object attrs to add p:UUID
        old_asm_obj = asm["match"].group(0)
        new_attrs = asm["attrs"]
        if "p:UUID" not in new_attrs:
            new_attrs = new_attrs + f' p:UUID="{asm_uuid}"'
        new_asm_obj = f'<object {new_attrs}>{new_body}</object>'
        new_main = new_main.replace(old_asm_obj, new_asm_obj, 1)

    # Remove mesh objects from main model (they're now in sub-files)
    for mesh_id, mesh_info in mesh_objects.items():
        # Only remove if it's referenced by an assembly
        is_referenced = any(
            re.search(rf'objectid="{mesh_id}"', asm["body"])
            for asm in assembly_info.values()
        )
        if is_referenced:
            new_main = new_main.replace(mesh_info["match"].group(0), "", 1)

    # Add UUIDs to build items
    def add_build_uuids(m):
        item_str = m.group(0)
        if "p:UUID" not in item_str:
            item_str = item_str.replace("/>", f' p:UUID="{make_uuid()}"/>')
            item_str = item_str.replace(">", f' p:UUID="{make_uuid()}">', 1) if "/>" not in item_str else item_str
        return item_str

    new_main = re.sub(r'<item\s+[^>]*/>', add_build_uuids, new_main)

    # Add UUID to build element
    def add_build_uuid(m):
        build_tag = m.group(0)
        if "p:UUID" not in build_tag:
            build_tag = build_tag.replace(">", f' p:UUID="{make_uuid()}">', 1)
        return build_tag

    new_main = re.sub(r'<build[^>]*>', add_build_uuid, new_main)

    # Add requiredextensions if missing
    if 'requiredextensions=' not in new_main[:500]:
        new_main = new_main.replace('<model ', '<model requiredextensions="p" ', 1)

    # Add p namespace if missing
    if 'xmlns:p=' not in new_main[:500]:
        new_main = new_main.replace(
            'xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02"',
            'xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02" xmlns:p="http://schemas.microsoft.com/3dmanufacturing/production/2015/06"',
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

    # Write model rels
    rels_entries = []
    for i, filename in enumerate(object_files.keys(), 1):
        rels_entries.append(f' <Relationship Target="/3D/Objects/{filename}" Id="rel-{i}" Type="{NS_3DMODEL_TYPE}"/>')
    rels = '<?xml version="1.0" encoding="UTF-8"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n' + "\n".join(rels_entries) + '\n</Relationships>'
    zout.writestr("3D/_rels/3dmodel.model.rels", rels.encode("utf-8"))

    # Copy other files
    skip = {"3D/3dmodel.model", "3D/_rels/3dmodel.model.rels"}
    for name in zin.namelist():
        if name not in skip and not name.startswith("3D/Objects/"):
            zout.writestr(name, zin.read(name))

    return True


def convert_3mf(input_path, output_path=None, force=False):
    """Convert a 3MF file to Orca-compatible format."""
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
            print(f"Not a valid 3MF (no 3D/3dmodel.model): {input_path}")
            return False

        # Read header + tail to classify.
        # Component/build tags often appear after huge mesh blocks,
        # so we need the tail too. For files under 10MB, just read it all.
        # For larger files, read in chunks keeping only head+tail.
        model_size = zin.getinfo("3D/3dmodel.model").file_size
        with zin.open("3D/3dmodel.model") as f:
            if model_size <= 10 * 1024 * 1024:  # under 10MB
                classify_bytes = f.read()
            else:
                # Stream through, keeping first 50KB and last 50KB
                head = f.read(50000)
                tail = b""
                while True:
                    chunk = f.read(1024 * 1024)  # 1MB chunks
                    if not chunk:
                        break
                    tail = chunk[-50000:] if len(chunk) >= 50000 else (tail + chunk)[-50000:]
                classify_bytes = head + tail

        fmt = classify_3mf(classify_bytes)
        print(f"Detected format: {fmt}")

        if fmt == "orca_ready":
            print(f"Already in Orca format: {input_path}")
            if input_path != output_path:
                shutil.copy2(input_path, output_path)
            return True

        if fmt == "unknown":
            print(f"Unrecognized 3MF structure: {input_path}")
            return False

        source_name = input_path.stem.replace(" ", "_").replace("-", "_")

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            if fmt == "bare":
                success = _stream_convert_bare(zin, zout, source_name, input_path.name)
            else:
                success = _stream_convert_bambu_old(zin, zout, source_name)

        if success:
            print(f"Converted: {input_path}")
            print(f"Output:    {output_path}")
        else:
            # Clean up failed output
            output_path.unlink(missing_ok=True)
            print(f"Conversion failed: {input_path}")

        return success


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
            if convert_3mf(input_path, tmp_path, force=True):
                tmp_path.replace(input_path)
                success_count += 1
            else:
                tmp_path.unlink(missing_ok=True)
        else:
            if convert_3mf(input_path, output_path, force=args.force):
                success_count += 1

    if len(args.input) > 1:
        print(f"\nConverted {success_count}/{len(args.input)} files.")

    sys.exit(0 if success_count > 0 else 1)


if __name__ == "__main__":
    main()
