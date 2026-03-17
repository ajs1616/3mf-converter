#!/usr/bin/env python3
"""
Convert old-format BambuStudio 3MF files to the current Orca Slicer format.

Old format: all mesh data inline in 3D/3dmodel.model, no production extension.
New format: mesh data in 3D/Objects/, component refs with p:path and p:UUID,
            3D/_rels/3dmodel.model.rels, production extension namespace.
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
NS_3DMODEL_TYPE = "http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel"

# Register namespaces so ET doesn't mangle them with ns0/ns1 prefixes
ET.register_namespace("", NS_CORE)
ET.register_namespace("p", NS_PRODUCTION)
ET.register_namespace("BambuStudio", NS_BAMBU)


def make_uuid():
    return str(uuid.uuid4())


def serialize_xml(root):
    """Serialize an Element to clean UTF-8 XML bytes, fixing ET quirks."""
    ET.indent(ET.ElementTree(root), space=" ")
    raw = ET.tostring(root, encoding="unicode", xml_declaration=False)

    # Remove duplicate namespace declarations that ET sometimes produces
    # e.g. xmlns:p="..." appearing twice
    seen_decls = set()

    def dedup_xmlns(match):
        decl = match.group(0)
        if decl in seen_decls:
            return ""
        seen_decls.add(decl)
        return decl

    raw = re.sub(r'xmlns(?::\w+)?="[^"]*"', dedup_xmlns, raw)
    # Clean up any double spaces left by removed duplicates
    raw = re.sub(r'  +', ' ', raw)
    # Clean up space before >
    raw = re.sub(r' >', '>', raw)
    raw = re.sub(r' />', ' />', raw)

    xml_decl = '<?xml version="1.0" encoding="UTF-8"?>\n'
    return (xml_decl + raw).encode("utf-8")


def needs_conversion(model_xml):
    """Check if this 3MF uses the old inline format (no production extension)."""
    root = ET.fromstring(model_xml)
    resources = root.find(f"{{{NS_CORE}}}resources")
    if resources is None:
        return False
    for obj in resources.findall(f"{{{NS_CORE}}}object"):
        components = obj.find(f"{{{NS_CORE}}}components")
        if components is not None:
            for comp in components.findall(f"{{{NS_CORE}}}component"):
                if comp.get(f"{{{NS_PRODUCTION}}}path"):
                    return False
        mesh = obj.find(f"{{{NS_CORE}}}mesh")
        if mesh is not None:
            for other_obj in resources.findall(f"{{{NS_CORE}}}object"):
                other_comps = other_obj.find(f"{{{NS_CORE}}}components")
                if other_comps is not None:
                    for c in other_comps.findall(f"{{{NS_CORE}}}component"):
                        if c.get("objectid") == obj.get("id"):
                            return True
    return False


def convert_model(model_xml, source_name="object"):
    """
    Convert old-format 3dmodel.model to new format.

    Returns:
        (new_main_xml_bytes, dict of {filename: xml_bytes} for 3D/Objects/)
    """
    root = ET.fromstring(model_xml)

    # Ensure production namespace and requiredextensions
    if "requiredextensions" not in root.attrib:
        root.set("requiredextensions", "p")

    resources = root.find(f"{{{NS_CORE}}}resources")
    build = root.find(f"{{{NS_CORE}}}build")

    # Categorize objects
    mesh_objects = {}
    assembly_objects = {}

    for obj in list(resources.findall(f"{{{NS_CORE}}}object")):
        obj_id = obj.get("id")
        if obj.find(f"{{{NS_CORE}}}mesh") is not None:
            mesh_objects[obj_id] = obj
        if obj.find(f"{{{NS_CORE}}}components") is not None:
            assembly_objects[obj_id] = obj

    if not mesh_objects:
        return model_xml, {}

    object_files = {}

    for asm_id, asm_obj in assembly_objects.items():
        components = asm_obj.find(f"{{{NS_CORE}}}components")
        asm_uuid = f"0000000{asm_id}-61cb-4c03-9d28-80fed5dfa1dc"
        asm_obj.set(f"{{{NS_PRODUCTION}}}UUID", asm_uuid)

        for comp in components.findall(f"{{{NS_CORE}}}component"):
            ref_id = comp.get("objectid")
            if ref_id not in mesh_objects:
                continue

            mesh_obj = mesh_objects[ref_id]
            obj_filename = f"{source_name}_{ref_id}.model"
            obj_path = f"/3D/Objects/{obj_filename}"
            part_uuid = f"000{ref_id}0000-b206-40ff-9872-83e8017abed1"

            # Build the sub-model XML
            sub_root = ET.Element(f"{{{NS_CORE}}}model")
            sub_root.set("unit", "millimeter")
            sub_root.set("xml:lang", "en-US")
            sub_root.set("requiredextensions", "p")

            sub_meta = ET.SubElement(sub_root, f"{{{NS_CORE}}}metadata")
            sub_meta.set("name", "BambuStudio:3mfVersion")
            sub_meta.text = "1"

            sub_resources = ET.SubElement(sub_root, f"{{{NS_CORE}}}resources")
            sub_obj = ET.SubElement(sub_resources, f"{{{NS_CORE}}}object")
            sub_obj.set("id", ref_id)
            sub_obj.set(f"{{{NS_PRODUCTION}}}UUID", part_uuid)
            sub_obj.set("type", "model")

            # Move mesh element to sub-model
            mesh_elem = mesh_obj.find(f"{{{NS_CORE}}}mesh")
            sub_obj.append(mesh_elem)

            object_files[obj_filename] = serialize_xml(sub_root)

            # Update component reference
            comp.set(f"{{{NS_PRODUCTION}}}path", obj_path)
            comp.set(f"{{{NS_PRODUCTION}}}UUID", part_uuid)

            # Remove old inline mesh object from main model
            resources.remove(mesh_obj)

    # Add UUIDs to build
    if build is not None:
        build.set(f"{{{NS_PRODUCTION}}}UUID", make_uuid())
        for item in build.findall(f"{{{NS_CORE}}}item"):
            item.set(f"{{{NS_PRODUCTION}}}UUID", make_uuid())

    # Update application version
    for meta in root.findall(f"{{{NS_CORE}}}metadata"):
        if meta.get("name") == "Application":
            meta.text = "BambuStudio-2.3.2-rc2"

    # Remove thumbnail metadata not used by Orca
    for meta in list(root.findall(f"{{{NS_CORE}}}metadata")):
        name = meta.get("name", "")
        if name in ("Thumbnail_Middle", "Thumbnail_Small"):
            root.remove(meta)

    return serialize_xml(root), object_files


def create_model_rels(object_files):
    """Create 3D/_rels/3dmodel.model.rels content."""
    root = ET.Element(f"{{{NS_RELS}}}Relationships")
    for i, filename in enumerate(object_files.keys(), 1):
        rel = ET.SubElement(root, f"{{{NS_RELS}}}Relationship")
        rel.set("Target", f"/3D/Objects/{filename}")
        rel.set("Id", f"rel-{i}")
        rel.set("Type", NS_3DMODEL_TYPE)

    # Register rels namespace for clean output
    ET.register_namespace("", NS_RELS)
    result = serialize_xml(root)
    # Re-register core namespace as default
    ET.register_namespace("", NS_CORE)
    return result


def convert_3mf(input_path, output_path=None, force=False):
    """Convert a 3MF file from old format to new Orca-compatible format."""
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
        model_xml = zin.read("3D/3dmodel.model")

        has_objects_dir = any(n.startswith("3D/Objects/") for n in names)
        has_model_rels = "3D/_rels/3dmodel.model.rels" in names

        if has_objects_dir and has_model_rels and not needs_conversion(model_xml):
            print(f"Already in new format, no conversion needed: {input_path}")
            if input_path != output_path:
                shutil.copy2(input_path, output_path)
            return True

        source_name = input_path.stem.replace(" ", "_").replace("-", "_")
        new_main_xml, object_files = convert_model(model_xml, source_name)

        if not object_files:
            print(f"No mesh objects to extract. File may have unusual structure.")
            return False

        model_rels_xml = create_model_rels(object_files)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name == "3D/3dmodel.model":
                    zout.writestr(name, new_main_xml)
                elif name.startswith("3D/Objects/") or name == "3D/_rels/3dmodel.model.rels":
                    continue
                else:
                    zout.writestr(name, zin.read(name))

            for filename, content in object_files.items():
                zout.writestr(f"3D/Objects/{filename}", content)
            zout.writestr("3D/_rels/3dmodel.model.rels", model_rels_xml)

    print(f"Converted: {input_path}")
    print(f"Output:    {output_path}")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Convert old BambuStudio 3MF files to current Orca Slicer format"
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
