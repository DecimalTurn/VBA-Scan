export const analyzePython: string = `
import json
import os
import re
import zipfile
from oletools.olevba import VBA_Parser, VBA_Project

def get_project_codepages(parser):
    codepages = []
    parsers = [parser]
    while parsers:
        current = parsers.pop()
        parsers.extend(getattr(current, "ole_subfiles", []) or [])
        if current.ole_file is None:
            continue
        try:
            projects = current.find_vba_projects() or []
        except Exception:
            continue
        for vba_root, project_path, dir_path in projects:
            try:
                project = VBA_Project(
                    current.ole_file, vba_root, project_path, dir_path, current.relaxed
                )
                codepages.append({
                    "project": str(project.projectname),
                    "value": project.codepage,
                    "name": project.codepage_name,
                })
            except Exception:
                continue
    return codepages

def scan_document(path, pcode_path, vba_source_path, vba_archive_path, xlm_path):
    parser = VBA_Parser(path)
    try:
        has_vba = parser.detect_vba_macros()
        has_xlm = parser.detect_xlm_macros()
        stomping = None
        stomping_error = None
        stomping_items = []
        try:
            stomping = parser.detect_vba_stomping()
        except Exception as error:
            stomping_error = str(error)
        if stomping:
          stomping_items = getattr(parser, "vba_stomping_items", [])

        pcode_size = 0
        pcode_error = None
        pcode = ""
        vba_source_size = 0
        vba_source_archive_size = 0
        vba_source_module_count = 0
        vba_source_error = None
        xlm_size = 0
        xlm_error = None
        xlm_sources = []
        xlm_findings = []
        source_module_names = {}
        if has_vba or stomping:
            try:
                pcode = parser.extract_pcode() or ""
                if pcode:
                    with open(pcode_path, "w", encoding="utf-8") as pcode_file:
                        pcode_file.write(pcode)
                    pcode_size = os.path.getsize(pcode_path)
            except Exception as error:
                pcode_error = str(error)

        if has_vba:
            try:
                used_names = set()
                source_parts = []
                with zipfile.ZipFile(vba_archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                    for index, (_, stream_path, vba_filename, vba_code) in enumerate(parser.extract_macros(), start=1):
                        if stream_path == "VBA P-code" and vba_filename == "VBA_P-code.txt":
                            continue
                        if stream_path == "xlm_macro" and vba_filename == "xlm_macro.txt":
                            continue
                        filename = str(vba_filename).replace("\\\\", "/").rsplit("/", 1)[-1]
                        if filename in ("", ".", ".."):
                            filename = f"Module-{index}.bas"
                        filename = re.sub(r'[<>:"/|?*\\x00-\\x1f]', "_", filename)
                        stem, extension = os.path.splitext(filename)
                        candidate = filename
                        suffix = 2
                        while candidate.casefold() in used_names:
                            candidate = f"{stem}-{suffix}{extension}"
                            suffix += 1
                        used_names.add(candidate.casefold())
                        source_module_names[index - 1] = candidate
                        archive.writestr(candidate, vba_code)
                        source_parts.append(f"' ===== {candidate} =====\\n{vba_code}\\n")
                        vba_source_module_count += 1
                if vba_source_module_count:
                    with open(vba_source_path, "w", encoding="utf-8", newline="") as source_file:
                        source_file.write("\\n".join(source_parts))
                    vba_source_size = os.path.getsize(vba_source_path)
                    vba_source_archive_size = os.path.getsize(vba_archive_path)
                else:
                    os.unlink(vba_archive_path)
            except Exception as error:
                vba_source_error = str(error)
                vba_source_module_count = 0
                if os.path.exists(vba_archive_path):
                    os.unlink(vba_archive_path)
                if os.path.exists(vba_source_path):
                    os.unlink(vba_source_path)

        if has_xlm:
          try:
            xlm_sources = parser.extract_xlm_macros() or []
            xlm_findings = [
              {"type": kind, "keyword": keyword, "description": description}
              for kind, keyword, description in parser.analyze_xlm_macros()
            ]
            if xlm_sources:
              with open(xlm_path, "w", encoding="utf-8", newline="") as xlm_file:
                for index, (location, code) in enumerate(xlm_sources):
                  if index > 0:
                    xlm_file.write("\\n")
                  xlm_file.write(f"===== {location} =====\\n")
                  xlm_file.write(code.rstrip() + "\\n")
              xlm_size = os.path.getsize(xlm_path)
          except Exception as error:
            xlm_error = str(error)

        findings_by_source = parser.analyze_macros_by_source() or {}
        vba_source_findings = [
          {"type": kind, "keyword": keyword, "description": description}
          for kind, keyword, description in findings_by_source.get("vba_source", [])
        ]
        pcode_findings = [
          {"type": kind, "keyword": keyword, "description": description}
          for kind, keyword, description in findings_by_source.get("pcode", [])
        ]
        vba_source_locations = []
        for location in findings_by_source.get("vba_source_locations", []):
            mapped_location = dict(location)
            mapped_location["module_name"] = source_module_names.get(
                location["module_index"], location["module_name"]
            )
            vba_source_locations.append(mapped_location)
        return json.dumps({
            "has_vba": has_vba,
            "has_xlm": has_xlm,
            "stomping": stomping,
            "stomping_error": stomping_error,
            "stomping_items": stomping_items,
            "pcode_size": pcode_size,
            "pcode_error": pcode_error,
            "vba_source_size": vba_source_size,
            "vba_source_archive_size": vba_source_archive_size,
            "vba_source_module_count": vba_source_module_count,
            "vba_source_error": vba_source_error,
            "xlm_size": xlm_size,
            "xlm_error": xlm_error,
            "xlm_macro_count": len(xlm_sources),
            "project_codepages": get_project_codepages(parser),
            "vba_source_findings": vba_source_findings,
            "vba_source_locations": vba_source_locations,
            "pcode_findings": pcode_findings,
            "xlm_findings": xlm_findings,
        })
    finally:
        parser.close()

`;
