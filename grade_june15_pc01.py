# -*- coding: utf-8 -*-
"""
Autodesk Inventor COM API 기반 자동 채점 스크립트 (grade_june15_pc01.py)
- PC-13 폴더를 기준(Reference)으로 삼아 PC-01 폴더의 제출물을 정밀 검사합니다.
- 임시 비번호 처리: 폴더 내 파일명에서 감지된 임시 비번호(예: 11번)를 기준으로 파일명 명명 규칙을 유연하게 검사합니다.
"""

import os
import sys
import re
import json
import collections
import traceback
from datetime import datetime

# 한글 출력 보장을 위한 설정
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

REF_DIR = r"D:\이갑종\26 3학년 전일제\6월15 TEST\PC-13_10.128.56.103"
TARGET_DIR = r"D:\이갑종\26 3학년 전일제\6월15 TEST\PC-01_10.128.56.91"

def get_inventor_app(visible=False):
    """실행 중인 Inventor를 찾거나 새로 시작합니다."""
    import win32com.client
    try:
        app = win32com.client.gencache.EnsureDispatch("Inventor.Application")
        try:
            # 작동 중인지 확인
            _ = app.Documents.Count
            print("[정보] 실행 중인 Autodesk Inventor 연결 성공.")
            return app, False
        except Exception:
            pass
    except Exception:
        pass

    # 새로 기동
    try:
        app = win32com.client.gencache.EnsureDispatch("Inventor.Application")
        app.Visible = visible
        print("[정보] Autodesk Inventor 새로 시작 성공.")
        return app, True
    except Exception as e:
        print(f"[오류] Inventor를 시작할 수 없습니다. COM API 확인 요망: {e}")
        sys.exit(1)

def guess_part_number(filepath):
    """파일명에서 Part 1 또는 Part 2인지 판별합니다."""
    name = os.path.basename(filepath).lower()
    if '01' in name or '_1' in name or 'part1' in name or '-1' in name:
        return 1
    if '02' in name or '_2' in name or 'part2' in name or '-2' in name:
        return 2
    return None

def extract_assembly_info(app, file_path):
    """어셈블리(.iam) 파일을 분석하여 각 부품의 체적(mm³)과 간섭, 누락 파일 목록을 추출합니다."""
    doc = None
    try:
        # 인벤터 경고창 비활성화 시도 (COM 오류 방지)
        try:
            app.SilentFiles = True
        except Exception:
            pass

        # 문서 열기 (보이지 않게 오픈)
        raw_doc = app.Documents.Open(file_path, OpenVisible=False)
        import win32com.client
        try:
            doc = win32com.client.CastTo(raw_doc, "AssemblyDocument")
        except Exception:
            doc = raw_doc
            
        comp_def = doc.ComponentDefinition
        occurrences = comp_def.Occurrences
        
        parts = {}
        # 1. 각 구성 부품(Occurrences) 체적 분석
        for i in range(1, occurrences.Count + 1):
            occ = occurrences.Item(i)
            if occ.Suppressed:
                continue
            try:
                mass_properties = occ.MassProperties
                vol_cm3 = mass_properties.Volume
                vol_mm3 = vol_cm3 * 1000.0 # cm^3 -> mm^3
                
                occ_doc = occ.Definition.Document
                filepath = occ_doc.FullFileName
                
                part_num = guess_part_number(filepath)
                if part_num is None:
                    part_num = len(parts) + 1
                    
                parts[f"part{part_num}"] = {
                    "name": occ.Name,
                    "filename": os.path.basename(filepath),
                    "volume": round(vol_mm3, 2)
                }
            except Exception as e:
                print(f"[경고] 구성 부품 '{occ.Name}' 속성 추출 실패: {e}")

        # 2. 부품 간 간섭 검사 (Interference Analysis)
        interference_vol_mm3 = 0.0
        has_interference = False
        
        if occurrences.Count >= 2:
            oGroup1 = app.TransientObjects.CreateObjectCollection()
            oGroup2 = app.TransientObjects.CreateObjectCollection()
            
            occ1 = occurrences.Item(1)
            occ2 = occurrences.Item(2)
            
            oGroup1.Add(occ1)
            oGroup2.Add(occ2)
            
            try:
                results = comp_def.AnalyzeInterference(oGroup1, oGroup2)
                if results.Count > 0:
                    has_interference = True
                    for r_idx in range(1, results.Count + 1):
                        interference_vol_mm3 += results.Item(r_idx).Volume * 1000.0
            except Exception as e:
                print(f"[경고] 간섭 분석 중 오류 발생: {e}")
                
        # 3. 문서 참조 상태 확인 (누락 파일 검사)
        unresolved_files = []
        for ref_desc in doc.ReferencedFileDescriptors:
            if ref_desc.ReferenceStatus == 2:  # kMissingReference
                unresolved_files.append(ref_desc.DisplayName)

        return {
            "success": True,
            "parts": parts,
            "has_interference": has_interference,
            "interference_volume": round(interference_vol_mm3, 2),
            "unresolved_files": unresolved_files,
            "components_count": occurrences.Count
        }
        
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        if doc is not None:
            try:
                doc.Close(True)
            except Exception:
                pass

def detect_bibunho_prefix(files):
    """파일명 목록에서 가장 많이 나타나는 비번호 숫자 접두사를 감지합니다."""
    prefixes = []
    for f in files:
        m = re.match(r'^(\d+)', f)
        if m:
            prefixes.append(m.group(1))
    if prefixes:
        counter = collections.Counter(prefixes)
        return counter.most_common(1)[0][0]
    return None

def check_naming_convention(folder_name, files):
    """감지된 비번호(또는 폴더명) 기준으로 파일명 규칙을 검사합니다."""
    # 파일명에서 감지된 비번호 접두사 사용 (임시 비번호 대응)
    detected_bibunho = detect_bibunho_prefix(files)
    bibunho = detected_bibunho if detected_bibunho else folder_name
    
    expected_part1 = f"{bibunho}_01.ipt"
    expected_part2 = f"{bibunho}_02.ipt"
    expected_assembly = f"{bibunho}_03.iam"
    expected_stl = f"{bibunho}_03.stl"
    
    status = {
        "part1": {"found": False, "exact_match": False, "filename": ""},
        "part2": {"found": False, "exact_match": False, "filename": ""},
        "assembly": {"found": False, "exact_match": False, "filename": ""},
        "stl": {"found": False, "exact_match": False, "filename": ""},
        "slicing": {"found": False, "exact_match": False, "filename": ""},
        "extra_files": [],
        "used_bibunho": bibunho,
        "is_temporary": (bibunho != folder_name)
    }
    
    matched_files = set()
    
    # 1. 단품 1 매칭
    for f in files:
        fl = f.lower()
        if fl == expected_part1.lower():
            status["part1"] = {"found": True, "exact_match": True, "filename": f}
            matched_files.add(f)
            break
    if not status["part1"]["found"]:
        for f in files:
            fl = f.lower()
            if fl.endswith(".ipt") and (f"_{bibunho}_01" in fl or f"{bibunho}_01" in fl or f"{bibunho}-01" in fl or f"_{bibunho}-01" in fl or fl.startswith(f"{bibunho}_01") or fl.startswith(f"{bibunho}-01")):
                status["part1"] = {"found": True, "exact_match": False, "filename": f}
                matched_files.add(f)
                break
            # 아주 단순 패턴 매치 보완
            elif fl.endswith(".ipt") and (f"_{bibunho}_1" in fl or f"{bibunho}_1" in fl or f"{bibunho}-1" in fl):
                status["part1"] = {"found": True, "exact_match": False, "filename": f}
                matched_files.add(f)
                break
                
    # 2. 단품 2 매칭
    for f in files:
        if f in matched_files:
            continue
        fl = f.lower()
        if fl == expected_part2.lower():
            status["part2"] = {"found": True, "exact_match": True, "filename": f}
            matched_files.add(f)
            break
    if not status["part2"]["found"]:
        for f in files:
            if f in matched_files:
                continue
            fl = f.lower()
            if fl.endswith(".ipt") and (f"{bibunho}_02" in fl or f"{bibunho}-02" in fl or f"{bibunho}_03" in fl or f"{bibunho}-03" in fl or f"{bibunho}_2" in fl or f"{bibunho}-2" in fl):
                status["part2"] = {"found": True, "exact_match": False, "filename": f}
                matched_files.add(f)
                break

    # 3. 어셈블리 매칭
    for f in files:
        fl = f.lower()
        if fl == expected_assembly.lower():
            status["assembly"] = {"found": True, "exact_match": True, "filename": f}
            matched_files.add(f)
            break
    if not status["assembly"]["found"]:
        for f in files:
            fl = f.lower()
            if fl.endswith(".iam") and (f"{bibunho}" in fl):
                status["assembly"] = {"found": True, "exact_match": False, "filename": f}
                matched_files.add(f)
                break
                
    # 4. STL 매칭
    for f in files:
        fl = f.lower()
        if fl == expected_stl.lower():
            status["stl"] = {"found": True, "exact_match": True, "filename": f}
            matched_files.add(f)
            break
    if not status["stl"]["found"]:
        for f in files:
            fl = f.lower()
            if fl.endswith(".stl") and (f"{bibunho}" in fl):
                status["stl"] = {"found": True, "exact_match": False, "filename": f}
                matched_files.add(f)
                break

    # 5. 슬라이싱 매칭
    slicing_pattern = re.compile(rf"^{bibunho}_04", re.IGNORECASE)
    for f in files:
        fl = f.lower()
        if fl.endswith(('.cfb', '.gcode', '.3gcode', '.zcode', '.hvs')):
            if slicing_pattern.match(f):
                status["slicing"] = {"found": True, "exact_match": True, "filename": f}
                matched_files.add(f)
                break
            elif f"{bibunho}" in fl:
                status["slicing"] = {"found": True, "exact_match": False, "filename": f}
                matched_files.add(f)
                break

    for f in files:
        if f not in matched_files:
            status["extra_files"].append(f)
            
    return status

def run_grading():
    if not os.path.exists(REF_DIR):
        print(f"[오류] 기준 폴더(REF_DIR)가 존재하지 않습니다: {REF_DIR}")
        return
    if not os.path.exists(TARGET_DIR):
        print(f"[오류] 대상 폴더(TARGET_DIR)가 존재하지 않습니다: {TARGET_DIR}")
        return

    # 폴더 목록 추출
    subdirs = sorted([d for d in os.listdir(TARGET_DIR) if os.path.isdir(os.path.join(TARGET_DIR, d)) and d.isdigit()])
    print(f"[정보] 감지된 비번호 폴더 목록: {', '.join(subdirs)}")

    app, launched = get_inventor_app(visible=False)
    
    # 1단계: 기준 폴더(PC-13)에서 정답 체적 데이터 추출
    print("\n--- [1단계] 기준 폴더(PC-13) 정답 체적 데이터 빌드 중 ---")
    ref_database = {}
    for folder in subdirs:
        ref_folder_path = os.path.join(REF_DIR, folder)
        if not os.path.exists(ref_folder_path):
            print(f"[경고] 기준 폴더 내 비번호 {folder}번 폴더가 없습니다.")
            continue
            
        ref_files = os.listdir(ref_folder_path)
        ref_naming = check_naming_convention(folder, ref_files)
        
        ref_iam_file = None
        if ref_naming["assembly"]["found"]:
            ref_iam_file = os.path.join(ref_folder_path, ref_naming["assembly"]["filename"])
        else:
            for f in ref_files:
                if f.lower().endswith(".iam"):
                    ref_iam_file = os.path.join(ref_folder_path, f)
                    break
                    
        if ref_iam_file and os.path.exists(ref_iam_file):
            print(f"  -> 기준 파일 로드: {os.path.basename(ref_iam_file)}")
            info = extract_assembly_info(app, ref_iam_file)
            if info.get("success"):
                ref_database[folder] = info["parts"]
            else:
                print(f"  -> [경고] 기준 파일 분석 실패: {info.get('error')}")
        else:
            print(f"  -> [경고] 기준 어셈블리 파일을 찾을 수 없습니다.")

    # 2단계: 대상 폴더(PC-01) 채점 진행
    print("\n--- [2단계] 대상 폴더(PC-01) 채점 진행 중 ---")
    results = {}
    
    for idx, folder in enumerate(subdirs, 1):
        folder_path = os.path.join(TARGET_DIR, folder)
        files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
        
        print(f"[{idx}/{len(subdirs)}] 비번호 {folder}번 검사 중...")
        
        # 파일명 규칙 검사
        naming_status = check_naming_convention(folder, files)
        
        # 어셈블리 검사
        assembly_info = None
        iam_file = None
        if naming_status["assembly"]["found"]:
            iam_file = os.path.join(folder_path, naming_status["assembly"]["filename"])
        else:
            for f in files:
                if f.lower().endswith(".iam"):
                    iam_file = os.path.join(folder_path, f)
                    break
                    
        if iam_file and os.path.exists(iam_file):
            print(f"  -> 어셈블리 분석 실행: {os.path.basename(iam_file)}")
            assembly_info = extract_assembly_info(app, iam_file)
        else:
            print(f"  -> [경고] 어셈블리(.iam) 파일을 찾을 수 없습니다.")
            assembly_info = {"success": False, "error": "어셈블리(.iam) 파일 없음"}

        results[folder] = {
            "naming": naming_status,
            "assembly": assembly_info
        }
        
    if launched:
        app.Quit()

    # 3단계: 보고서 작성 (PC-01 폴더 아래에 생성)
    report_path = os.path.join(TARGET_DIR, "grade_report_june15.md")
    
    md = []
    md.append(f"# 3D 프린터 운용기능사 실기 시험 자동 채점 보고서 (PC-01)")
    md.append(f"- **채점 일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    md.append(f"- **기준 폴더(PC-13)**: `{REF_DIR}`")
    md.append(f"- **대상 폴더(PC-01)**: `{TARGET_DIR}`")
    md.append(f"- **검사 대상 비번호**: {len(subdirs)}개 ({', '.join(subdirs)}번)\n")
    
    md.append("## 1. 종합 채점 결과 요약")
    md.append("| 비번호 | 적용 비번호 | 파일명 규칙 | 간섭 여부 | 링크 유실 여부 | 체적 오차 검사 | 최종 판정 |")
    md.append("| :---: | :---: | :---: | :---: | :---: | :---: | :---: |")
    
    for folder in subdirs:
        res = results[folder]
        naming = res["naming"]
        assembly = res["assembly"]
        ref_parts = ref_database.get(folder, {})
        
        used_bibunho = naming["used_bibunho"]
        bibunho_type = f"**{used_bibunho}번** (임시)" if naming["is_temporary"] else f"{used_bibunho}번"
        
        # 파일명 규칙 판정
        naming_ok = True
        naming_errors = []
        for key in ["part1", "part2", "assembly", "stl", "slicing"]:
            val = naming[key]
            if not val["found"]:
                naming_ok = False
                naming_errors.append(f"{key} 누락")
            elif not val["exact_match"]:
                naming_ok = False
                naming_errors.append(f"{key} 이름 오류 (`{val['filename']}`)")
        naming_str = "✔️ 정상" if naming_ok else f"❌ 오류 ({', '.join(naming_errors)})"
        
        # 간섭 및 링크 판정
        interference_str = "-"
        link_str = "-"
        vol_str = "-"
        decision = "합격"
        decision_reasons = []
        
        if assembly and assembly.get("success"):
            # 간섭 여부
            if assembly["has_interference"]:
                interference_str = f"❌ 발생 ({assembly['interference_volume']:.1f} mm³)"
                decision = "실격"
                decision_reasons.append("부품 간 간섭")
            else:
                interference_str = "✔️ 없음"
                
            # 링크 유실 여부
            if assembly["unresolved_files"]:
                link_str = f"❌ 유실 ({len(assembly['unresolved_files'])}개)"
                decision = "실격"
                decision_reasons.append("단품 링크 유실")
            else:
                link_str = "✔️ 정상"
                
            # 체적 오차 검사
            if ref_parts:
                vol_ok = True
                vol_errors = []
                for p_key in ["part1", "part2"]:
                    ref_p = ref_parts.get(p_key)
                    std_p = assembly["parts"].get(p_key)
                    
                    if not ref_p:
                        continue
                    if not std_p:
                        vol_ok = False
                        vol_errors.append(f"{p_key} 누락")
                        continue
                        
                    ref_vol = ref_p["volume"]
                    std_vol = std_p["volume"]
                    
                    if ref_vol > 0:
                        diff_pct = abs(std_vol - ref_vol) / ref_vol
                    else:
                        diff_pct = 0.0
                        
                    if diff_pct > 0.05:
                        vol_ok = False
                        vol_errors.append(f"{p_key} 초과 ({diff_pct*100:.1f}%)")
                        
                if vol_ok:
                    vol_str = "✔️ 정상"
                else:
                    vol_str = f"❌ 오차초과 ({', '.join(vol_errors)})"
                    decision = "실격"
                    decision_reasons.append(f"체적 오차 초과 ({', '.join(vol_errors)})")
            else:
                vol_str = "N/A (기준데이터 없음)"
        else:
            err_msg = assembly.get("error") if assembly else "분석 불가"
            interference_str = "N/A"
            link_str = "N/A"
            vol_str = "N/A"
            decision = "실격"
            decision_reasons.append(f"어셈블리 분석 실패 ({err_msg})")
            
        # 파일명 감점 처리
        if not naming_ok:
            has_missing = any(not naming[k]["found"] for k in ["part1", "part2", "assembly", "stl", "slicing"])
            if has_missing:
                decision = "실격"
                decision_reasons.append("필수 파일 누락")
            else:
                if decision != "실격":
                    decision = "감점"
                    decision_reasons.append("파일명 규칙 위반")
                    
        decision_str = f"**{decision}**"
        if decision_reasons:
            decision_str += f" ({', '.join(decision_reasons)})"
            
        md.append(f"| {folder}번 | {bibunho_type} | {naming_str} | {interference_str} | {link_str} | {vol_str} | {decision_str} |")
        
    md.append("\n")
    md.append("## 2. 비번호별 상세 검사 정보")
    
    for folder in subdirs:
        res = results[folder]
        naming = res["naming"]
        assembly = res["assembly"]
        ref_parts = ref_database.get(folder, {})
        bibunho = naming["used_bibunho"]
        
        md.append(f"### 👤 비번호 {folder}번 상세 (실제 비번호: {bibunho}번)")
        
        # 파일 구성 상태
        md.append("#### 📂 파일 구성 및 명명 규칙 상태")
        md.append("| 항목 | 예상 파일명 (기준 비번호: {0}) | 실제 파일명 | 상태 |".format(bibunho))
        md.append("| :--- | :--- | :--- | :--- |")
        
        expected_names = {
            "part1": f"{bibunho}_01.ipt",
            "part2": f"{bibunho}_02.ipt",
            "assembly": f"{bibunho}_03.iam",
            "stl": f"{bibunho}_03.stl",
            "slicing": f"{bibunho}_04[...].(cfb/gcode)"
        }
        
        for key in ["part1", "part2", "assembly", "stl", "slicing"]:
            val = naming[key]
            status_cell = ""
            if not val["found"]:
                status_cell = "❌ **누락**"
            elif val["exact_match"]:
                status_cell = "✔️ **일치**"
            else:
                status_cell = "⚠️ **파일명 오류**"
                
            actual_name = val["filename"] if val["found"] else "-"
            md.append(f"| {key} | `{expected_names[key]}` | `{actual_name}` | {status_cell} |")
            
        if naming["extra_files"]:
            md.append(f"\n- **기타 파일 목록**: {', '.join([f'`{x}`' for x in naming['extra_files']])}")
            
        # 3D 모델링 분석 및 체적 분석
        md.append("\n#### 🛠️ 3D 모델링 분석 상태 (Autodesk Inventor)")
        if assembly and assembly.get("success"):
            md.append(f"- **구성 부품 수**: {assembly['components_count']}개")
            if assembly["has_interference"]:
                md.append(f"- **부품 간 간섭**: ❌ **발생** (간섭 체적: `{assembly['interference_volume']:.2f}` mm³)")
            else:
                md.append("- **부품 간 간섭**: ✔️ **없음 (정상)**")
                
            if assembly["unresolved_files"]:
                md.append(f"- **링크 유실 파일**: ❌ **유실** ({', '.join([f'`{x}`' for x in assembly['unresolved_files']])})")
            else:
                md.append("- **링크 유실 파일**: ✔️ **없음 (정상)**")
                
            # 체적 비교 출력
            md.append("\n**📐 체적(Volume) 비교 분석 (vs PC-13 기준)**")
            md.append("| 부품 | 기준 체적 (PC-13) | 학생 체적 (PC-01) | 오차 비율 | 상태 |")
            md.append("| :--- | :---: | :---: | :---: | :---: |")
            
            for p_key in ["part1", "part2"]:
                ref_p = ref_parts.get(p_key)
                std_p = assembly["parts"].get(p_key)
                
                ref_vol_val = ref_p["volume"] if ref_p else 0.0
                std_vol_val = std_p["volume"] if std_p else 0.0
                
                if ref_vol_val > 0:
                    diff_pct = abs(std_vol_val - ref_vol_val) / ref_vol_val
                else:
                    diff_pct = 0.0
                    
                status_vol = "✔️ **정상**" if diff_pct <= 0.05 else "❌ **초과**"
                if not ref_p:
                    status_vol = "N/A"
                    
                md.append(f"| {p_key} | {ref_vol_val:.2f} mm³ | {std_vol_val:.2f} mm³ | {diff_pct*100:.2f}% | {status_vol} |")
        else:
            err_msg = assembly.get("error") if assembly else "분석 정보 없음"
            md.append(f"- ❌ **어셈블리 분석 실패 사유**: `{err_msg}`")
            
        md.append("\n---\n")
        
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md))
        
    print("\n" + "="*75)
    print(" 6월 15일 TEST PC-01 채점 완료 (기준: PC-13)")
    print("="*75)
    print(f" 생성된 보고서: {report_path}")
    print("="*75)

if __name__ == "__main__":
    run_grading()
