# -*- coding: utf-8 -*-
"""
Autodesk Inventor 2022 COM API 기반 자동 채점 스크립트 (grade_inventor.py)
- 3D 프린터 운용기능사 공개도면 01-27번 채점용
- 정답 등록 모드: --register <번호> --file <정답_어셈블리_경로>
- 학생 채점 모드: --grade <학생_ZIP_경로_또는_폴더> --task <번호> (기본값: 폴더명/파일명에서 자동 추출)
"""

import os
import sys
import json
import re
import zipfile
import shutil
import tempfile
import argparse
import traceback
from datetime import datetime

# UTF-8 출력 강제 설정
sys.stdout.reconfigure(encoding='utf-8')

# COM API 모듈 로드
try:
    import win32com.client
    from win32com.client import Dispatch
except ImportError:
    print("[오류] pywin32 패키지가 설치되어 있지 않습니다. 'pip install pywin32'를 실행해 주세요.")
    sys.exit(1)

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reference_db.json")

def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_db(db):
    try:
        with open(DB_FILE, "w", encoding="utf-8") as f:
            json.dump(db, f, indent=4, ensure_ascii=False)
        print(f"[정보] 정답 데이터베이스가 업데이트되었습니다: {DB_FILE}")
    except Exception as e:
        print(f"[오류] 데이터베이스 저장 실패: {e}")

def get_inventor_app(visible=False):
    """실행 중인 인벤터를 가져오거나 새로 실행합니다."""
    try:
        import win32com.client.gencache
        app = win32com.client.gencache.EnsureDispatch("Inventor.Application")
        try:
            # 이미 실행 중인 경우
            _ = app.Documents.Count
            print("[정보] 실행 중인 Autodesk Inventor 연결 성공.")
            return app, False
        except Exception:
            # 새로 실행되는 경우
            app.Visible = visible
            print("[정보] Autodesk Inventor 새로 시작 성공.")
            return app, True
    except Exception as e:
        print(f"[오류] Autodesk Inventor를 제어할 수 없습니다: {e}")
        sys.exit(1)

def guess_part_number(filepath):
    """파일명에서 Part 1 또는 Part 2인지 판별합니다."""
    name = os.path.basename(filepath).lower()
    # 01_Part1, 01_1, 22_01 등
    if '01' in name or '_1' in name or 'part1' in name:
        return 1
    if '02' in name or '_2' in name or 'part2' in name:
        return 2
    return None

def extract_assembly_info(app, file_path):
    """어셈블리(.iam) 파일을 열어 각 부품의 체적(mm³)과 간섭을 분석합니다."""
    doc = None
    try:
        # 파일이 존재하지 않는 경우 예외 처리
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {file_path}")

        # 인벤터 경고창 비활성화 시도 (COM 오류 방지)
        try:
            app.SilentFiles = True
        except Exception:
            pass
        
        # 문서 열기 및 AssemblyDocument로 캐스팅
        raw_doc = app.Documents.Open(file_path, OpenVisible=True)
        import win32com.client
        try:
            doc = win32com.client.CastTo(raw_doc, "AssemblyDocument")
        except Exception:
            doc = raw_doc
            
        comp_def = doc.ComponentDefinition
        
        occurrences = comp_def.Occurrences
        parts = {}
        
        # 1. 각 Occurrences(부품) 분석
        for i in range(1, occurrences.Count + 1):
            occ = occurrences.Item(i)
            if occ.Suppressed:
                continue
                
            try:
                # 물리적 성질 추출 (Inventor 내부 단위는 cm이며, 체적은 cm^3 단위임)
                mass_properties = occ.MassProperties
                vol_cm3 = mass_properties.Volume
                area_cm2 = mass_properties.Area
                mass_kg = mass_properties.Mass
                
                # mm 단위로 변환
                vol_mm3 = vol_cm3 * 1000.0
                area_mm2 = area_cm2 * 100.0
                mass_g = mass_kg * 1000.0
                
                occ_doc = occ.Definition.Document
                filepath = occ_doc.FullFileName
                
                part_num = guess_part_number(filepath)
                if part_num is None:
                    # 파일명으로 판별 불가시 순서대로 또는 이름으로 저장
                    part_num = len(parts) + 1
                
                parts[f"part{part_num}"] = {
                    "name": occ.Name,
                    "filename": os.path.basename(filepath),
                    "volume": round(vol_mm3, 2),
                    "area": round(area_mm2, 2),
                    "mass": round(mass_g, 2)
                }
            except Exception as e:
                print(f"[경고] 구성 부품 '{occ.Name}' 속성 추출 실패: {e}")

        # 2. 부품 간 간섭 검사 (Interference Analysis)
        interference_vol_mm3 = 0.0
        has_interference = False
        
        if occurrences.Count >= 2:
            oGroup1 = app.TransientObjects.CreateObjectCollection()
            oGroup2 = app.TransientObjects.CreateObjectCollection()
            
            # 2개 부품이므로 1번과 2번의 간섭 검사 진행
            occ1 = occurrences.Item(1)
            occ2 = occurrences.Item(2)
            
            oGroup1.Add(occ1)
            oGroup2.Add(occ2)
            
            try:
                results = comp_def.AnalyzeInterference(oGroup1, oGroup2)
                if results.Count > 0:
                    has_interference = True
                    for r_idx in range(1, results.Count + 1):
                        # cm^3 -> mm^3 변환
                        interference_vol_mm3 += results.Item(r_idx).Volume * 1000.0
            except Exception as e:
                print(f"[경고] 간섭 분석 중 오류 발생: {e}")
                
        # 문서 참조 상태 확인 (누락 파일 검사)
        unresolved_files = []
        for ref_desc in doc.ReferencedFileDescriptors:
            if ref_desc.ReferenceStatus == 2:  # kMissingReference (참조 파일 유실)
                unresolved_files.append(ref_desc.DisplayName)

        return {
            "success": True,
            "parts": parts,
            "has_interference": has_interference,
            "interference_volume": round(interference_vol_mm3, 2),
            "unresolved_files": unresolved_files
        }
        
    except Exception as e:
        traceback.print_exc()
        return {"success": False, "error": str(e)}
    finally:
        if doc is not None:
            try:
                doc.Close(True)  # 변경 사항 저장 없이 닫기 (SkipSave=True)
            except Exception:
                pass

def handle_register(args):
    """선택한 과제 번호에 맞춰 정답 데이터를 등록합니다."""
    print(f"\n[정답 등록 시작] 과제 번호: {args.register} | 파일: {args.file}")
    
    app, launched = get_inventor_app(visible=True)
    try:
        abs_path = os.path.abspath(args.file)
        res = extract_assembly_info(app, abs_path)
        
        if not res["success"]:
            print(f"[오류] 정답 추출에 실패했습니다: {res.get('error')}")
            return
            
        db = load_db()
        db[str(args.register)] = {
            "parts": res["parts"],
            "registered_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        save_db(db)
        
        print("\n=== 등록된 정답 정보 ===")
        print(json.dumps(db[str(args.register)], indent=2, ensure_ascii=False))
        
    finally:
        if launched:
            app.Quit()

def inspect_student_folder(app, folder_path, task_no, ref_data):
    """개별 학생 폴더 안의 어셈블리를 찾아 채점합니다."""
    # iam 파일 검색
    # 구조: task_no_3/ 폴더 아래의 .iam 파일
    iam_files = []
    task_dir_name = f"{task_no}_3"
    
    for root, dirs, files in os.walk(folder_path):
        if task_dir_name in root:
            for f in files:
                if f.lower().endswith(".iam"):
                    iam_files.append(os.path.join(root, f))
                    
    # 지정 경로 하위에 해당 폴더가 없는 경우 전체 폴더에서 해당 과제 번호 iam 검색
    if not iam_files:
        for root, dirs, files in os.walk(folder_path):
            for f in files:
                if f.lower().endswith(".iam") and f"_{task_no}" in f:
                    iam_files.append(os.path.join(root, f))
                elif f.lower().endswith(".iam") and str(task_no) in root:
                    iam_files.append(os.path.join(root, f))

    if not iam_files:
        return {"result": "실격", "reason": "조립 파일(.iam) 누락"}
        
    target_iam = iam_files[0]
    print(f"  -> 분석 대상 어셈블리: {os.path.basename(target_iam)}")
    
    res = extract_assembly_info(app, target_iam)
    if not res["success"]:
        return {"result": "에러", "reason": f"인벤터 로드 실패: {res.get('error')}"}
        
    if res["unresolved_files"]:
        return {"result": "실격", "reason": f"링크 유실(단품 누락): {', '.join(res['unresolved_files'])}"}

    if res["has_interference"]:
        return {"result": "실격", "reason": f"부품 간 간섭(겹침) 발생 (체적: {res['interference_volume']:.1f} mm³)"}

    # 체적 비교 (가변치수 A, B 오차 허용 비율 설정: 기본 +-5%)
    tolerance_ratio = 0.05
    parts_status = {}
    is_all_passed = True
    reasons = []

    for part_key, ref_part in ref_data["parts"].items():
        student_part = res["parts"].get(part_key)
        if not student_part:
            is_all_passed = False
            reasons.append(f"{part_key} 누락")
            continue
            
        ref_vol = ref_part["volume"]
        std_vol = student_part["volume"]
        
        diff_pct = abs(std_vol - ref_vol) / ref_vol
        
        if diff_pct <= tolerance_ratio:
            parts_status[part_key] = "PASS"
        else:
            is_all_passed = False
            parts_status[part_key] = "FAIL"
            reasons.append(f"{part_key} 체적 오차 초과 (정답: {ref_vol:.1f}, 제출: {std_vol:.1f}, 오차: {diff_pct*100:.1f}%)")

    if is_all_passed:
        return {
            "result": "합격",
            "details": res,
            "parts_status": parts_status
        }
    else:
        return {
            "result": "감점/불합격",
            "reason": "; ".join(reasons),
            "details": res,
            "parts_status": parts_status
        }

def handle_grade(args):
    """지정된 학생 제출물(ZIP 또는 폴더)을 대상으로 채점을 수행합니다."""
    is_json_mode = getattr(args, 'json', False)
    import io
    original_stdout = sys.stdout
    if is_json_mode:
        sys.stdout = io.StringIO()
        
    db = load_db()
    
    target_path = os.path.abspath(args.grade)
    if not os.path.exists(target_path):
        if is_json_mode:
            sys.stdout = original_stdout
            print(json.dumps({"result": "에러", "reason": f"대상 경로가 존재하지 않습니다: {target_path}"}, ensure_ascii=False))
        else:
            print(f"[오류] 대상 경로가 존재하지 않습니다: {target_path}")
        return

    # 과제 번호 추출
    task_no = args.task
    if not task_no:
        m = re.search(r'(0[1-9]|1[0-9]|2[0-7]|\b[1-9]\b)', os.path.basename(target_path))
        if m:
            task_no = int(m.group(1))
        else:
            if is_json_mode:
                sys.stdout = original_stdout
                print(json.dumps({"result": "에러", "reason": "과제 번호를 판별할 수 없습니다. --task 옵션을 명시해 주세요."}, ensure_ascii=False))
            else:
                print("[오류] 과제 번호를 판별할 수 없습니다. --task <번호> 옵션을 명시해 주세요.")
            return
            
    task_str = str(task_no)
    if task_str not in db:
        if is_json_mode:
            sys.stdout = original_stdout
            print(json.dumps({"result": "에러", "reason": f"과제 {task_no}번에 대한 정답 데이터가 등록되어 있지 않습니다."}, ensure_ascii=False))
        else:
            print(f"[오류] 과제 {task_no}번에 대한 정답 데이터가 등록되어 있지 않습니다.")
            print("먼저 '--register <번호> --file <정답_경로.iam>' 명령으로 정답을 등록해 주세요.")
        return
        
    ref_data = db[task_str]
    
    # 임시 작업용 디렉토리 생성
    temp_dir = tempfile.mkdtemp(prefix="grade_temp_")
    
    app, launched = get_inventor_app(visible=args.visible)
    
    result = {"result": "에러", "reason": "알 수 없는 오류"}
    try:
        if zipfile.is_zipfile(target_path):
            with zipfile.ZipFile(target_path, 'r') as zf:
                zf.extractall(temp_dir)
            grade_folder = temp_dir
        else:
            grade_folder = target_path
            
        result = inspect_student_folder(app, grade_folder, task_no, ref_data)
        
        if not is_json_mode:
            print("\n" + "="*50)
            print(f" 채점 결과 - 과제 {task_no}번 ({os.path.basename(target_path)})")
            print("="*50)
            print(f" 판정: {result['result']}")
            if "reason" in result:
                print(f" 사유: {result['reason']}")
                
            if "parts_status" in result:
                print("\n [부품별 통과 상태]")
                for pk, status in result["parts_status"].items():
                    ref_vol = ref_data["parts"][pk]["volume"]
                    std_vol = result["details"]["parts"].get(pk, {}).get("volume", 0.0)
                    print(f"  - {pk}: {status} (정답: {ref_vol:.1f} mm³ | 제출: {std_vol:.1f} mm³)")
                    
            if "details" in result and result["details"].get("success"):
                print(f"\n 간섭 검사: {'[실격] 간섭 발생' if result['details']['has_interference'] else '통과 (간섭 없음)'}")
                if result['details']['has_interference']:
                    print(f"  - 간섭 체적: {result['details']['interference_volume']:.1f} mm³")
            print("="*50)
            
    except Exception as e:
        result = {"result": "에러", "reason": f"처리 중 예외 발생: {e}"}
        if not is_json_mode:
            traceback.print_exc()
    finally:
        # 임시 디렉토리 제거
        try:
            shutil.rmtree(temp_dir)
        except Exception:
            pass
            
        if launched:
            app.Quit()
            
        # JSON 모드이면 원래 stdout을 복구하고 JSON만 콘솔에 출력
        if is_json_mode:
            sys.stdout = original_stdout
            print(json.dumps(result, ensure_ascii=False))

def handle_batch(args):
    """기준 디렉토리 내의 모든 학생 ZIP 파일에 대해 채점을 일괄 수행합니다."""
    db = load_db()
    
    batch_dir = os.path.abspath(args.batch)
    if not os.path.exists(batch_dir):
        print(f"[오류] 대상 배치 디렉토리가 존재하지 않습니다: {batch_dir}")
        return

    task_no = args.task
    if not task_no:
        # 경로명에서 과제번호 추출 시도
        m = re.search(r'(0[1-9]|1[0-9]|2[0-7]|\b[1-9]\b)', os.path.basename(batch_dir))
        if m:
            task_no = int(m.group(1))
            print(f"[정보] 경로명에서 과제 번호 {task_no}번을 자동으로 감지했습니다.")
        else:
            print("[오류] 배치 채점을 위해서는 과제 번호를 명시해야 합니다. --task <번호> 옵션을 명시해 주세요.")
            return

    task_str = str(task_no)
    if task_str not in db:
        print(f"[오류] 과제 {task_no}번에 대한 정답 데이터가 등록되어 있지 않습니다.")
        print("먼저 '--register <번호> --file <정답_경로.iam>' 명령으로 정답을 등록해 주세요.")
        return
        
    ref_data = db[task_str]
    
    # 학생 폴더 스캔
    folders = sorted([f for f in os.listdir(batch_dir) if os.path.isdir(os.path.join(batch_dir, f))])
    
    zip_tasks = []
    for folder in folders:
        folder_path = os.path.join(batch_dir, folder)
        zip_files = [f for f in os.listdir(folder_path) if f.endswith('.zip')]
        for zf in zip_files:
            zip_tasks.append({
                'folder': folder,
                'zip_name': zf,
                'zip_path': os.path.join(folder_path, zf)
            })
            
    if not zip_tasks:
        print(f"[경고] {batch_dir} 하위 폴더에서 ZIP 파일을 찾을 수 없습니다.")
        return
        
    print(f"\n[일괄 채점 시작] 과제 번호: {task_no}번 | 대상 ZIP 파일: {len(zip_tasks)}개")
    
    results = []
    import subprocess
    
    for idx, task in enumerate(zip_tasks, 1):
        student_label = f"{task['folder']}/{task['zip_name']}"
        student_name = task['zip_name'].replace('.zip', '')
        print(f"\n[{idx}/{len(zip_tasks)}] {student_label} 채점 중...")
        
        # 각 학생을 독립된 프로세스로 실행 (타임아웃으로 블로킹/다운 방지)
        cmd = [
            sys.executable,
            "-u",
            __file__,
            "--grade", task['zip_path'],
            "--task", str(task_no),
            "--json"
        ]
        if args.visible:
            cmd.append("--visible")
            
        try:
            # 학생 한 명당 최대 30초의 제한 시간 적용
            res_sub = subprocess.run(cmd, capture_output=True, text=True, timeout=30, encoding="utf-8")
            if res_sub.returncode == 0:
                try:
                    res_json = json.loads(res_sub.stdout.strip())
                    res_json['student_name'] = student_name
                    res_json['folder'] = task['folder']
                    res_json['zip_name'] = task['zip_name']
                    results.append(res_json)
                    print(f"  -> 결과: {res_json['result']}")
                    if 'reason' in res_json:
                        print(f"  -> 사유: {res_json['reason']}")
                except Exception as e:
                    err_msg = f"결과 파싱 실패 (오류: {e})"
                    print(f"  -> 결과: 에러 ({err_msg})")
                    if res_sub.stdout:
                        print(f"  [디버그 출력]: {res_sub.stdout}")
                    results.append({
                        'student_name': student_name,
                        'folder': task['folder'],
                        'zip_name': task['zip_name'],
                        'result': '에러',
                        'reason': err_msg
                    })
            else:
                err_msg = f"프로세스 비정상 종료 (종료 코드: {res_sub.returncode})"
                if res_sub.stderr:
                    err_msg += f" - {res_sub.stderr.strip()}"
                print(f"  -> 결과: 에러 ({err_msg})")
                results.append({
                    'student_name': student_name,
                    'folder': task['folder'],
                    'zip_name': task['zip_name'],
                    'result': '에러',
                    'reason': err_msg
                })
        except subprocess.TimeoutExpired:
            print("  -> 결과: 실격 (로딩 시간 초과 - 외부 참조 오류 또는 다이얼로그 차단)")
            results.append({
                'student_name': student_name,
                'folder': task['folder'],
                'zip_name': task['zip_name'],
                'result': '실격',
                'reason': '로딩 시간 초과 (외부 참조 파일 유실 또는 CATIA 등 파일 형식 불일치)'
            })
            # 블로킹된 인벤터 강제 종료하여 다음 검사가 진행되도록 처리
            print("  [경고] 타임아웃 발생으로 차단된 인벤터 프로세스를 강제 청소합니다...")
            try:
                subprocess.run(["taskkill", "/f", "/im", "Inventor.exe"], capture_output=True)
            except Exception:
                pass
        except Exception as e:
            err_msg = f"실행 중 예외 발생: {e}"
            print(f"  -> 결과: 에러 ({err_msg})")
            results.append({
                'student_name': student_name,
                'folder': task['folder'],
                'zip_name': task['zip_name'],
                'result': '에러',
                'reason': err_msg
            })
            
    # 결과 요약 및 보고서 작성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_filename = f"grade_report_task_{task_no}_{timestamp}.md"
    report_path = os.path.join(batch_dir, report_filename)
    
    total_count = len(results)
    pass_count = sum(1 for r in results if r['result'] == '합격')
    fail_count = sum(1 for r in results if r['result'] in ('감점/불합격', '실격'))
    err_count = sum(1 for r in results if r['result'] == '에러')
    
    # 마크다운 작성
    md_content = []
    md_content.append(f"# 3D 프린터 운용기능사 자동 채점 보고서 (인벤터 API 기반)")
    md_content.append(f"- **채점 과제**: {task_no}번 도면")
    md_content.append(f"- **채점 일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    md_content.append(f"- **총 검사 대상**: {total_count}명 (합격: {pass_count}명, 불합격/실격: {fail_count}명, 오류: {err_count}명)")
    if total_count > 0:
        md_content.append(f"- **합격률**: {pass_count/total_count*100:.1f}%\n")
    else:
        md_content.append(f"- **합격률**: 0.0%\n")
    
    md_content.append("## 채점 결과 리스트")
    md_content.append("| 번호 | 학생명 | 판정 | 사유 및 상세 |")
    md_content.append("| --- | --- | --- | --- |")
    
    for r in results:
        reason = r.get('reason', '-')
        if r['result'] == '합격':
            # 부품 체적 오차 세부 정보 추가
            parts_info = []
            for pk, ref_part in ref_data['parts'].items():
                std_vol = r.get('details', {}).get('parts', {}).get(pk, {}).get('volume', 0.0)
                ref_vol = ref_part['volume']
                diff_pct = abs(std_vol - ref_vol) / ref_vol * 100 if ref_vol > 0 else 0
                parts_info.append(f"{pk}: 오차 {diff_pct:.1f}%")
            reason = "정상 (" + ", ".join(parts_info) + ")"
            
        md_content.append(f"| {r['folder']} | {r['student_name']} | **{r['result']}** | {reason} |")
        
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_content))
        
    print("\n" + "="*70)
    print(f" 일괄 채점 완료")
    print("="*70)
    print(f" 대상 학생 수: {total_count}명")
    print(f" 합격자 수: {pass_count}명 | 불합격/실격: {fail_count}명 | 에러: {err_count}명")
    if total_count > 0:
        print(f" 합격률: {pass_count/total_count*100:.1f}%")
    print(f" 상세 마크다운 보고서가 생성되었습니다: {report_path}")
    print("="*70)

def main():
    parser = argparse.ArgumentParser(description="Autodesk Inventor 3D 프린터 운용기능사 자동 채점 프로그램")
    group = parser.add_mutually_exclusive_group(required=True)
    
    group.add_argument("--register", type=int, help="정답 모델을 등록할 과제 번호 (1-27)")
    group.add_argument("--grade", type=str, help="채점할 학생의 ZIP 파일 또는 폴더 경로")
    group.add_argument("--batch", type=str, help="여러 학생 폴더(ZIP 포함)가 들어있는 기준 디렉토리 경로")
    
    parser.add_argument("--file", type=str, help="정답 모델의 어셈블리(.iam) 파일 경로 (정답 등록 시 필수)")
    parser.add_argument("--task", type=int, help="채점할 과제 번호 (생략 시 경로/파일명에서 자동 추출)")
    parser.add_argument("--visible", action="store_true", help="인벤터 창을 화면에 띄웁니다.")
    parser.add_argument("--json", action="store_true", help="결과를 JSON 형식으로 콘솔에 출력합니다 (내부 호출용).")
    
    args = parser.parse_args()
    
    if args.register:
        if not args.file:
            parser.error("--register 사용 시 --file <정답_어셈블리_경로.iam> 지정이 필수적입니다.")
        handle_register(args)
    elif args.grade:
        handle_grade(args)
    elif args.batch:
        handle_batch(args)

if __name__ == "__main__":
    main()
