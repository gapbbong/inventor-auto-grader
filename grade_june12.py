# -*- coding: utf-8 -*-
import os
import sys
import zipfile
import hashlib
import re
import tempfile
import shutil
import json
import traceback
import argparse
from datetime import datetime

# 한글 출력 보장을 위한 설정
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

BASE_DIR = r"d:\이갑종\전일제"
JUNE12_DIR = os.path.join(BASE_DIR, "6월12")

def get_inventor_app(visible=False):
    """실행 중인 Inventor를 찾거나 새로 시작합니다."""
    import win32com.client
    try:
        app = win32com.client.gencache.EnsureDispatch("Inventor.Application")
        try:
            # 작동 중인지 확인
            _ = app.Documents.Count
            return app, False
        except Exception:
            pass
    except Exception:
        pass

    # 새로 기동
    try:
        app = win32com.client.gencache.EnsureDispatch("Inventor.Application")
        app.Visible = visible
        return app, True
    except Exception as e:
        print(f"[오류] Inventor를 시작할 수 없습니다. COM API 확인 요망: {e}")
        sys.exit(1)

def guess_part_number(filename):
    """파일명에서 부품 번호(1 또는 2)를 판별합니다."""
    base = os.path.splitext(os.path.basename(filename))[0].lower()
    # "01_01" -> 1, "01_02" -> 2 등 뒷자리 숫자로 구별 시도
    m = re.search(r'[-_](0?1|0?2)$', base)
    if m:
        return f"part{int(m.group(1))}"
    if "01" in base or "part1" in base:
        return "part1"
    if "02" in base or "part2" in base:
        return "part2"
    return None

def extract_assembly_info(app, file_path):
    """어셈블리(.iam) 파일을 분석하여 부품 정보, 간섭 여부, 누락 파일 목록을 추출합니다."""
    doc = None
    try:
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
        # 1. 각 구성 부품 속성 추출
        for i in range(1, occurrences.Count + 1):
            occ = occurrences.Item(i)
            if occ.Suppressed:
                continue
            try:
                mass_properties = occ.MassProperties
                vol_mm3 = mass_properties.Volume * 1000.0 # cm^3 -> mm^3
                occ_doc = occ.Definition.Document
                filepath = occ_doc.FullFileName
                
                part_key = guess_part_number(filepath)
                if not part_key:
                    part_key = f"part{len(parts) + 1}"
                    
                parts[part_key] = {
                    "name": occ.Name,
                    "filename": os.path.basename(filepath),
                    "volume": round(vol_mm3, 2)
                }
            except Exception:
                pass

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
            except Exception:
                pass
                
        # 3. 문서 참조 상태 확인 (누락 파일 검사)
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
        return {"success": False, "error": str(e)}
    finally:
        if doc is not None:
            try:
                doc.Close(True) # 변경 사항 저장 없이 닫기
            except Exception:
                pass

def determine_task_no(iam_path, root_dir):
    """상대 경로를 기반으로 1~14번 과제 번호를 판별합니다."""
    rel_path = os.path.relpath(iam_path, root_dir)
    parts = rel_path.replace('\\', '/').split('/')
    
    # 1. 상위 폴더 이름에서 과제 번호 추출 (가장 깊은 폴더부터 역순으로)
    for part in reversed(parts[:-1]):
        m = re.search(r'\b(0[1-9]|1[0-4]|[1-9])\b', part)
        if m:
            return int(m.group(1))
        m2 = re.match(r'^(\d+)', part)
        if m2:
            num = int(m2.group(1))
            if 1 <= num <= 14:
                return num
                
    # 2. 파일명에서 추출 시도 (예: "03_03.iam" -> 3)
    filename = parts[-1]
    m = re.match(r'^(\d+)', filename)
    if m:
        num = int(m.group(1))
        if 1 <= num <= 14:
            return num
            
    m = re.search(r'\b(0[1-9]|1[0-4]|[1-9])\b', filename)
    if m:
        return int(m.group(1))
        
    return None

def analyze_student_zip(zip_path):
    """개별 학생 ZIP 파일을 압축 해제하여 1~14번 과제 어셈블리 파일을 전부 채점합니다."""
    temp_dir = tempfile.mkdtemp(prefix="grade12_student_")
    results = {}
    
    app, launched = get_inventor_app(visible=False)
    
    # DB 파일 로드
    db = {}
    db_path = os.path.join(BASE_DIR, "reference_db.json")
    if os.path.exists(db_path):
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                db = json.load(f)
        except Exception:
            pass
            
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(temp_dir)
            
        # 모든 .iam 파일 탐색 (OldVersions 제외)
        iam_files = []
        for root, dirs, files in os.walk(temp_dir):
            if 'OldVersions' in root:
                continue
            for f in files:
                if f.lower().endswith('.iam'):
                    iam_files.append(os.path.join(root, f))
                    
        # 각 어셈블리별 과제 판별 및 검사
        for iam_path in iam_files:
            task_no = determine_task_no(iam_path, temp_dir)
            if task_no is None:
                continue
                
            # 한 과제에 여러 iam이 있으면 첫 번째 것만 분석
            if task_no in results:
                continue
                
            # 인벤터 검사 실행
            info = extract_assembly_info(app, iam_path)
            
            if not info["success"]:
                results[task_no] = {
                    "result": "에러",
                    "reason": f"인벤터 로드 실패: {info.get('error')}"
                }
            elif info["unresolved_files"]:
                results[task_no] = {
                    "result": "실격",
                    "reason": f"링크 유실(단품 누락): {', '.join(info['unresolved_files'])}"
                }
            elif info["has_interference"]:
                results[task_no] = {
                    "result": "실격",
                    "reason": f"부품 간 간섭(겹침) 발생 (체적: {info['interference_volume']:.1f} mm³)"
                }
            else:
                # 체적 비교 (허용 오차 5%)
                task_str = str(task_no)
                tolerance_ratio = 0.05
                is_volume_ok = True
                vol_reasons = []
                
                if task_str in db:
                    ref_data = db[task_str]
                    for part_key, ref_part in ref_data.get("parts", {}).items():
                        student_part = info["parts"].get(part_key)
                        if not student_part:
                            is_volume_ok = False
                            vol_reasons.append(f"{part_key} 누락")
                            continue
                            
                        ref_vol = ref_part["volume"]
                        std_vol = student_part["volume"]
                        
                        # 체적이 0인 경우 예외 처리
                        if ref_vol > 0:
                            diff_pct = abs(std_vol - ref_vol) / ref_vol
                        else:
                            diff_pct = 0
                            
                        if diff_pct > tolerance_ratio:
                            is_volume_ok = False
                            vol_reasons.append(f"{part_key} 체적 오차 초과 (기준: {ref_vol:.1f}, 학생: {std_vol:.1f}, 오차: {diff_pct*100:.1f}%)")
                
                if not is_volume_ok:
                    results[task_no] = {
                        "result": "실격",
                        "reason": f"치수 불일치: {'; '.join(vol_reasons)}"
                    }
                else:
                    results[task_no] = {
                        "result": "합격",
                        "reason": "정상 (간섭 없음, 치수 일치)"
                    }
                
    except Exception as e:
        results["error"] = str(e)
    finally:
        try:
            shutil.rmtree(temp_dir)
        except Exception:
            pass
        if launched:
            app.Quit()
            
    return results

def run_main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--student_zip", type=str, help="검사할 개별 학생 ZIP 파일 경로")
    parser.add_argument("--dir", type=str, default="6월12", help="채점할 배치 디렉토리 이름")
    args = parser.parse_args()
    
    if args.student_zip:
        # 개별 학생 검사 (JSON 출력 모드)
        res = analyze_student_zip(args.student_zip)
        print(json.dumps(res, ensure_ascii=False))
        return

    target_dir_name = args.dir
    target_dir_path = os.path.join(BASE_DIR, target_dir_name)

    # 일괄 검사 실행 모드
    if not os.path.exists(target_dir_path):
        print(f"[오류] 폴더가 존재하지 않습니다: {target_dir_path}")
        return
        
    folders = sorted([d for d in os.listdir(target_dir_path) if os.path.isdir(os.path.join(target_dir_path, d))])
    
    zip_tasks = []
    for folder in folders:
        folder_path = os.path.join(target_dir_path, folder)
        zip_files = [f for f in os.listdir(folder_path) if f.endswith('.zip')]
        for zf in zip_files:
            cleaned_name = re.sub(r'^[0-9_\-\s]+', '', zf.replace('.zip', '')).strip()
            if not cleaned_name:
                cleaned_name = zf.replace('.zip', '')
            zip_tasks.append({
                'folder': folder,
                'zip_name': zf,
                'zip_path': os.path.join(folder_path, zf),
                'student_name': cleaned_name
            })
            
    if not zip_tasks:
        print(f"[경고] {target_dir_name} 하위 폴더에서 ZIP 파일을 찾을 수 없습니다.")
        return
        
    print(f"\n[{target_dir_name} 일괄 CAD 검사 시작] 대상 학생 수: {len(zip_tasks)}명")
    
    student_results = {}
    import subprocess
    
    for idx, task in enumerate(zip_tasks, 1):
        label = f"{task['folder']}/{task['zip_name']}"
        print(f"[{idx}/{len(zip_tasks)}] {label} 채점 중 (1~14번 과제 스캔)...")
        
        cmd = [
            sys.executable,
            "-u",
            __file__,
            "--student_zip", task['zip_path']
        ]
        
        try:
            # 학생 한 명당 최대 60초 타임아웃 (최대 14개 과제를 검사하므로 넉넉하게 대기)
            res_sub = subprocess.run(cmd, capture_output=True, text=True, timeout=60, encoding="utf-8")
            if res_sub.returncode == 0:
                try:
                    res_json = json.loads(res_sub.stdout.strip())
                    student_results[task['student_name']] = {
                        'folder': task['folder'],
                        'zip_name': task['zip_name'],
                        'tasks': res_json
                    }
                    # 요약 콘솔 출력
                    passed_cnt = sum(1 for t, r in res_json.items() if r.get('result') == '합격')
                    failed_cnt = sum(1 for t, r in res_json.items() if r.get('result') == '실격')
                    print(f"  -> 완료: 합격 {passed_cnt}개 | 실격 {failed_cnt}개")
                except Exception as e:
                    print(f"  -> 결과 파싱 실패: {e}")
                    if res_sub.stdout:
                        print(f"     [디버그]: {res_sub.stdout}")
                    student_results[task['student_name']] = {
                        'folder': task['folder'], 'zip_name': task['zip_name'], 'error': f'결과 파싱 실패 ({e})'
                    }
            else:
                err_msg = f"프로세스 비정상 종료 (코드: {res_sub.returncode})"
                if res_sub.stderr:
                    err_msg += f" - {res_sub.stderr.strip()}"
                print(f"  -> 오류: {err_msg}")
                student_results[task['student_name']] = {
                    'folder': task['folder'], 'zip_name': task['zip_name'], 'error': err_msg
                }
        except subprocess.TimeoutExpired:
            print("  -> 결과: 실격 (로딩 시간 초과 - 외부 참조 오류 또는 다이얼로그 차단)")
            student_results[task['student_name']] = {
                'folder': task['folder'],
                'zip_name': task['zip_name'],
                'error': '로딩 시간 초과 (차단 또는 무한 대기)'
            }
            # 인벤터 강제 종료
            try:
                subprocess.run(["taskkill", "/f", "/im", "Inventor.exe"], capture_output=True)
            except Exception:
                pass
        except Exception as e:
            print(f"  -> 예외 발생: {e}")
            student_results[task['student_name']] = {
                'folder': task['folder'], 'zip_name': task['zip_name'], 'error': str(e)
            }

    # 종합 리포트 마크다운 파일 작성
    report_filename = f"grade_report_{target_dir_name}.md"
    report_path = os.path.join(target_dir_path, report_filename)
    
    md = []
    md.append(f"# 3D 프린터 운용기능사 {target_dir_name} 과제 자동 채점 보고서")
    md.append(f"- **채점 일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    md.append(f"- **대상 범위**: 1번 ~ 14번 과제 (간섭 및 링크 오류 전수 검사)")
    md.append(f"- **대상 학생**: {len(zip_tasks)}명\n")
    
    md.append("## 1. 학생별 과제 합격/실격 매트릭스")
    md.append("각 셀은 **합격(정상)**, **실격(간섭/누락)** 또는 **미제출(-)** 상태를 나타냅니다.\n")
    
    headers = ["학생명", "반/좌석"] + [f"{t}번" for t in range(1, 15)] + ["제출 개수", "최종 판정"]
    md.append("| " + " | ".join(headers) + " |")
    md.append("| " + " | ".join(["---"] * len(headers)) + " |")
    
    sorted_students = sorted(student_results.keys())
    for name in sorted_students:
        data = student_results[name]
        folder = data['folder']
        
        row = [name, folder]
        tasks = data.get('tasks', {})
        
        if 'error' in data:
            # 에러 발생 학생
            row += ["에러"] * 14 + ["0/14", f"**오류** ({data['error']})"]
        else:
            submitted_cnt = 0
            disqualified_cnt = 0
            disq_reasons = []
            
            for t in range(1, 15):
                t_str = str(t)
                t_val = tasks.get(t_str) or tasks.get(t) # 키 타입 호환성
                
                if t_val:
                    submitted_cnt += 1
                    if t_val['result'] == '합격':
                        row.append("✔️ 합격")
                    else:
                        row.append(f"❌ 실격")
                        disqualified_cnt += 1
                        disq_reasons.append(f"{t}번: {t_val['reason']}")
                else:
                    row.append("-")
            
            row.append(f"{submitted_cnt}/14")
            
            if submitted_cnt == 0:
                row.append("미제출")
            elif disqualified_cnt > 0:
                row.append(f"**실격** ({', '.join(disq_reasons)})")
            else:
                row.append("**합격 (전원 통과)**")
                
        md.append("| " + " | ".join(row) + " |")
        
    md.append("\n")
    md.append("## 2. 세부 실격 사유 및 내용")
    has_any_disq = False
    for name in sorted_students:
        data = student_results[name]
        tasks = data.get('tasks', {})
        if 'error' in data:
            continue
            
        student_disqs = []
        for t, val in sorted(tasks.items(), key=lambda x: int(x[0]) if str(x[0]).isdigit() else 0):
            if val['result'] == '실격':
                student_disqs.append(f"- **{t}번 과제**: {val['reason']}")
                
        if student_disqs:
            has_any_disq = True
            md.append(f"### 👤 {name} ({data['folder']})")
            md.append("\n".join(student_disqs))
            md.append("")
            
    if not has_any_disq:
        md.append("- 실격 또는 감점 사항이 있는 학생이 없습니다. 전원 합격입니다!\n")
        
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md))
        
    print("\n" + "="*75)
    print(" 6월 12일 일괄 채점 완료")
    print("="*75)
    print(f" 생성된 보고서: {report_path}")
    print("="*75)

if __name__ == "__main__":
    run_main()
