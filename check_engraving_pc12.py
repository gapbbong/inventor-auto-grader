# -*- coding: utf-8 -*-
"""
Autodesk Inventor COM API 기반 비번호 각인 누락 검사 스크립트 (check_engraving_pc12.py)
- PC-12 폴더 내의 모든 과제(15번~27번)를 대상으로 비번호 각인(스케치 텍스트 상자) 누락 여부를 검사합니다.
"""

import os
import sys
import re
import win32com.client
from datetime import datetime

# 한글 출력 보장을 위한 설정
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

TARGET_DIR = r"D:\이갑종\26 3학년 전일제\6월15 TEST\PC-12_10.128.56.102"

def get_inventor_app(visible=False):
    """실행 중인 Inventor를 찾거나 새로 시작합니다."""
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

def check_engraving_in_assembly(app, iam_path):
    """어셈블리 내 단품들의 스케치 텍스트 상자 유무를 검사합니다."""
    doc = None
    try:
        raw_doc = app.Documents.Open(iam_path, OpenVisible=False)
        try:
            doc = win32com.client.CastTo(raw_doc, "AssemblyDocument")
        except Exception:
            doc = raw_doc
            
        comp_def = doc.ComponentDefinition
        occurrences = comp_def.Occurrences
        
        engraving_status = {}
        for i in range(1, occurrences.Count + 1):
            occ = occurrences.Item(i)
            if occ.Suppressed:
                continue
                
            raw_occ_doc = occ.Definition.Document
            filename = os.path.basename(raw_occ_doc.FullFileName)
            
            # 파트 문서로 명시적 캐스팅
            try:
                occ_doc = win32com.client.CastTo(raw_occ_doc, "PartDocument")
            except Exception:
                occ_doc = raw_occ_doc
            
            # 파트 문서(12290)인 경우에만 검사
            if occ_doc.DocumentType == 12290:
                part_def = occ_doc.ComponentDefinition
                sketches = part_def.Sketches
                
                texts_found = []
                for s_idx in range(1, sketches.Count + 1):
                    sketch = sketches.Item(s_idx)
                    tb_count = sketch.TextBoxes.Count
                    if tb_count > 0:
                        for t_idx in range(1, tb_count + 1):
                            texts_found.append(sketch.TextBoxes.Item(t_idx).Text)
                            
                engraving_status[filename] = {
                    "has_engraving": (len(texts_found) > 0),
                    "texts": texts_found
                }
        return {"success": True, "parts": engraving_status}
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"success": False, "error": str(e)}
    finally:
        if doc is not None:
            try:
                doc.Close(True)
            except Exception:
                pass

def run_analysis():
    if not os.path.exists(TARGET_DIR):
        print(f"[오류] 대상 폴더가 존재하지 않습니다: {TARGET_DIR}")
        return

    # 비번호 폴더 목록 추출
    subdirs = sorted([d for d in os.listdir(TARGET_DIR) if os.path.isdir(os.path.join(TARGET_DIR, d)) and d.isdigit()], key=lambda x: int(x))
    print(f"[정보] 감지된 비번호 폴더 목록: {', '.join(subdirs)}")

    app, launched = get_inventor_app(visible=False)
    
    results = {}
    
    try:
        app.SilentFiles = True
    except:
        pass

    try:
        for folder in subdirs:
            folder_path = os.path.join(TARGET_DIR, folder)
            files = os.listdir(folder_path)
            
            # 어셈블리 파일 탐색
            iam_file = None
            for f in files:
                if f.lower().endswith(".iam"):
                    iam_file = os.path.join(folder_path, f)
                    break
                    
            if iam_file and os.path.exists(iam_file):
                print(f"[{folder}번 과제] 각인 검사 중...")
                info = check_engraving_in_assembly(app, iam_file)
                results[folder] = info
            else:
                results[folder] = {"success": False, "error": "어셈블리(.iam) 파일 없음"}
                print(f"[{folder}번 과제] [경고] 어셈블리 파일 없음")
                
    finally:
        if launched:
            app.Quit()

    # 보고서 작성
    report_path = os.path.join(TARGET_DIR, "engraving_report_pc12.md")
    
    md = []
    md.append(f"# PC-12 수험생 비번호 각인 누락 정밀 검사 보고서")
    md.append(f"- **검사 일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    md.append(f"- **대상 폴더**: `{TARGET_DIR}`\n")
    
    md.append("## 1. 각인 누락 여부 종합 테이블")
    md.append("| 과제 번호 | 각인 검출 결과 | 부품별 상세 내용 | 판정 |")
    md.append("| :---: | :--- | :--- | :---: |")
    
    missing_count = 0
    for folder in subdirs:
        res = results[folder]
        
        status_str = ""
        detail_parts = []
        decision = "정상"
        
        if res.get("success"):
            parts = res["parts"]
            total_texts = []
            
            for fname, p_info in parts.items():
                if p_info["has_engraving"]:
                    text_list = ", ".join([f"'{t}'" for t in p_info["texts"]])
                    detail_parts.append(f"`{fname}`: 각인 있음 ({text_list})")
                    total_texts.extend(p_info["texts"])
                else:
                    detail_parts.append(f"`{fname}`: 각인 없음")
                    
            if len(total_texts) > 0:
                status_str = f"✔️ 각인 검출 ({', '.join(set(total_texts))})"
            else:
                status_str = "❌ **각인 누락**"
                decision = "**각인 누락**"
                missing_count += 1
        else:
            status_str = "❌ 분석 불가"
            decision = "**분석 실패**"
            detail_parts.append(f"이유: {res.get('error', '어셈블리 없음')}")
            missing_count += 1
            
        md.append(f"| {folder}번 | {status_str} | {'; '.join(detail_parts)} | {decision} |")
        
    md.append(f"\n- **총 검사 대상**: {len(subdirs)}개 과제")
    md.append(f"- **각인 누락(또는 분석 실패) 과제**: {missing_count}개")
    
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md))
        
    print("\n" + "="*75)
    print(" PC-12 각인 분석 완료!")
    print("="*75)
    print(f" 생성된 보고서: {report_path}")
    print("="*75)

if __name__ == "__main__":
    run_analysis()
