# -*- coding: utf-8 -*-
import os
import sys
import zipfile
import hashlib
import re
from collections import defaultdict
from datetime import datetime

# 한글 출력 보장을 위한 설정
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

BASE_DIR = r"d:\이갑종\전일제"

def extract_task_number(filepath):
    """파일 경로에서 과제 번호(1~27)를 추출합니다."""
    # 경로를 분할하여 파일명과 폴더명을 각각 체크
    parts = filepath.replace('\\', '/').split('/')
    for part in parts:
        # 단어 경계를 기준으로 1~27 범위의 숫자를 찾음
        m = re.search(r'\b(0[1-9]|1[0-9]|2[0-7]|[1-9])\b', part)
        if m:
            return int(m.group(1))
    return None

def clean_student_name(filename):
    """학번 등 숫자를 제외하고 깨끗한 학생 이름을 추출합니다."""
    name = filename.replace('.zip', '')
    # 숫자나 공백 제거 (예: "3606박찬희" -> "박찬희", "3419_한채민" -> "한채민")
    name = re.sub(r'^[0-9_\-\s]+', '', name)
    name = re.sub(r'[0-9_\-\s]+$', '', name)
    return name.strip()

def run_analysis():
    print("="*75)
    print(" 전체 날짜별 학생 제출물 분석 시작")
    print("="*75)
    
    # 날짜 폴더 목록 찾기 (예: 5월21, 6월11 등 '월'이 포함되고 폴더인 것)
    date_folders = []
    for item in os.listdir(BASE_DIR):
        item_path = os.path.join(BASE_DIR, item)
        if os.path.isdir(item_path) and '월' in item:
            date_folders.append(item)
            
    # 날짜 정렬 (월과 일 숫자를 추출하여 정렬)
    def date_key(name):
        m = re.match(r'(\d+)월(\d+)', name)
        if m:
            return (int(m.group(1)), int(m.group(2)))
        return (0, 0)
        
    date_folders = sorted(date_folders, key=date_key)
    print(f"[정보] 감지된 날짜 폴더: {', '.join(date_folders)}")
    
    # 데이터 수집용 딕셔너리
    student_submissions = defaultdict(lambda: defaultdict(set))
    student_seats = defaultdict(set)
    hash_map = defaultdict(list)
    
    INVENTOR_EXT = {'.ipt', '.iam', '.idw', '.ipn', '.dwg', '.stp', '.step', '.stl'}
    
    for date_f in date_folders:
        date_path = os.path.join(BASE_DIR, date_f)
        seat_folders = sorted([d for d in os.listdir(date_path) if os.path.isdir(os.path.join(date_path, d))])
        
        for seat in seat_folders:
            seat_path = os.path.join(date_path, seat)
            zip_files = [f for f in os.listdir(seat_path) if f.endswith('.zip')]
            
            for zf_name in zip_files:
                zip_path = os.path.join(seat_path, zf_name)
                student_name = clean_student_name(zf_name)
                if not student_name:
                    student_name = zf_name # 빈 경우 예외 처리
                    
                student_seats[student_name].add(seat)
                student_label = f"{seat}/{zf_name}"
                
                try:
                    with zipfile.ZipFile(zip_path, 'r') as zf:
                        for entry in zf.infolist():
                            if entry.is_dir():
                                continue
                                
                            ext = os.path.splitext(entry.filename)[1].lower()
                            if ext not in INVENTOR_EXT or 'OldVersions' in entry.filename:
                                continue
                                
                            # 과제 번호 추출
                            task_no = extract_task_number(entry.filename)
                            if task_no is not None:
                                student_submissions[student_name][date_f].add(task_no)
                                
                            # MD5 해시 생성 (파일 사이즈가 0보다 큰 경우만)
                            if entry.file_size > 0:
                                try:
                                    data = zf.read(entry.filename)
                                    md5 = hashlib.md5(data).hexdigest()
                                    hash_map[md5].append({
                                        'date': date_f,
                                        'student': student_name,
                                        'label': student_label,
                                        'filename': os.path.basename(entry.filename),
                                        'size': entry.file_size
                                    })
                                except Exception:
                                    pass
                except Exception as e:
                    print(f"[경고] {date_f}/{student_label} 압축파일 읽기 오류: {e}")

    # 마크다운 보고서 작성
    report_path = os.path.join(BASE_DIR, "전체_학생_제출_현황_보고서.md")
    md = []
    md.append(f"# 전체 학생 3D 모델링 제출 현황 및 표절 분석 보고서")
    md.append(f"- **분석 일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    md.append(f"- **대상 날짜**: {len(date_folders)}개 날짜 ({', '.join(date_folders)})")
    md.append(f"- **총 참여 학생**: {len(student_submissions)}명\n")
    
    # 1. 학생별 요약 테이블
    md.append("## 1. 학생별 누적 제출 현황")
    md.append("학생별로 각 날짜에 제출한 과제 번호와 누적 제출 개수를 보여줍니다.\n")
    
    # 테이블 헤더 구성
    header = ["학생명", "반/좌석"] + date_folders + ["누적 제출 개수"]
    md.append("| " + " | ".join(header) + " |")
    md.append("| " + " | ".join(["---"] * len(header)) + " |")
    
    # 학생 이름 가나다 순 정렬
    sorted_students = sorted(student_submissions.keys())
    
    for name in sorted_students:
        seats_str = ", ".join(sorted(student_seats[name]))
        row = [name, seats_str]
        
        total_tasks = set()
        for date_f in date_folders:
            tasks = student_submissions[name].get(date_f, set())
            total_tasks.update(tasks)
            if tasks:
                tasks_str = ", ".join(str(t) for t in sorted(tasks))
                row.append(f"{tasks_str}번")
            else:
                row.append("-")
                
        row.append(f"총 {len(total_tasks)}개 도면")
        md.append("| " + " | ".join(row) + " |")
        
    md.append("\n")
    
    # 2. 날짜별 제출률 통계
    md.append("## 2. 날짜별 제출 현황 및 통계")
    for date_f in date_folders:
        md.append(f"### 📅 {date_f} 제출 현황")
        submitted_students = []
        for name in sorted_students:
            if date_f in student_submissions[name]:
                tasks = sorted(student_submissions[name][date_f])
                submitted_students.append(f"**{name}** ({', '.join(str(t) for t in tasks)}번)")
                
        md.append(f"- **제출 학생 수**: {len(submitted_students)}명 / {len(sorted_students)}명")
        md.append(f"- **제출자 명단**: {', '.join(submitted_students) if submitted_students else '제출 없음'}\n")
        
    # 3. 표절 의심 검사 (서로 다른 학생 간 동일 MD5 해시 소유 여부)
    md.append("## 3. 학생 간 표절(복사본 제출) 의심 분석")
    md.append("서로 다른 학생이 내용이 완전히 일치하는 파일(동일 MD5)을 제출한 경우입니다. (폴더 구조나 파일 이름이 달라도 검출됨)\n")
    
    plagiarism_cases = []
    for md5, instances in hash_map.items():
        # 고유한 학생 이름 집합 구하기
        unique_students = set(inst['student'] for inst in instances)
        if len(unique_students) > 1:
            plagiarism_cases.append((md5, instances))
            
    if plagiarism_cases:
        md.append(f"> [!WARNING]\n> **총 {len(plagiarism_cases)}건의 파일 공유(복제본) 의심 사례가 발견되었습니다.**\n")
        md.append("| 번호 | 의심 파일명 | 크기 | 날짜 / 학생 및 파일 경로 |")
        md.append("| --- | --- | --- | --- |")
        for idx, (md5, instances) in enumerate(plagiarism_cases, 1):
            file_name = instances[0]['filename']
            file_size = f"{instances[0]['size'] / 1024:.1f} KB"
            
            details = []
            for inst in instances:
                details.append(f"[{inst['date']}] {inst['student']} ({inst['label']} -> {inst['filename']})")
                
            md.append(f"| {idx} | `{file_name}` | {file_size} | {', '.join(details)} |")
    else:
        md.append("> [!NOTE]\n> **동일한 MD5 파일을 중복 제출한 표절 의심 사례가 발견되지 않았습니다. 학생들의 제출물이 모두 고유합니다.**\n")
        
    # 파일 쓰기
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md))
        
    print("="*75)
    print(f" 분석 완료! 보고서가 저장되었습니다: {report_path}")
    print("="*75)
    print(f"  - 대상 날짜 수: {len(date_folders)}개")
    print(f"  - 총 학생 수: {len(student_submissions)}명")
    print(f"  - 표절 의심 파일 수: {len(plagiarism_cases)}개")
    print("="*75)

if __name__ == '__main__':
    run_analysis()
