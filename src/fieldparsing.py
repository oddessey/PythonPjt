import pandas as pd
import re

def process_data_standardization(input_path, output_path):
    xls = pd.ExcelFile(input_path)
    
    # 1. 헤더 행(Header Row) 자동 탐색 로더
    def load_sheet_smart(sheet_name):
        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        header_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = ' '.join(row.dropna().astype(str))
            if any(k in row_str for k in ['단어', '도메인', '컬럼', '객체', '순번', '데이터베이스']):
                header_idx = idx
                break
        df = pd.read_excel(xls, sheet_name=sheet_name, header=header_idx)
        df.columns = df.columns.astype(str).str.strip().str.replace(r'\s+', '', regex=True)
        return df

    # 시트 로딩
    df_table = load_sheet_smart('전체테이블목록')
    df_dict = load_sheet_smart('단어사전')
    df_domain = load_sheet_smart('도메인정의서')

    # 2. 단어사전 매핑 데이터 구성 (NaN 값 문자열 안전 처리)
    dict_sorted_kr = df_dict.dropna(subset=['단어']).sort_values(
        by='단어', key=lambda x: x.astype(str).str.len(), ascending=False
    )
    
    kr_word_to_abbr = {}
    for _, row in dict_sorted_kr.iterrows():
        word = str(row['단어']).strip()
        # 약어가 NaN인 경우 원본 단어로 대체
        abbr = str(row['영문약어']).strip() if pd.notna(row['영문약어']) and str(row['영문약어']).strip().lower() != 'nan' else word
        kr_word_to_abbr[word] = abbr

    # 영문 물리명 -> 표준 약어 매핑
    eng_abbr_map = {}
    for _, row in df_dict.iterrows():
        abbr = str(row['영문약어']).strip().upper() if pd.notna(row['영문약어']) and str(row['영문약어']).strip().lower() != 'nan' else ""
        cur_abbr = str(row['현행영문약어']).strip().upper() if pd.notna(row['현행영문약어']) and str(row['현행영문약어']).strip().lower() != 'nan' else ""
        full_eng = str(row['전체영문명']).strip().upper() if pd.notna(row['전체영문명']) and str(row['전체영문명']).strip().lower() != 'nan' else ""
        
        if abbr:
            eng_abbr_map[abbr] = abbr
        if cur_abbr and abbr:
            eng_abbr_map[cur_abbr] = abbr
        if full_eng and abbr:
            eng_abbr_map[full_eng] = abbr

    # 3. 도메인 매핑 데이터 구성
    domain_map_kr = {}
    domain_map_abbr = {}

    for _, row in df_domain.iterrows():
        dom_name = str(row['도메인명']).strip() if pd.notna(row['도메인명']) else ""
        dom_group = str(row['도메인그룹화']).strip() if pd.notna(row['도메인그룹화']) else ""
        word_kr = str(row['단어명']).strip() if pd.notna(row['단어명']) else dom_name
        word_abbr = str(row['단어약어']).strip().upper() if pd.notna(row['단어약어']) else str(row['도메인약어']).strip().upper()
        
        if word_kr and word_kr.lower() != 'nan':
            domain_map_kr[word_kr] = (dom_name, dom_group)
        if word_abbr and word_abbr.lower() != 'nan':
            domain_map_abbr[word_abbr] = (dom_name, dom_group)

    # 4. 행 단위 변환 처리 함수
    def parse_row(row):
        col_physical = str(row['컬럼물리명']).strip() if pd.notna(row['컬럼물리명']) else ""
        col_logical = str(row['컬럼논리명']).strip() if pd.notna(row['컬럼논리명']) else ""
        
        std_field, dom_name, dom_group = "", "", ""
        
        # 4-1. 컬럼논리명 기반 처리
        if col_logical and col_logical.lower() != 'nan':
            matched_chunks = []
            i, name, unmatched = 0, col_logical, ""
            
            while i < len(name):
                if name[i] == ' ':
                    if unmatched:
                        matched_chunks.append(str(unmatched))
                        unmatched = ""
                    i += 1
                    continue
                matched = False
                for word, abbr in kr_word_to_abbr.items():
                    if name.startswith(word, i):
                        if unmatched:
                            matched_chunks.append(str(unmatched))
                            unmatched = ""
                        matched_chunks.append(str(abbr)) # 강제 str 변환
                        i += len(word)
                        matched = True
                        break
                if not matched:
                    unmatched += name[i]
                    i += 1
            if unmatched:
                matched_chunks.append(str(unmatched))
            
            raw_std = "_".join(matched_chunks).upper()
            raw_std = re.sub(r'[^A-Z0-9_]', '', raw_std)
            std_field = re.sub(r'_+', '_', raw_std).strip('_')
            
            for d_key in sorted(domain_map_kr.keys(), key=len, reverse=True):
                if col_logical.endswith(d_key):
                    dom_name, dom_group = domain_map_kr[d_key]
                    break
                    
        # 4-2. 컬럼물리명 기반 보완
        if not std_field and col_physical and col_physical.lower() != 'nan':
            parts = col_physical.split('_')
            converted = [str(eng_abbr_map.get(p.upper(), p.upper())) for p in parts]
            std_field = "_".join(converted)
            
        if not dom_name and col_physical and col_physical.lower() != 'nan':
            parts = col_physical.split('_')
            last_part = parts[-1].upper()
            if last_part in domain_map_abbr:
                dom_name, dom_group = domain_map_abbr[last_part]
                
        return std_field, dom_name, dom_group

    # 5. 적용 및 결과 저장
    results = df_table.apply(parse_row, axis=1)
    df_table['표준필드명'] = [r[0] for r in results]
    df_table['도메인명'] = [r[1] for r in results]
    df_table['도메인그룹화'] = [r[2] for r in results]

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_table.to_excel(writer, sheet_name='전체테이블목록', index=False)
        df_dict.to_excel(writer, sheet_name='단어사전', index=False)
        df_domain.to_excel(writer, sheet_name='도메인정의서', index=False)

    print(f"작업 완료: '{output_path}' 파일이 정상 생성되었습니다.")

# 실행
input_file = "d:\\ai\\데이터표준화_입력.xlsx"
output_file = "d:\\ai\\데이터표준화_결과.xlsx"
process_data_standardization(input_file, output_file)