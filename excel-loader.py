# 엑셀 파일을 읽어서 PostgreSQL에 업로드하는 스크립트
import win32com.client as win32     # pywin32 패키지 필요 (pip install pywin32)
import pandas as pd
import psycopg                      # psycopg 패키지 필요 (pip install psycopg[binary])
from io import StringIO


def export_drm_excel_to_csv(file_path, output_csv):
    # Excel 애플리케이션 시작
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # 백그라운드에서 실행
    excel.ScreenUpdating = False  # 화면 업데이트 비활성화
    excel.DisplayAlerts = False  # 경고 메시지 비활성화
    excel.EnableEvents = False  # 이벤트 비활성화
    excel.AskToUpdateLinks = False  # 링크 업데이트 묻는 메시지 비활성화

    try:
        # Excel 파일 열기
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.ActiveSheet    #.Sheets(1)

        # 시트의 모든 데이터를 읽어서 DataFrame으로 변환
        raw_data = worksheet.UsedRange.Value
        df = pd.DataFrame(raw_data[1:], columns=raw_data[0])  # 첫 번째 행을 헤더로 사용
        print(df)

        # DataFrame을 CSV로 저장
        df.to_csv(output_csv, index=False, encoding='utf-8-sig')
        print(f"Excel 파일이 성공적으로 CSV로 변환되었습니다: {output_csv}")

    except Exception as e:
        print(f"엑셀 파일을 읽는 중 오류가 발생하였습니다: {e}")

    finally:
        # Excel 애플리케이션 종료
        workbook.Close(SaveChanges=False)
        excel.Quit()


def upload_large_excel_to_postgres_with_copy(file_path, connection_string, copy_sql):
    # Excel 애플리케이션 시작
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # 백그라운드에서 실행
    excel.ScreenUpdating = False  # 화면 업데이트 비활성화
    excel.DisplayAlerts = False  # 경고 메시지 비활성화
    excel.EnableEvents = False  # 이벤트 비활성화
    excel.AskToUpdateLinks = False  # 링크 업데이트 묻는 메시지 비활성화

    try:
        # Excel 파일 열기
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.ActiveSheet

        excel.Calculation = -4135  # 수동 계산 모드로 설정 (자동 계산 비활성화)

        # 시트의 모든 데이터를 읽어서 DataFrame으로 변환
        raw_data = worksheet.UsedRange.Value
        df = pd.DataFrame(raw_data[1:], columns=raw_data[0])  # 첫 번째 행을 헤더로 사용

        last_row = worksheet.Cells(worksheet.Rows.Count, "A").End(-4162).Row  # A열의 마지막 행 번호
        print(f"마지막 행 번호: {last_row}")
        chunk_size = 10000

        # DataFrame을 CSV로 변환하여 메모리에 저장
        csv_buffer = StringIO()
        df.to_csv(csv_buffer, sep='\t', index=False, encoding='utf-8-sig')
        csv_buffer.seek(0)

        # PostgreSQL에 연결하여 COPY 명령어로 데이터 업로드
        with psycopg.connect(connection_string) as conn:
            with conn.cursor() as cur:
                for start_row in range(1, last_row + 1, chunk_size):
                    end_row = min(start_row + chunk_size - 1, last_row)

                    raw_data = worksheet.Range(f"A{start_row}:AX{end_row}").Value
                    df_chunk = pd.DataFrame(list(raw_data))

                    if start_row == 1:
                        df_chunk.columns = df_chunk.iloc[0]  # 첫 번째 행을 헤더로 사용
                        df_chunk = df_chunk[1:]  # 헤더 행 제거

                    output_buffer = StringIO()
                    df_chunk.to_csv(output_buffer, sep='\t', index=False, header=True, encoding='utf-8-sig')
                    output_buffer.seek(0)

                    with cur.copy(copy_sql) as copy:
                        while data := output_buffer.read(8192):
                            copy.write(data)

                conn.commit()
                print(f"데이터를 성공적으로 업로드하였습니다.")

    except Exception as e:
        print(f"엑셀 파일을 읽거나 PostgreSQL에 업로드하는 중 오류가 발생하였습니다: {e}")

    finally:
        # Excel 애플리케이션 종료
        workbook.Close(SaveChanges=False)
        excel.Quit()


def upload_excel_to_postgres_with_copy(file_path, connection_string, copy_sql):
    # Excel 애플리케이션 시작
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # 백그라운드에서 실행
    excel.ScreenUpdating = False  # 화면 업데이트 비활성화
    excel.DisplayAlerts = False  # 경고 메시지 비활성화
    excel.EnableEvents = False  # 이벤트 비활성화
    excel.AskToUpdateLinks = False  # 링크 업데이트 묻는 메시지 비활성화

    try:
        # Excel 파일 열기
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.ActiveSheet

        excel.Calculation = -4135  # 수동 계산 모드로 설정 (자동 계산 비활성화)

        # 시트의 모든 데이터를 읽어서 DataFrame으로 변환
        raw_data = worksheet.UsedRange.Value
        df = pd.DataFrame(raw_data[1:], columns=raw_data[0])  # 첫 번째 행을 헤더로 사용

        # DataFrame을 CSV로 변환하여 메모리에 저장
        csv_buffer = StringIO()
        df.to_csv(csv_buffer, sep='\t', index=False, header=True, encoding='utf-8-sig')
        csv_buffer.seek(0)

        # PostgreSQL에 연결하여 COPY 명령어로 데이터 업로드
        with psycopg.connect(connection_string) as conn:
            with conn.cursor() as cur:
                with cur.copy(copy_sql) as copy:
                    while data := csv_buffer.read(8192):
                        copy.write(data)

                conn.commit()
                print(f"데이터를 성공적으로 업로드하였습니다.")


    except Exception as e:
        print(f"엑셀 파일을 읽거나 PostgreSQL에 업로드하는 중 오류가 발생하였습니다: {e}")

    finally:
        # Excel 애플리케이션 종료
        workbook.Close(SaveChanges=False)
        excel.Quit()


def upload_excel_to_postgres_with_write_row(file_path, connection_string, copy_sql):
    # Excel 애플리케이션 시작
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # 백그라운드에서 실행
    excel.ScreenUpdating = False  # 화면 업데이트 비활성화
    excel.DisplayAlerts = False  # 경고 메시지 비활성화
    excel.EnableEvents = False  # 이벤트 비활성화
    excel.AskToUpdateLinks = False  # 링크 업데이트 묻는 메시지 비활성화

    try:
        # Excel 파일 열기
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.ActiveSheet

        excel.Calculation = -4135  # 수동 계산 모드로 설정 (자동 계산 비활성화)

        # 시트의 모든 데이터를 읽어서 DataFrame으로 변환
        raw_data = worksheet.UsedRange.Value
        # df = pd.DataFrame(raw_data[1:], columns=raw_data[0])  # 첫 번째 행을 헤더로 사용

        # PostgreSQL에 연결하여 COPY 명령어로 데이터 업로드
        with psycopg.connect(connection_string) as conn:
            with conn.cursor() as cur:
                with cur.copy(copy_sql) as copy:
                    for row in raw_data[1:]:  # 첫 번째 행은 헤더이므로 제외
                        # 튜플을 리스트로 변환 및 기본 공백 처리
                        processed_row = [None if val == "" or val is None else val for val in row]
                        
                        # 엑셀이 숫자를 float 값으로 처리하는 경우 소수점 제거 
                        for i in range(len(processed_row)):
                            val = processed_row[i]
                            
                            if isinstance(val, float):
                                # fiscal_year (인덱스 0) 처리: "2026"
                                if i == 0:
                                    processed_row[i] = str(int(val))
                                
                                # 금액 관련 컬럼 (인덱스 31 ~ 44) 처리
                                # total_amt(31), sok12(32), amt1~12(33~44)
                                elif 31 <= i <= 44:
                                    if val.is_integer(): # 소수점이 .0으로 끝나는 경우만 정수화
                                        processed_row[i] = int(val)
                        
                        copy.write_row(processed_row)

                conn.commit()
                print(f"데이터를 성공적으로 업로드하였습니다.")

    except Exception as e:
        print(f"엑셀 파일을 읽거나 PostgreSQL에 업로드하는 중 오류가 발생하였습니다: {e}")

    finally:
        # Excel 애플리케이션 종료
        workbook.Close(SaveChanges=False)
        excel.Quit()


# 실행 예시
conn_info = "dbname=postgres user=postgres password=rltjs606 host=localhost"
copy_sql = """ 
        COPY cost_ledger (fiscal_year, version_id, cost_type_code, budgeting_type,
                         mo_desc, cost_element, functional_area, acct_lv1, acct_lv2, acct_lv3, acct_lv4,
                         org_cctr, org_cctr_desc, org_job_category, org_bu_desc, mo_t, mo_t_desc, cost_center,
                         job_category, budget_unit, line, vendor, connection_name, org_cost_code, org_cost_code_name,
                         item_text, cost_code_type1, cost_code_type2, cost_code_type3, cost_code_type4, cost_code_type5,
                         total_amt, sok12, amt1, amt2, amt3, amt4, amt5, amt6, amt7, amt8, amt9, amt10, amt11, amt12,
                         create_user_id, pic_id, record_type, deletion_indicator, deleted)
        FROM STDIN WITH (FORMAT csv, DELIMITER '\t', HEADER true, ENCODING 'UTF8')
        """

copy_sql_write_row = """ 
        COPY cost_ledger (fiscal_year, version_id, cost_type_code, budgeting_type,
                         mo_desc, cost_element, functional_area, acct_lv1, acct_lv2, acct_lv3, acct_lv4,
                         org_cctr, org_cctr_desc, org_job_category, org_bu_desc, mo_t, mo_t_desc, cost_center,
                         job_category, budget_unit, line, vendor, connection_name, org_cost_code, org_cost_code_name,
                         item_text, cost_code_type1, cost_code_type2, cost_code_type3, cost_code_type4, cost_code_type5,
                         total_amt, sok12, amt1, amt2, amt3, amt4, amt5, amt6, amt7, amt8, amt9, amt10, amt11, amt12,
                         create_user_id, pic_id, record_type, deletion_indicator, deleted)
        FROM STDIN 
        """

filename = "C:\\Users\\lieph\\repos\\ledger-data-generator\\가데이터작성.xlsx"

# export_drm_excel_to_csv(filename, "output.csv")
# upload_large_excel_to_postgres_with_copy(filename, conn_info, copy_sql)
# upload_excel_to_postgres_with_copy(filename, conn_info, copy_sql)
upload_excel_to_postgres_with_write_row(filename, conn_info, copy_sql_write_row)