import json
import pandas as pd
from datetime import datetime, timedelta

# 📌 엑셀 파일에서 "2025第2期" 시트를 로드
excel_file = "templates/template.xlsx"
sheet_name = "2025第2期"
df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

# 📌 1주차 시작 날짜 설정
start_date = datetime(2025, 4, 5)

# 📌 전체 데이터 저장할 딕셔너리
final_data = {}

# 📌 rowNumber 시작 값 (6부터 18까지 증가)
row_number = 6  

for week in range(1, 14):  # 1thWeek ~ 13thWeek
    week_key = f"{week}thWeek"

    # 📌 sabbathSchool, worshipService, common 데이터 유지
    week_data = {
        "config": {"rowNumber": row_number},
        "sabbathSchool": {
            "reception": df.iloc[row_number - 1, 2] if pd.notna(df.iloc[row_number - 1, 2]) else None,  # C열
            "pianist": df.iloc[row_number - 1, 3] if pd.notna(df.iloc[row_number - 1, 3]) else None,  # D열
            "greetings": df.iloc[row_number - 1, 5] if pd.notna(df.iloc[row_number - 1, 5]) else None,  # F열
            "songService": None,
            "miniPrayerTime": df.iloc[row_number - 1, 5] if pd.notna(df.iloc[row_number - 1, 5]) else None,  # F열
            "memoryText": None,
            "lessonStudy": None,
            "program": df.iloc[row_number - 1, 6] if pd.notna(df.iloc[row_number - 1, 6]) else None,  # G열
            "breakTime": None,
            "announcements": None
        },
        "worshipService": {
            "pianist": df.iloc[row_number - 1, 3] if pd.notna(df.iloc[row_number - 1, 3]) else None,  # D열
            "translator": df.iloc[row_number - 1, 11] if pd.notna(df.iloc[row_number - 1, 11]) else None,  # L열
            "presider": df.iloc[row_number - 1, 8] if pd.notna(df.iloc[row_number - 1, 8]) else None,  # I열
            "hymn": None,
            "openingPrayer": None,
            "openingSong": None,
            "scriptureReading": None,
            "specialMusic": df.iloc[row_number - 1, 9] if pd.notna(df.iloc[row_number - 1, 9]) else None,  # J열
            "preacher": df.iloc[row_number - 1, 10] if pd.notna(df.iloc[row_number - 1, 10]) else None,  # K열
            "sermonTitleJP": None,
            "sermonTitleEN": None,
            "offeringPromotion": None,
            "offeringService": df.iloc[row_number - 1, 13] if pd.notna(df.iloc[row_number - 1, 13]) else None,  # N열
            "offeringPrayer": df.iloc[row_number - 1, 14] if pd.notna(df.iloc[row_number - 1, 14]) else None,  # O열
            "closingSong": None,
            "closingPrayer": None
        },
        "common": {
            "sunset": None,
            "titleJP": None,
            "titleEN": None,
            "memoryTextJP": None,
            "memoryTextEN": None,
            "scriptureJP": None,
            "scriptureEN": None,
            "meditationJP": None,
            "meditationEN": None,
            "quizQuestion": None,
            "quizAnswer": None
        }
    }

    # 📌 schedule 데이터 처리
    schedule_data = {}

    for i in range(3):  # 1st, 2nd, 3rd
        key = f"{i+1}st" if i == 0 else f"{i+1}nd" if i == 1 else f"{i+1}rd"

        # ✅ 13주차 → schedule만 null 처리 (sabbathSchool, worshipService 유지)
        if week == 13 and i >= 1:
            schedule_data[key] = None
        # ✅ 12주차 → 3rd만 null 처리
        elif week == 12 and i == 2:
            schedule_data[key] = None
        else:
            # ✅ 정상적인 주차 데이터 입력
            row_index = row_number + i - 1
            schedule_data[key] = {
                "rowNumber": str(row_number + i),
                "date": (start_date + timedelta(weeks=i)).strftime("%m/%d"),
                "reception": df.iloc[row_index, 2] if pd.notna(df.iloc[row_index, 2]) else None,  # C열
                "pianist": df.iloc[row_index, 3] if pd.notna(df.iloc[row_index, 3]) else None,  # D열
                "greetings": df.iloc[row_index, 5] if pd.notna(df.iloc[row_index, 5]) else None,  # F열
                "program": df.iloc[row_index, 6] if pd.notna(df.iloc[row_index, 6]) else None,  # G열
                "announcementTranslator": df.iloc[row_index, 7] if pd.notna(df.iloc[row_index, 7]) else None,  # H열
                "specialMusic": df.iloc[row_index, 9] if pd.notna(df.iloc[row_index, 9]) else None,  # J열
                "preacher": df.iloc[row_index, 10] if pd.notna(df.iloc[row_index, 10]) else None,  # K열
                "SermonTranslator": df.iloc[row_index, 11] if pd.notna(df.iloc[row_index, 11]) else None,  # L열
                "offeringService": df.iloc[row_index, 13] if pd.notna(df.iloc[row_index, 13]) else None,  # N열
                "offeringPrayer": df.iloc[row_index, 14] if pd.notna(df.iloc[row_index, 14]) else None  # O열
            }

    # 📌 week_data에 schedule 추가 후 저장
    week_data["schedule"] = schedule_data
    final_data[week_key] = week_data

    # 📌 rowNumber & start_date 업데이트
    row_number += 1
    start_date += timedelta(weeks=1)

# 📌 JSON 파일로 저장
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(final_data, f, indent=4, ensure_ascii=False)

print("✅ data.json 생성 완료!") 
