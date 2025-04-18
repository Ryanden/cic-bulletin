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
            "reception": df.iloc[row_number - 1, 2] if pd.notna(df.iloc[row_number - 1, 2]) else None,
            "pianist": df.iloc[row_number - 1, 3] if pd.notna(df.iloc[row_number - 1, 3]) else None,
            "greetings": df.iloc[row_number - 1, 5] if pd.notna(df.iloc[row_number - 1, 5]) else None,
            "songService": None,
            "miniPrayerTime": df.iloc[row_number - 1, 5] if pd.notna(df.iloc[row_number - 1, 5]) else None,
            "memoryText": None,
            "lessonStudy": None,
            "program": df.iloc[row_number - 1, 6] if pd.notna(df.iloc[row_number - 1, 6]) else None,
            "breakTime": None,
            "announcements": None
        },
        "worshipService": {
            "pianist": df.iloc[row_number - 1, 3] if pd.notna(df.iloc[row_number - 1, 3]) else None,
            "translator": df.iloc[row_number - 1, 11] if pd.notna(df.iloc[row_number - 1, 11]) else None,
            "presider": df.iloc[row_number - 1, 8] if pd.notna(df.iloc[row_number - 1, 8]) else None,
            "hymn": None,
            "openingPrayer": None,
            "openingSong": None,
            "scriptureReading": None,
            "specialMusic": df.iloc[row_number - 1, 9] if pd.notna(df.iloc[row_number - 1, 9]) else None,
            "preacher": df.iloc[row_number - 1, 10] if pd.notna(df.iloc[row_number - 1, 10]) else None,
            "sermonTitleJP": None,
            "sermonTitleEN": None,
            "offeringPromotion": None,
            "offeringService": df.iloc[row_number - 1, 13] if pd.notna(df.iloc[row_number - 1, 13]) else None,
            "offeringPrayer": df.iloc[row_number - 1, 14] if pd.notna(df.iloc[row_number - 1, 14]) else None,
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

    for i in range(3):
        key = f"{i+1}st" if i == 0 else f"{i+1}nd" if i == 1 else f"{i+1}rd"

        # ✅ 13주차 → schedule만 null 처리 (sabbathSchool, worshipService 유지)
        if week == 13 and i >= 1:
            schedule_data[key] = None
        # ✅ 12주차 → 3rd만 null 처리
        elif week == 12 and i == 2:
            schedule_data[key] = None
        else:
            row_index = row_number + i
            schedule_data[key] = {
                "rowNumber": str(row_index + 1),
                "date": (start_date + timedelta(weeks=i + 1)).strftime("%m/%d"),
                "reception": df.iloc[row_index, 2].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 2]) else None,
                "pianist": df.iloc[row_index, 3].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 3]) else None,
                "greetings": df.iloc[row_index, 5].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 5]) else None,
                "program": df.iloc[row_index, 6].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 6]) else None,
                "announcementTranslator": df.iloc[row_index, 7].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 7]) else None,
                "presider": df.iloc[row_index, 8].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 8]) else None,
                "specialMusic": df.iloc[row_index, 9].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 9]) else None,
                "preacher": df.iloc[row_index, 10].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 10]) else None,
                "SermonTranslator": df.iloc[row_index, 11].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 11]) else None,
                "offeringService": df.iloc[row_index, 13].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 13]) else None,
                "offeringPrayer": df.iloc[row_index, 14].replace("\n", " / ").replace("\r", " / ") if pd.notna(df.iloc[row_index, 14]) else None
            }

    week_data["schedule"] = schedule_data
    final_data[week_key] = week_data

    row_number += 1
    start_date += timedelta(weeks=1)

# 📌 JSON 파일로 저장
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(final_data, f, indent=4, ensure_ascii=False)

print("✅ data.json 생성 완료!")
