import json
import pandas as pd
from datetime import datetime, timedelta

# 📌 엑셀 파일에서 "2025第2期" 시트를 로드
excel_file = "templates/template.xlsx"
sheet_name = "2025第2期"
df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

# 📌 1주차 시작 날짜 설정
start_date = datetime(2025, 4, 5)

# 📌 셀 값 정리 함수 (개행문자 → " / "로 변환)
def clean_cell(value):
    if pd.isna(value):
        return None
    return str(value).replace("\n", " / ").replace("\r", " / ")

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
            "reception": clean_cell(df.iloc[row_number - 1, 2]),
            "pianist": clean_cell(df.iloc[row_number - 1, 3]),
            "greetings": clean_cell(df.iloc[row_number - 1, 5]),
            "songService": None,
            "miniPrayerTime": clean_cell(df.iloc[row_number - 1, 5]),
            "memoryText": None,
            "lessonStudy": None,
            "program": clean_cell(df.iloc[row_number - 1, 6]),
            "breakTime": None,
            "announcements": None
        },
        "worshipService": {
            "pianist": clean_cell(df.iloc[row_number - 1, 3]),
            "translator": clean_cell(df.iloc[row_number - 1, 11]),
            "presider": clean_cell(df.iloc[row_number - 1, 8]),
            "hymn": None,
            "openingPrayer": None,
            "openingSong": None,
            "scriptureReading": None,
            "specialMusic": clean_cell(df.iloc[row_number - 1, 9]),
            "preacher": clean_cell(df.iloc[row_number - 1, 10]),
            "sermonTitleJP": None,
            "sermonTitleEN": None,
            "offeringPromotion": None,
            "offeringService": clean_cell(df.iloc[row_number - 1, 13]),
            "offeringPrayer": clean_cell(df.iloc[row_number - 1, 14]),
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

        if week == 13 and i >= 1:
            schedule_data[key] = None
        elif week == 12 and i == 2:
            schedule_data[key] = None
        else:
            row_index = row_number + i - 1
            schedule_data[key] = {
                "rowNumber": str(row_number + i),
                "date": (start_date + timedelta(weeks=i)).strftime("%m/%d"),
                "reception": clean_cell(df.iloc[row_index, 2]),
                "pianist": clean_cell(df.iloc[row_index, 3]),
                "greetings": clean_cell(df.iloc[row_index, 5]),
                "program": clean_cell(df.iloc[row_index, 6]),
                "announcementTranslator": clean_cell(df.iloc[row_index, 7]),
                "specialMusic": clean_cell(df.iloc[row_index, 9]),
                "preacher": clean_cell(df.iloc[row_index, 10]),
                "SermonTranslator": clean_cell(df.iloc[row_index, 11]),
                "offeringService": clean_cell(df.iloc[row_index, 13]),
                "offeringPrayer": clean_cell(df.iloc[row_index, 14])
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
