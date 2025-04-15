import openpyxl
import os

def save_to_excel(data, filename='business_cards.xlsx'):
    if os.path.exists(filename):
        workbook = openpyxl.load_workbook(filename)
        sheet = workbook.active
    else:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["회사", "이름", "부서", "직급", "이메일", "전화번호"])

    sheet.append([
        data["회사"], data["이름"], data["부서"],
        data["직급"], data["이메일"], data["전화번호"]
    ])
    workbook.save(filename)

def input_business_card():
    print("명함 정보를 입력하세요.")
    company = input("회사명: ")
    name = input("이름: ")
    department = input("부서: ")
    position = input("직급: ")
    email = input("이메일: ")
    phone = input("전화번호: ")

    return {
        "회사": company,
        "이름": name,
        "부서": department,
        "직급": position,
        "이메일": email,
        "전화번호": phone
    }

if __name__ == "__main__":
    while True:
        data = input_business_card()
        save_to_excel(data)
        cont = input("계속 입력하시겠습니까? (y/n): ")
        if cont.lower() != 'y':
            print("엑셀 저장을 완료하고 프로그램을 종료합니다.")
            break
