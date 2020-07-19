# excels-python
Control excel file with python library. Can be applied with the ability to automatically collect answers in questionnaires

# excel 파일을 이용하여 질문서를 취합하는 경우 개발도구의 Combo box, Drop box, Radio button 등을 사용하면, 사용자 UX가 좋아지게 된다.
이 경우에 단순 엑셀의 Cell 참조 방식으로는 선택한 값을 제대로 읽어서 저장하는 등의 업무를 수행할 수 없다.
그래서 MS의 COM Object를 이용하는 win32com.client, xlwt 를 이용하여 Excel의 workbook, worksheet, shapes를 다루는 기술 특징을 정리하였다.

Radio button은 그룹을 별도로 구성을 할 수 있는지?는 확인을 못해서, 예시로 만든 Radio Button은 전체가 연계되어 체크되는 문제점은 활용에서 참고하기 바란다.

# win32com library는 pip install pywin32를 이용해서 설치를 한다.
# xlwt는 pip install xlwt를 이용해서 설치를 한다.

# 해당 프로그램은 미리 정의해 놓은 Objects의 Name 규칙을 이용해서 처리하도록 되어 있기 때문에,
  excel_objects.xlsx 의 Object 이름과 소스에서 참고하고 있는 이름을 비교하면 쉽게 이해를 할 수 있을 것이다.
  
# 개발소스의 기본 처리 흐름을 이해한다면, 확장 / 응용을 다양하게 할 수 있을 것으로 판단한다.

## 개발도구의 Object를 사용하지 않는 경우는 openpyxl library를 사용하면 보다 쉽게 처리할 수 있고, 참고 Site도 많으니 필요한 상황에 맞게 선택하면 좋다.

