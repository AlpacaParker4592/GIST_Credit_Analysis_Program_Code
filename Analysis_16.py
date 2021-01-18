from os import getcwd
import sys
from pandas import read_excel
from numpy import array

program_directory = getcwd()[:getcwd().index('Design')]


# 수업 이수 여부 판별 함수
def is_completed(list1, string):
    if list1.count(string) >= 1:
        return True
    else:
        return False


def analysis(df_grade):
    df_grade_code = list(array(df_grade['교과목'].tolist()))  # 교과목 코드 리스트
    # 1. 기초교양 학점
    # 1-1. 영어
    # 영어1(GS1601 또는 GS1603) 및 영어2(GS2652) 필수 이수

    english_token = False  # 1-1. 영어 토큰
    # 미션내용
    temp_english_content = 0
    english1_token = False
    english2_token = False

    if is_completed(df_grade_code, 'GS1601') or is_completed(df_grade_code, 'GS1603'):
        # 미션내용
        temp_english_content += 1
        english1_token = True

    if is_completed(df_grade_code, 'GS1604') or is_completed(df_grade_code, 'GS2652'):
        # 미션내용
        temp_english_content += 1
        english2_token = True

    if english1_token == True and english2_token == True:
        english_token = True

    # 미션내용
    english_content = str(temp_english_content) + '/2(과목)'

    # 1-2. 글쓰기
    # 글쓰기 과목(GS1511, GS1512, G1513, GS1531, GS1532, GS1533, GS1534) 중 하나 이상 필수 이수

    writing_token = False  # 1-2. 글쓰기 토큰
    # 미션내용
    writing_content = '미이수'
    writing_code_list = ['GS1511', 'GS1512', 'GS1513', 'GS1531', 'GS1532', 'GS1533', 'GS1534']
    for writing_code in writing_code_list:
        if is_completed(df_grade_code, writing_code):
            writing_token = True
            # 미션내용
            writing_content = '이수(과목코드: ' + writing_code + ')'
            break

    # 1-3. 교양(+기초교육학부 개별연구 이수학점)
    # HUS(6학점 이상), PPE(6학점 이상), GSC 합 24학점 이상 이수 필수

    # 수강한 필수 교양 과목 리스트 반환 함수
    def now_cpltd_elective_list(obj_list, obj_year_list, now_list, now_year_list):
        completed_list = []
        now_count = 0
        for now_cpltd_lecture in now_list:
            obj_count = 0
            for obj_lecture in obj_list:
                if now_cpltd_lecture == obj_lecture and now_year_list[now_count] == obj_year_list[obj_count]:
                    completed_list.append(now_cpltd_lecture)
                    break
                obj_count += 1
            now_count += 1
        return completed_list

    # 수강한 필수 교양 과목 학점 반환 함수
    def cpltd_elective_credit(your_df_grade, your_elective_list):
        your_total_credit = 0
        for element in your_elective_list:
            your_total_credit += your_df_grade[your_df_grade['교과목'] == element]['학점'].sum()
        return your_total_credit

    elective_token = False  # 1-3. 교양 토큰
    hus_token = False  # 1-3-1. hus 토큰
    ppe_token = False  # 1-3-2. ppe 토큰
    # 각 년도마다 HUS, PPE, GSC 과목 배정이 다름 -> 각 년도 학사 편람에 따라 엑셀 파일을 참조해야 함
    # (15학번까지, 14학번 이상은 교양 단일 파일(모든 과목 GSC로 들어가게 처리))
    try:
        hus_file = program_directory + '\\edited\\elective_list\\hus.xlsx'
        df_hus = read_excel(hus_file)
        ppe_file = program_directory + '\\edited\\elective_list\\ppe.xlsx'
        df_ppe = read_excel(ppe_file)
        gsc_file = program_directory + '\\edited\\elective_list\\gsc.xlsx'
        df_gsc = read_excel(gsc_file)
    except:
        print('교양 데이터 파일이 존재하지 않습니다.')
        sys.exit(0)

    hus_code_list = list(array(df_hus['교과목'].tolist()))  # hus 코드 리스트
    hus_year_list = list(array(df_hus['년도'].tolist()))  # hus 과목 년도 리스트
    ppe_code_list = list(array(df_ppe['교과목'].tolist()))  # ppe 코드 리스트
    ppe_year_list = list(array(df_ppe['년도'].tolist()))  # ppe 과목 년도 리스트
    gsc_code_list = list(array(df_gsc['교과목'].tolist()))  # gsc 코드 리스트
    gsc_year_list = list(array(df_gsc['년도'].tolist()))  # gsc 과목 년도 리스트

    # 각 강좌별 수강 년도 리스트(교양 과목 리스트 추출용)
    cpltd_semester_list = list(array(df_grade['수강 학기'].tolist()))
    cpltd_year_list = []
    for element in cpltd_semester_list:
        year = int(element[0:4])
        cpltd_year_list.append(year)

    cpltd_hus_list = now_cpltd_elective_list(hus_code_list, hus_year_list, df_grade_code, cpltd_year_list)
    cpltd_ppe_list = now_cpltd_elective_list(ppe_code_list, ppe_year_list, df_grade_code, cpltd_year_list)
    cpltd_gsc_list = now_cpltd_elective_list(gsc_code_list, gsc_year_list, df_grade_code, cpltd_year_list)

    hus_credit = cpltd_elective_credit(df_grade, cpltd_hus_list)
    ppe_credit = cpltd_elective_credit(df_grade, cpltd_ppe_list)
    gsc_credit = cpltd_elective_credit(df_grade, cpltd_gsc_list)

    # gsc에서 기초교육학부 개별연구 이수학점 제외
    for lecture in cpltd_gsc_list:
        if lecture.find('GS920') != -1:
            gsc_credit -= 1

    if hus_credit >= 6:
        hus_token = True
    if ppe_credit >= 6:
        ppe_token = True
    elective_credit = hus_credit + ppe_credit + gsc_credit  # 교양 학점
    if hus_token == True and ppe_token == True and elective_credit >= 24:
        elective_token = True

    # gsc에서 기초교육학부 개별연구 이수학점 추가
    for lecture in cpltd_gsc_list:
        if lecture.find('GS920') != -1:
            gsc_credit += 1
    elective_credit = hus_credit + ppe_credit + gsc_credit  # 교양 학점
    if elective_credit >= 36:
        elective_credit = 36  # (최대 36학점 제한)

    # 미션내용
    hus_content = 'HUS: ' + str(hus_credit) + '/6(학점)     '
    ppe_content = 'PPE: ' + str(ppe_credit) + '/6(학점)'
    elective_content = '총 교양학점: ' + str(elective_credit) + '/24(학점)<br>(최대 36학점 인정)<br>' + hus_content + \
                       ppe_content

    # 1-4. 기초과학
    basic_sci_token = False  # 1-4. 기초과학 토큰
    # 미션내용
    basic_sci_content = ''

    # 1-4-1. 물리: 일물I(GS1101 또는 GS1103) 및 일물실(GS1111) 이수
    basic_physics_token = False  # 1-4-1. 기초물리 토큰

    if (is_completed(df_grade_code, 'GS1101') or is_completed(df_grade_code, 'GS1103')) and is_completed(df_grade_code,
                                                                                                         'GS1111'):
        # 미션내용
        basic_sci_content += '일반물리학'
        basic_physics_token = True

    # 1-4-2. 화학: 일화I(GS1201) 및 일화실(GS1211) 이수
    basic_chemistry_token = False  # 1-4-2. 기초화학 토큰

    if is_completed(df_grade_code, 'GS1201') and is_completed(df_grade_code, 'GS1211'):
        # 미션내용
        if basic_sci_content != '':
            basic_sci_content += ', '
        basic_sci_content += '일반화학'
        basic_chemistry_token = True

    # 1-4-3. 생명: 일생 또는 인생(GS1301, GS1302 또는 GS1303) 및 일생실(GS1311) 이수

    basic_biology_token = False  # 1-4-3. 기초생물학 토큰

    if (is_completed(df_grade_code, 'GS1301') or is_completed(df_grade_code, 'GS1302')) and is_completed(df_grade_code,
                                                                                                         'GS1311'):
        # 미션내용
        if basic_sci_content != '':
            basic_sci_content += ', '
        basic_sci_content += '생물학'
        basic_biology_token = True

    # 1-4-4. 전컴: 컴프(GS1401) 이수

    basic_eecs_token = False  # 1-4-4. 기초전컴 토큰

    if is_completed(df_grade_code, 'GS1401'):
        # 미션내용
        if basic_sci_content != '':
            basic_sci_content += ', '
        if len(basic_sci_content) >= 25:
            basic_sci_content += '<br>'
        basic_sci_content += '컴퓨터 프로그래밍'
        basic_eecs_token = True

    # 1-4-5. 소기코(GS1490) 이수

    basic_sbc_token = False  # 1-4-5. 소기코 토큰

    if is_completed(df_grade_code, 'GS1490'):
        # 미션내용
        if basic_sci_content != '':
            basic_sci_content += ', '
        if len(basic_sci_content) >= 25 and basic_sci_content.count('<br>') == 0:
            basic_sci_content += '<br>'
        basic_sci_content += '소프트웨어 기초와 코딩'
        basic_sbc_token = True

    # 위에서 1235, 124, 134, 234 조합 중 하나 완료 필수
    basic_sci_token_list = [basic_physics_token, basic_chemistry_token, basic_biology_token, basic_eecs_token,
                            basic_sbc_token]

    if basic_sci_token_list[0] and basic_sci_token_list[1] and basic_sci_token_list[2] and basic_sci_token_list[4]:
        basic_sci_token = True

    elif basic_sci_token_list[3]:
        if basic_sci_token_list[0] and basic_sci_token_list[1]:
            basic_sci_token = True
        elif basic_sci_token_list[0] and basic_sci_token_list[2]:
            basic_sci_token = True
        elif basic_sci_token_list[1] and basic_sci_token_list[2]:
            basic_sci_token = True

    # 1-5. 새내기: GIST 새내기 또는 신입생세미나(GS1901) 이수 필수
    freshman_token = False  # 1-5. 새내기 토큰
    # 미션내용
    freshman_content = '미이수'

    if is_completed(df_grade_code, 'GS9301') or is_completed(df_grade_code, 'GS1901'):  # GS1901은 19학번 이전용
        # 미션내용
        freshman_content = '이수'
        freshman_token = True

    # 1-6. 콜로퀴움(UC9331) 2학기 필수 이수
    colloquium_token = False  # 1-7. 콜로퀴움 토큰
    if df_grade_code.count('UC9331') >= 2:
        colloquium_token = True
    colloquium_content = str(df_grade_code.count('UC9331')) + '/2(학기)'

    # 1-7. 수학
    math_token = False
    calculus_token = False  # 미적분학
    analysis_token = False  # 다변수해석학
    equation_token = False  # 미분방정식
    algebra_token = False  # 선형대수학
    equation_algebra_token = False  # 기초미분방정식 및 선형대수
    # 미션내용
    math_content = ''
    if is_completed(df_grade_code, 'GS1001') or is_completed(df_grade_code, 'GS1011'):
        # 미션내용
        math_content += '미적분학'
        calculus_token = True
    if is_completed(df_grade_code, 'GS1002'):
        # 미션내용
        math_content += ', 다변수해석학'
        analysis_token = True
    if is_completed(df_grade_code, 'GS2002'):
        # 미션내용
        math_content += ', 미분방정식'
        equation_token = True
    if is_completed(df_grade_code, 'GS2013'):
        # 미션내용
        math_content += ', 기초미분방정식&선형대수'
        equation_algebra_token = True
    if is_completed(df_grade_code, 'GS2004'):
        # 미션내용
        math_content += ', 선형대수학'
        algebra_token = True

    if calculus_token and (
            calculus_token or analysis_token or equation_token or algebra_token or equation_algebra_token):
        math_token = True

    # 2. 연구학점
    research_token = False  # 2. 연구 토큰
    # 미션내용
    temp_research_content = 0
    # 2-1. 기계공학 전공생
    # 학논연II 또는 캡스톤II(MC9103 또는 MC9104) 및 학논연I(MC9102) 이수 필수
    # 2-2. 비기계공학 전공생
    # 학논연I(??9102) 및 학논연II(??9103) 이수 필수
    research1_token = False
    research2_token = False

    for code in df_grade_code:
        if code.find('9102') != -1:
            # 미션내용
            temp_research_content += 1
            research1_token = True
            break
    for code in df_grade_code:
        if code.find('9103') != -1 or code.find('MC9104') != -1:
            # 미션내용
            temp_research_content += 1
            research2_token = True
            break

    if research1_token and research2_token:
        research_token = True
    # 미션내용
    research_content = str(temp_research_content) + '/2(과목)'

    # 3. 대학 공통과목
    # 사회봉사(GS0201), 해외봉사(GS0203): 두 과목 모두 이수한 경우에도 최대 1학점만 인정
    # 창의함양(GS0202): 몇 번 이수해도 1학점만 인정
    volunteer_log_token = True  # 봉사활동용 조건토큰
    creative_log_token = True  # 창의함양용 조건토큰

    for i in range(0, len(df_grade)):
        # 2회차 이후의 봉사활동의 학점을 0 처리(사회봉사, 해외봉사)
        if df_grade.loc[i, '교과목'] == 'UC0201' or df_grade.loc[i, '교과목'] == 'UC0203':
            if volunteer_log_token:
                volunteer_log_token = False
            else:
                df_grade.loc[i, '학점'] = 0
        # 2회차 이후의 창의함양활동의 학점을 0 처리
        if df_grade.loc[i, '교과목'] == 'UC0202':
            if creative_log_token:
                creative_log_token = False
            else:
                df_grade.loc[i, '학점'] = 0

    # 4. 예체능: 예능(GS0201~0214) 및 체능(GS0101~0114) 각각 2학기 이상 이수 필수
    art_token = False
    art_token_count = 0

    for code in df_grade_code:
        if code.find('GS02') != -1:
            art_token_count += 1

    if art_token_count >= 4:
        art_token = True

    physical_token = False
    physical_token_count = 0

    for code in df_grade_code:
        if code.find('GS01') != -1:
            physical_token_count += 1

    if physical_token_count >= 4:
        physical_token = True
    # 미션내용
    art_content = str(art_token_count) + '/4(학기)'
    physical_content = str(physical_token_count) + '/4(학기)'

    # In[2]:

    # 전공계열 과목 목록 및 그 학점
    major_token = False
    nec_major_token = False
    my_major_credit = 0

    # 수강한 필수 교양 과목 학점 반환 함수
    def cpltd_major_credit(your_df_grade, your_major_list):
        your_total_credit = 0
        for element in your_major_list:
            your_total_credit += your_df_grade[your_df_grade['교과목'] == element]['학점'].sum()
            if your_total_credit >= 36:
                your_total_credit = 36
        return your_total_credit

    # 수강한 필수 교양 과목 리스트 반환 함수
    def now_cpltd_major_list(obj_major_code_list, your_grade_code):
        completed_list = []
        for now_cpltd_lecture in your_grade_code:
            for lecture_list in obj_major_code_list:
                for lecture in lecture_list:
                    if now_cpltd_lecture == lecture:
                        completed_list.append(now_cpltd_lecture)
                        break
        return completed_list

    info_file = program_directory + '\\edited\\info\\personalinfo2.xlsx'
    df_info = read_excel(info_file, header=None)
    my_major = df_info.values[3][1]

    major_list = ('물리전공', '화학전공', '생명과학전공', '전기전자컴퓨터전공', '기계공학전공', '신소재공학전공', '지구·환경공학전공')
    # 미션내용
    nec_major_list = []
    my_cpltd_major_list = []
    my_now_cpltd_nec_major_list = []
    # 물리 전공 ~ 지환공 전공 입력

    if my_major == major_list[0]:  # 물리전공

        my_cpltd_major_list = [x for x in df_grade_code if x[:2] == 'PS' or x == 'AP602Y']
        nec_major_list = [['PS2101', 'PS3204'], ['PS2103', 'PS3101'],
                          ['PS3103'], ['PS3104'], ['PS3105'], ['PS3106'],
                          ['PS3107']]
        # 고전역학 및 연습I, 전자기학 및 연습 II, 양자물리 및 연습I / II,
        # 열력학 및 통계물리, 물리실험 I, 수리물리 및 연습(수리물리 I)
        my_now_cpltd_nec_major_list = now_cpltd_major_list(nec_major_list, my_cpltd_major_list)
        my_major_credit = cpltd_major_credit(df_grade, my_cpltd_major_list)
        if len(my_now_cpltd_nec_major_list) >= len(nec_major_list):
            nec_major_token = True
        if nec_major_token == True and my_major_credit >= 30:
            major_token = True

    elif my_major == major_list[1]:  # 화학전공

        my_cpltd_major_list = [x for x in df_grade_code if
                               x[:2] == 'CH' or x[:2] == 'CM' or x == 'GS2203' or x == 'GS2205']
        nec_major_list = [['CH2102', 'CH3104', 'CM3104'], ['CH2201', 'CH3105', 'CM3105'],
                          ['CH2105', 'CH3102', 'CM3102'], ['CH3106', 'CM3203', 'GS2203'],
                          ['CH3103', 'CH3207', 'CM3103']]
        # 물리화학A(물리화학 II), 유기화학II,
        # 화학합성실험, 생화학I, 고급화학실험
        my_now_cpltd_nec_major_list = now_cpltd_major_list(nec_major_list, my_cpltd_major_list)
        my_major_credit = cpltd_major_credit(df_grade, my_cpltd_major_list)
        if len(my_now_cpltd_nec_major_list) >= len(nec_major_list):
            nec_major_token = True
        if nec_major_token == True and my_major_credit >= 30:
            major_token = True

    elif my_major == major_list[2]:  # 생명전공

        my_cpltd_major_list = [x for x in df_grade_code if x[:2] == 'BS']
        nec_major_list = [['BS3101'], ['BS3105'], ['BS2103', 'BS3111'],
                          ['BS3112'], ['BS2104', 'BS3113']]
        # 생화학II, 세포생물학, 생화학분자생물학실험
        # 세포발생생물학실험, 생화학I
        my_now_cpltd_nec_major_list = now_cpltd_major_list(nec_major_list, my_cpltd_major_list)
        my_major_credit = cpltd_major_credit(df_grade, my_cpltd_major_list)
        if len(my_now_cpltd_nec_major_list) >= len(nec_major_list):
            nec_major_token = True
        if nec_major_token == True and my_major_credit >= 30:
            major_token = True

    elif my_major == major_list[3]:  # 전컴전공

        my_cpltd_major_list = [x for x in df_grade_code if x[:2] == 'EC'
                               or x == "MM4010"
                               or x == "CT4201"]  # 이산수학, 그래픽스
        nec_major_list = [['EC3101', 'EC3102']]
        # 전자공학실험 or 컴시이실
        my_now_cpltd_nec_major_list = now_cpltd_major_list(nec_major_list, my_cpltd_major_list)
        my_major_credit = cpltd_major_credit(df_grade, my_cpltd_major_list)
        if len(my_now_cpltd_nec_major_list) >= len(nec_major_list):
            nec_major_token = True
        if nec_major_token == True and my_major_credit >= 30:
            major_token = True

    elif my_major == major_list[4]:  # 기계전공

        my_cpltd_major_list = [x for x in df_grade_code if x[:2] == 'MC']
        nec_major_list = [['MC2102', 'MC3105'], ['MC3103']]
        # 유체역학, 기구동역학
        my_now_cpltd_nec_major_list = now_cpltd_major_list(nec_major_list, my_cpltd_major_list)
        my_major_credit = cpltd_major_credit(df_grade, my_cpltd_major_list)
        if len(my_now_cpltd_nec_major_list) >= len(nec_major_list):
            nec_major_token = True
        if nec_major_token == True and my_major_credit >= 30:
            major_token = True

    elif my_major == major_list[5]:  # 소재전공

        my_cpltd_major_list = [x for x in df_grade_code if x[:2] == 'MA']
        nec_major_list = [['MA2101', 'MA3101'],
                          ['MA2104', 'MA3102'], ['MA3104'], ['MA3105']]
        # 재료과학, 고분자과학, 전자재료실험, 유기재료실험
        my_now_cpltd_nec_major_list = now_cpltd_major_list(nec_major_list, my_cpltd_major_list)
        my_major_credit = cpltd_major_credit(df_grade, my_cpltd_major_list)
        if len(my_now_cpltd_nec_major_list) >= len(nec_major_list):
            nec_major_token = True
        if nec_major_token == True and my_major_credit >= 30:
            major_token = True

    elif my_major == major_list[5]:  # 환경전공

        my_cpltd_major_list = [x for x in df_grade_code if x[:2] == 'EV']
        nec_major_list = [['EV3101'], ['EV3106'], ['EV3111'], ['EV3103', 'EV4106'],
                          ['EV3102', 'EV3221', 'EV4105', 'EV4218']]
        # 환경공학, 환경분석실험I, 지구시스템과학, 지구환경이동현상,지구환경열역학

        # 대기학, 해양학 모두 이수시 지구시스템과학 이수 처리
        if not 'EV3111' in my_cpltd_major_list:
            if 'EV3104' in my_cpltd_major_list and 'EV3105' in my_cpltd_major_list:
                my_cpltd_major_list.append('EV3111')

        my_now_cpltd_nec_major_list = now_cpltd_major_list(nec_major_list, my_cpltd_major_list)
        my_major_credit = cpltd_major_credit(df_grade, my_cpltd_major_list)
        if len(my_now_cpltd_nec_major_list) >= len(nec_major_list):
            nec_major_token = True
        if nec_major_token == True and my_major_credit >= 30:
            major_token = True

    # 미션내용
    if my_major == '기초교육학부':
        nec_major_content = '전공 미선언'
        major_content = '전공 미선언'
    else:
        nec_major_content = str(len(my_now_cpltd_nec_major_list)) + '/' + str(len(nec_major_list)) + '(과목)'
        major_content = str(my_major_credit) + '/30(학점)<br>(최대 36학점 인정)'

    # In[3]:

    # 비전공계열 과목 목록 및 그 학점
    df_grade_general_code = [x for x in df_grade['교과목'] if not x in my_cpltd_major_list]
    df_grade_general_code = [x for x in df_grade_general_code if not x in cpltd_hus_list]
    df_grade_general_code = [x for x in df_grade_general_code if not x in cpltd_ppe_list]
    df_grade_general_code = [x for x in df_grade_general_code if not x in cpltd_gsc_list]
    my_general_credit = cpltd_major_credit(df_grade, df_grade_general_code)

    # 총 학점(일반 학점+전공 학점(최대 36학점)+교양 학점(최대 36학점))
    total_credit = my_general_credit + my_major_credit + elective_credit

    # 달성한 졸업요건 개수 출력
    get_major_token = False
    if my_major == '기초교육학부':
        graduate_mission_list = [english_token, writing_token, elective_token, basic_sci_token,
                                 freshman_token, colloquium_token, research_token, math_token,
                                 art_token, physical_token, get_major_token]
        mission_list_text = '\t[english_token, writing_token, elective_token, basic_sci_token,\n' \
                            '\t\t\tfreshman_token, colloquium_token, research_token, math_token\n' \
                            '\t\t\tart_token, physical_token, get_major_token]'
        mission_list = ['영어 1 및 2', '글쓰기 과목', '교양 과목', '기초과학 과목', 'GIST 새내기/신입생 세미나',
                        '콜로퀴움(2학기 이상)', '학논연 과목', '수학 과목', '예능 과목(4학기 이상)',
                        '체능 과목(4학기 이상)', '전공 선언']
        mission_content_list = [english_content, writing_content, elective_content, basic_sci_content, freshman_content,
                                colloquium_content, research_content, math_content, art_content,
                                physical_content, '전공 미선언 상태']  # 미션 내용 및 달성 현황 내용 관련 리스트
    else:
        graduate_mission_list = [english_token, writing_token, elective_token, basic_sci_token,
                                 freshman_token, colloquium_token, research_token, math_token,
                                 art_token, physical_token, major_token, nec_major_token]
        mission_list_text = '\t[english_token, writing_token, elective_token, basic_sci_token,\n' \
                            '\t\t\tfreshman_token, colloquium_token, research_token, math_token, art_token,\n' \
                            '\t\t\tphysical_token, major_token, nec_major_token]'
        mission_list = ['영어 1 및 2', '글쓰기 과목', '교양 과목', '기초과학 과목', 'GIST 새내기/신입생 세미나',
                        '콜로퀴움(2학기 이상)', '학논연 과목', '수학 과목', '예능 과목(4학기 이상)',
                        '체능 과목(4학기 이상)', '전공 과목', '전공 필수 과목']
        mission_content_list = [english_content, writing_content, elective_content, basic_sci_content, freshman_content,
                                colloquium_content, research_content, math_content, art_content,
                                physical_content, major_content, nec_major_content]  # 미션 내용 및 달성 현황 내용 관련 리스트
    token_count = 0
    for token in graduate_mission_list:
        if token == True:
            token_count += 1

    print('Total Credit: ' + str(total_credit) + '/130 credit(s)')
    print('Missions:' + mission_list_text)
    print('Completed Missions: ' + str(token_count) + '/' + str(len(graduate_mission_list)) + ' mission(s)')

    return my_general_credit, my_major_credit, elective_credit, graduate_mission_list, my_cpltd_major_list, mission_list, mission_content_list
