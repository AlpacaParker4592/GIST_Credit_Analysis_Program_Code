from os import getcwd, path
from sys import exit
from pandas import read_excel


def executeAnalyse():
    elective_list = ('gsc', 'hus', 'ppe')
    # 성적 파일 및 교과목 파일 불러오기(없으면 종료)
    program_directory = getcwd()[:getcwd().index('Design')]

    grade_file = program_directory + 'edited\\grade\\grade_edited.xlsx'
    if path.isfile(grade_file) == 0:
        print('성적 파일(grade_edited.xlsx)이 존재하지 않습니다.')
        exit(-1)

    present_subject = program_directory + 'edited\\grade\\present_subject_info.xlsx'
    if path.isfile(present_subject) == 0:
        print('현재 수강 파일(present_subject_info.xlsx)이 존재하지 않습니다.')
        exit(-1)

    info_file = program_directory + 'edited\\info\\personalinfo2.xlsx'
    if path.isfile(info_file) == 0:
        print('현재 수강 파일(present_subject_info.xlsx)이 존재하지 않습니다.')
        exit(-1)

    for elective_file_name in elective_list:
        elective_file = program_directory + 'edited\\elective_list\\' + elective_file_name + '.xlsx'
        if path.isfile(elective_file) == 0:
            print('교양 과목 목록 파일(' + elective_file_name + '.xlsx)이 존재하지 않습니다.')
            exit(-1)

    # 이수 정의: 성적표에 F 또는 U가 산출되는 과목은 이수 완료된 과목
    # 최종 학점: 130학점 이상, 평점평균 2.0 이상
    # 성적 테이블 데이터
    df_grade = read_excel(grade_file)
    df_grade = df_grade[~df_grade.평점.isin(['F', 'U'])]  # F 또는 U 등급이 나온 과목 제거
    credit_list = df_grade['평점']
    credit_number_list = []  # 평점 숫자 환산치 리스트

    for credit in credit_list:
        if credit.find('S') != -1 or credit.find('U') != -1:
            credit_number_list.append('-')
        elif credit == 'A+':
            credit_number_list.append(4.5)
        elif credit == 'A0':
            credit_number_list.append(4)
        elif credit == 'B+':
            credit_number_list.append(3.5)
        elif credit == 'B0':
            credit_number_list.append(3)
        elif credit == 'C+':
            credit_number_list.append(2.5)
        elif credit == 'C0':
            credit_number_list.append(2)
        elif credit == 'D+':
            credit_number_list.append(1.5)
        elif credit == 'D0':
            credit_number_list.append(1)
        elif credit == 'F':
            credit_number_list.append(0)
        else:
            credit_number_list.append('-')

    df_grade['평점 환산치'] = credit_number_list

    # 재수강 과목 학점 갱신 처리
    temp_df_grade = df_grade.sort_values(by=['교과목', '평점 환산치'], axis=0).reset_index(drop=True)
    for i in range(1, len(temp_df_grade)):
        if temp_df_grade['교과목'][i] == temp_df_grade['교과목'][i - 1]:
            if temp_df_grade['평점 환산치'][i] > temp_df_grade['평점 환산치'][i - 1]:
                temp_df_grade = temp_df_grade.drop[i - 1]

    df_grade = temp_df_grade.sort_values(by='수강 학기', axis=0).reset_index(drop=True)

    # 현재 수강 중인 과목 포함 과목 정보 리스트
    subject_file = program_directory + 'edited\\grade\\present_subject_info.xlsx'
    df_subject = read_excel(subject_file)

    # 학점 및 학생 정보 수집
    import Analysis_20, Analysis_19, Analysis_18, Analysis_17, Analysis_16, Analysis_15, Analysis_14

    info_file = program_directory + 'edited\\info\\personalinfo2.xlsx'
    df_info = read_excel(info_file, header=None)
    my_number = str(df_info.values[0][1])
    my_major = df_info.values[3][1]
    my_entered_year = my_number[0:4]

    my_general_credit = 0
    my_major_credit = 0
    my_elective_credit = 0
    my_mission_list = []
    my_major_list = []
    my_mission_text_list = []
    my_mission_content_list = []

    if my_entered_year == '2020':
        my_general_credit, my_major_credit, my_elective_credit, my_mission_list, my_major_list, my_mission_text_list, my_mission_content_list = Analysis_20.analysis(
            df_subject)
    elif my_entered_year == '2019':
        my_general_credit, my_major_credit, my_elective_credit, my_mission_list, my_major_list, my_mission_text_list, my_mission_content_list = Analysis_19.analysis(
            df_subject)
    elif my_entered_year == '2018':
        my_general_credit, my_major_credit, my_elective_credit, my_mission_list, my_major_list, my_mission_text_list, my_mission_content_list = Analysis_18.analysis(
            df_subject)
    elif my_entered_year == '2017':
        my_general_credit, my_major_credit, my_elective_credit, my_mission_list, my_major_list, my_mission_text_list, my_mission_content_list = Analysis_17.analysis(
            df_subject)
    elif my_entered_year == '2016':
        my_general_credit, my_major_credit, my_elective_credit, my_mission_list, my_major_list, my_mission_text_list, my_mission_content_list = Analysis_16.analysis(
            df_subject)
    elif my_entered_year == '2015':
        my_general_credit, my_major_credit, my_elective_credit, my_mission_list, my_major_list, my_mission_text_list, my_mission_content_list = Analysis_15.analysis(
            df_subject)
    elif my_entered_year == '2014':
        my_general_credit, my_major_credit, my_elective_credit, my_mission_list, my_major_list, my_mission_text_list, my_mission_content_list = Analysis_14.analysis(
            df_subject)

    my_total_credit = my_major_credit + my_elective_credit + my_general_credit
    my_true_count = 0
    for token in my_mission_list:
        if token:
            my_true_count += 1

    # 평점 계산
    import math

    df_grade = read_excel(grade_file)
    df_grade = df_grade.loc[~df_grade['평점'].isin(['S', 'U'])]  # S 또는 U 등급이 나온 과목(패논패 과목) 제거
    credit_list = df_grade['평점']
    credit_number_list = []  # 평점 숫자 환산치 리스트

    for credit in credit_list:
        if credit == 'A+':
            credit_number_list.append(4.5)
        elif credit == 'A0':
            credit_number_list.append(4)
        elif credit == 'B+':
            credit_number_list.append(3.5)
        elif credit == 'B0':
            credit_number_list.append(3)
        elif credit == 'C+':
            credit_number_list.append(2.5)
        elif credit == 'C0':
            credit_number_list.append(2)
        elif credit == 'D+':
            credit_number_list.append(1.5)
        elif credit == 'D0':
            credit_number_list.append(1)
        elif credit == 'F':
            credit_number_list.append(0)

    df_grade['평점 환산치'] = credit_number_list

    # 재수강 과목 학점 갱신 처리
    temp_df_grade = df_grade.sort_values(by=['교과목', '평점 환산치'], axis=0).reset_index(drop=True)
    for i in range(1, len(temp_df_grade)):
        if temp_df_grade['교과목'][i] == temp_df_grade['교과목'][i - 1]:
            if temp_df_grade['평점 환산치'][i] > temp_df_grade['평점 환산치'][i - 1]:
                temp_df_grade = temp_df_grade.drop[i - 1]

    df_grade = temp_df_grade.sort_values(by='수강 학기', axis=0).reset_index(drop=True)

    average_credit_list = []
    semester_list = df_grade['수강 학기수'].values.tolist()

    temp_credit_list = []
    temp_credit_count_list = []
    if len(semester_list) != 0:  # 성적표에 기록된 과목이 있을 경우
        for i in range(0, len(semester_list)):
            if len(temp_credit_list) == 0:
                temp_credit_list.append(df_grade['평점 환산치'][i] * df_grade['학점'][i])
                temp_credit_count_list.append(df_grade['학점'][i])
            elif semester_list[i] == semester_list[i - 1]:
                temp_credit_list.append(df_grade['평점 환산치'][i] * df_grade['학점'][i])
                temp_credit_count_list.append(df_grade['학점'][i])
            elif semester_list[i] != semester_list[i - 1]:
                average = 0
                credit_count = 0
                for element in temp_credit_list:
                    average += element
                for element in temp_credit_count_list:
                    credit_count += element
                average /= credit_count
                average_credit_list.append(math.floor(average * 100) / 100)
                temp_credit_list = [df_grade['평점 환산치'][i] * df_grade['학점'][i]]
                temp_credit_count_list = [df_grade['학점'][i]]

            if i == len(semester_list) - 1:
                average = 0
                credit_count = 0
                for element in temp_credit_list:
                    average += element
                for element in temp_credit_count_list:
                    credit_count += element
                average /= credit_count
                average_credit_list.append(math.floor(average * 100) / 100)
    else:  # 성적표에 기록된 과목이 없을 경우(+전부 S/U 과목만 있을 경우)
        average_credit_list.append(0)

    average_major_credit_list = []

    df_major_grade = df_grade[df_grade.교과목.isin(my_major_list)].reset_index(drop=True)  # 전공과목만 추출
    temp_credit_list = []
    temp_credit_count_list = []
    major_semester_list = df_major_grade['수강 학기수'].values.tolist()
    if len(major_semester_list) != 0:  # 기초교육학부가 아닌 경우
        for i in range(1, major_semester_list[0]):
            average_major_credit_list.append(None)

        for i in range(0, len(major_semester_list)):
            if len(temp_credit_list) == 0:
                temp_credit_list.append(df_major_grade['평점 환산치'][i] * df_major_grade['학점'][i])
                temp_credit_count_list.append(df_major_grade['학점'][i])
            elif major_semester_list[i] == major_semester_list[i - 1]:
                temp_credit_list.append(df_major_grade['평점 환산치'][i] * df_major_grade['학점'][i])
                temp_credit_count_list.append(df_major_grade['학점'][i])
            elif major_semester_list[i] != major_semester_list[i - 1]:
                average = 0
                credit_count = 0
                for element in temp_credit_list:
                    average += element
                for element in temp_credit_count_list:
                    credit_count += element
                average /= credit_count
                average_major_credit_list.append(math.floor(average * 100) / 100)
                temp_credit_list = [df_major_grade['평점 환산치'][i] * df_major_grade['학점'][i]]
                temp_credit_count_list = [df_major_grade['학점'][i]]

            if i == len(major_semester_list) - 1:
                average = 0
                credit_count = 0
                for element in temp_credit_list:
                    average += element
                for element in temp_credit_count_list:
                    credit_count += element
                average /= credit_count
                average_major_credit_list.append(math.floor(average * 100) / 100)

    # 데이터프레임 제작
    from pandas import DataFrame as df

    # First Graph
    if my_total_credit >= 130:
        f_outer_index_label = ['전공학점', '교양학점', '일반/기타학점']
        f_outer_label = [my_major_credit, my_elective_credit, my_general_credit]
    else:
        f_outer_index_label = ['전공학점', '교양학점', '일반/기타학점', '미취득학점']
        my_left_credit = 130 - my_major_credit - my_elective_credit - my_general_credit
        f_outer_label = [my_major_credit, my_elective_credit, my_general_credit, my_left_credit]

    f_inner_index_label = my_mission_text_list
    f_inner_tf_data = my_mission_list  # 안쪽 그래프 색상 결정 데이타
    f_inner_formal_data = [1] * len(my_mission_list)

    df_f_outer_graph = df(data={'이수 학점 분포': f_outer_label}, index=f_outer_index_label)
    df_f_inner_graph = df(data={'졸업요건 현황': f_inner_formal_data}, index=f_inner_index_label)

    # Second Graph
    s_label = []
    semester_list = df_grade['수강 학기'].values.tolist()
    for semester_name in semester_list:
        semester_name2 = semester_name.replace(" ", "<br>")  # x축 라벨 띄어쓰기를 엔터로 바꾸기
        if len(s_label) == 0:
            s_label.append(semester_name2)
        elif s_label[-1] != semester_name2:
            s_label.append(semester_name2)

    if len(average_major_credit_list) != 0:  # 수강한 전공 과목이 있을 때
        df_s_graph = df(data={'학기별 전체 평점': average_credit_list, '학기별 전공 평점': average_major_credit_list}, index=s_label)
    else:  # 없다면
        df_s_graph = df(data={'학기별 전체 평점': average_credit_list}, index=s_label)

    # Completed/Incomplemented Mission
    df_mission = df_f_inner_graph
    df_mission['이수 여부'] = my_mission_list
    df_mission['이수 현황'] = my_mission_content_list

    df_completed_mission_list = df_mission[df_mission['이수 여부'] == True].index.tolist()
    df_incompleted_mission_list = df_mission[df_mission['이수 여부'] == False].index.tolist()
    df_completed_mission_list = df(data={'졸업 요건': df_completed_mission_list})
    df_incompleted_mission_list = df(data={'졸업 요건': df_incompleted_mission_list})
    df_completed_mission_list['이수 현황'] = df_mission[df_mission['이수 여부'] == True]['이수 현황'].tolist()
    df_incompleted_mission_list['이수 현황'] = df_mission[df_mission['이수 여부'] == False]['이수 현황'].tolist()
    df_completed_mission_list['번호'] = [str(x) for x in range(1, len(df_completed_mission_list) + 1)]
    df_incompleted_mission_list['번호'] = [str(x) for x in range(1, len(df_incompleted_mission_list) + 1)]

    # 표 외관 균형을 위해 번호 및 졸업 요건 뒤에 <br> 붙이기
    for row in range(0, len(df_completed_mission_list)):
        for i in range(0, df_completed_mission_list.iloc[row, 1].count('<br>')):  # 이수 현황
            # print(df_completed_mission_list[row, 1])
            df_completed_mission_list.iloc[row, 2] = df_completed_mission_list.iloc[row, 2] + '<br>'  # 번호
            df_completed_mission_list.iloc[row, 0] = df_completed_mission_list.iloc[row, 0] + '<br>'  # 졸업 요건
    for row in range(0, len(df_incompleted_mission_list)):
        for i in range(0, df_incompleted_mission_list.iloc[row, 1].count('<br>')):  # 이수 현황
            df_incompleted_mission_list.iloc[row, 2] = df_incompleted_mission_list.iloc[row, 2] + '<br>'  # 번호
            df_incompleted_mission_list.iloc[row, 0] = df_incompleted_mission_list.iloc[row, 0] + '<br>'  # 졸업 요건

    # import plotly.offline as pyo
    import plotly.graph_objs as go

    def iscompletedmission(tf_list, true_value, false_value):
        converted_list = []
        for tf_data in tf_list:
            if tf_data:
                converted_list.append(true_value)
            else:
                converted_list.append(false_value)
        return converted_list

    # 그래프 제작

    # 학점 수에 따른 색상 테마를 정하기 위한 리스트
    source1 = [160, 80, 240]  # 외부 그래프 색상(순서대로 RGB)
    source1_sub1 = [96, 61, 217]
    source1_sub2 = [69, 77, 248]
    line_source = [155, 61, 217]

    def DetermineColorTheme(rgb_list, differ, intercept, color_parameter):  # intercept: 0 <= x < 1
        converted_intercept = intercept * color_parameter / 12
        for i in range(0, my_total_credit):
            if i < color_parameter / 12 - converted_intercept:
                rgb_list[0] -= differ / color_parameter * 12
            if color_parameter / 12 - converted_intercept <= i < color_parameter / 12 * 2 - converted_intercept:
                rgb_list[1] += differ / color_parameter * 12
            if color_parameter / 12 * 2 - converted_intercept <= i < color_parameter / 12 * 3 - converted_intercept:
                rgb_list[1] += differ / color_parameter * 12
            if color_parameter / 12 * 3 - converted_intercept <= i < color_parameter / 12 * 4 - converted_intercept:
                rgb_list[2] -= differ / color_parameter * 12
            if color_parameter / 12 * 4 - converted_intercept <= i < color_parameter / 12 * 5 - converted_intercept:
                rgb_list[2] -= differ / color_parameter * 12
            if color_parameter / 12 * 5 - converted_intercept <= i < color_parameter / 12 * 6 - converted_intercept:
                rgb_list[0] += differ / color_parameter * 12
            if color_parameter / 12 * 6 - converted_intercept <= i < color_parameter / 12 * 7 - converted_intercept:
                rgb_list[0] += differ / color_parameter * 12
            if color_parameter / 12 * 7 - converted_intercept <= i < color_parameter / 12 * 8 - converted_intercept:
                rgb_list[1] -= differ / color_parameter * 12
            if color_parameter / 12 * 8 - converted_intercept <= i < color_parameter / 12 * 9 - converted_intercept:
                rgb_list[1] -= differ / color_parameter * 12
            if color_parameter / 12 * 9 - converted_intercept <= i < color_parameter / 12 * 10 - converted_intercept:
                rgb_list[2] += differ / color_parameter * 12
            if color_parameter / 12 * 10 - converted_intercept <= i < color_parameter / 12 * 11 - converted_intercept:
                rgb_list[2] += differ / color_parameter * 12
            if color_parameter / 12 * 11 - converted_intercept <= i < color_parameter - converted_intercept:
                rgb_list[0] -= differ / color_parameter * 12
        return rgb_list

    color_parameter = 195
    after_source1 = DetermineColorTheme(source1, 80, 0, color_parameter)
    after_source1_sub1 = DetermineColorTheme(source1_sub1, 78, 0, color_parameter)  # 1/2
    after_source1_sub2 = DetermineColorTheme(source1_sub2, 80, 0, color_parameter)  # 1/2-1/179
    after_line_source = DetermineColorTheme(line_source, 78, 0, color_parameter)  # 1/2

    present_color = 'rgb(' + str(int(after_source1[0]) % 256) + ',' + str(
        int(after_source1[1]) % 256) + ',' + str(int(after_source1[2]) % 256) + ')'
    sub1_present_color = 'rgb(' + str(int(after_source1_sub1[0]) % 256) + ',' + str(
        int(after_source1_sub1[1]) % 256) + ',' + str(int(after_source1_sub1[2]) % 256) + ')'
    sub2_present_color = 'rgb(' + str(int(after_source1_sub2[0]) % 256) + ',' + str(
        int(after_source1_sub2[1]) % 256) + ',' + str(int(after_source1_sub2[2]) % 256) + ')'
    line_present_color = 'rgb(' + str(int(after_line_source[0]) % 256) + ',' + str(
        int(after_line_source[1]) % 256) + ',' + str(int(after_line_source[2]) % 256) + ')'

    outer_color_list = [present_color, sub1_present_color, sub2_present_color, 'rgba(255,255,255,0)']
    outer_line_list = [present_color, sub1_present_color, sub2_present_color, line_present_color]
    title_font_size = 30

    # 현재 이수한 학점 및 졸업요건
    # graph1 - inner graph
    iscompleted_list = iscompletedmission(df_mission['이수 여부'].tolist(), '<b>%{label}</b>:<br>완료</br><extra></extra>',
                                          '<b>%{label}</b>:<br>미완료</br><extra></extra>')
    lighter_red = int(after_source1_sub1[0]) % 256 + (255 - int(after_source1_sub1[0]) % 256) * 4 / 5
    lighter_green = int(after_source1_sub1[1]) % 256 + (255 - int(after_source1_sub1[1]) % 256) * 4 / 5
    lighter_blue = int(after_source1_sub1[2]) % 256 + (255 - int(after_source1_sub1[2]) % 256) * 4 / 5
    lighter_color = 'rgb(' + str(lighter_red) + ',' + str(int(lighter_green)) + ',' + str(lighter_blue) + ')'
    tf_color_list = iscompletedmission(df_mission['이수 여부'].tolist(), sub1_present_color, lighter_color)
    graph1_data1 = go.Pie(
        hole=0.7,
        sort=False,
        direction='clockwise',
        domain={'x': [0.1, 0.9], 'y': [0.15, 0.85]},
        labels=df_f_inner_graph.index,
        values=df_f_inner_graph['졸업요건 현황'],
        textinfo='label',
        textposition='none',
        hovertemplate=iscompleted_list,
        marker={'colors': tf_color_list, 'line': {'color': 'white', 'width': 1}}
    )

    # graph1 - outer graph
    graph1_data2 = go.Pie(
        hole=0.8,
        sort=False,
        direction='clockwise',
        labels=df_f_outer_graph.index,
        values=df_f_outer_graph['이수 학점 분포'],
        textinfo='label',
        textposition=['outside', 'outside', 'outside', 'none'],
        hovertemplate="<b>%{label}</b>: <br>%{value}학점</br><extra></extra>",
        marker={'colors': outer_color_list, 'line': {'color': outer_line_list, 'width': 2}}
    )
    data1 = [graph1_data1, graph1_data2]
    layout1 = go.Layout(title='<b>현재 취득 학점 및 졸업 요건 달성 현황</b>', titlefont=dict(family='맑은 고딕', size=title_font_size),
                        font=dict(family='맑은 고딕', size=17), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), dragmode=False)
    fig1 = go.Figure(data=data1, layout=layout1)
    fig1.update_layout(showlegend=False)
    fig1.update_layout(annotations=[dict(text='현재 취득 학점<br><b>' + str(my_total_credit) + '</b>/130(학점)<br>'
                                                                                         '달성한 졸업 요건<br><b>' + str(
        my_true_count) + '</b>/' + str(len(my_mission_list)) + '(개)</br></br></br>',
                                         x=0.5, y=0.5, font_size=28, font_family='맑은 고딕', showarrow=False)])

    # 학기별 평점 관련 꺾은선 그래프
    data2 = []
    graph2_data1 = go.Scatter(x=df_s_graph.index, y=df_s_graph['학기별 전체 평점'], mode='lines+markers',
                              name='학기별 전체 평점', line=dict(color="#6c6c6c"))
    data2.append(graph2_data1)
    if len(average_major_credit_list) != 0:  # 기초교육학부가 아닐 때
        graph2_data2 = go.Scatter(x=df_s_graph.index, y=df_s_graph['학기별 전공 평점'], mode='lines+markers', name='학기별 전공 평점',
                                  line=dict(color="#de3026"))
        data2.append(graph2_data2)
    layout2 = go.Layout(title='<b>전체 및 전공 평균 평점</b>', titlefont=dict(family='맑은 고딕', size=title_font_size),
                        font=dict(family='맑은 고딕', size=15), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), dragmode=False)
    fig2 = go.Figure(data=data2, layout=layout2)
    fig2.update_traces(mode="markers+lines", hovertemplate=None)
    fig2.update_layout(hovermode="x")
    average_credit = 0
    for credit in average_credit_list:
        average_credit += credit
    average_credit /= len(average_credit_list)
    average_credit = round(average_credit, 2)

    average_major_credit = 0
    if len(average_major_credit_list) != 0:
        for credit in average_major_credit_list:
            if credit is not None:
                average_major_credit += credit
        average_major_credit /= len(list(filter(lambda x: x is not None, average_major_credit_list)))
    average_major_credit = round(average_major_credit, 2)

    fig2.add_trace(go.Scatter(
        x=df_s_graph.index,
        y=[average_credit] * len(df_s_graph),
        name="전체 평점 평균",
        mode="lines",
        line=dict(color="#6c6c6c", width=2, dash='dash'),
        showlegend=False
    ))
    if len(average_major_credit_list) != 0:
        fig2.add_trace(go.Scatter(
            x=df_s_graph.index,
            y=[average_major_credit] * len(df_s_graph),
            name="전공 평점 평균",
            mode="lines",
            line=dict(color="#de3026", width=2, dash='dash'),
            showlegend=False
        ))

    fig2['layout']['yaxis'].update(range=[1, 4.5], dtick=1, autorange=False)
    fig2.update_xaxes(showline=True, linewidth=1.5, linecolor='LightGray')  # mirror=False
    fig2.update_yaxes(showline=True, linewidth=1.5, linecolor='LightGray')  # mirror=False
    fig2.update_xaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')
    fig2.update_yaxes(showgrid=True, gridwidth=1, gridcolor='LightGray')

    # 달성한 졸업요건 관련 표 데이터
    layout3 = go.Layout(title='<b>달성한 졸업 요건</b>', titlefont=dict(family='맑은 고딕', size=title_font_size),
                        font=dict(family='맑은 고딕', size=15), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), dragmode=False)
    rowEvenColor = lighter_color
    rowOddColor = '#FFFFFF'
    rowColorList = []
    for i in range(0, len(df_completed_mission_list)):
        if i % 2 == 0:
            rowColorList.append(rowOddColor)
        else:
            rowColorList.append(rowEvenColor)
    # 글자 색 결정하는 상수(present_color 를 기반으로 함)
    header_luma = 0.2126 * (int(after_source1_sub1[0]) % 256) + \
                  0.7152 * (int(after_source1_sub1[1]) % 256) + \
                  0.0722 * (int(after_source1_sub1[2]) % 256)
    if header_luma >= 127.5:
        header_color = "#333333"
    else:
        header_color = "white"
    fig3 = go.Figure(data=[go.Table(
        columnwidth=[80, 400],
        header=dict(values=['<b>Index</b>', '<b>졸업 요건</b>', '<b>이수 현황</b>'],
                    fill_color=sub1_present_color,
                    align=['center', 'left', 'left'],
                    font=dict(color="white", size=25),
                    height=30),
        cells=dict(values=[df_completed_mission_list['번호'], df_completed_mission_list['졸업 요건'],
                           df_completed_mission_list['이수 현황']],
                   fill_color=[rowColorList],
                   align=['center', 'left', 'left'],
                   font_size=20,
                   height=30))],
        layout=layout3
    )

    # 미달성한 졸업요건 관련 표 데이터
    layout4 = go.Layout(title='<b>미달성한 졸업 요건</b>', titlefont=dict(family='맑은 고딕', size=title_font_size),
                        font=dict(family='맑은 고딕', size=15), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1), dragmode=False)
    rowEvenColor = lighter_color
    rowOddColor = '#FFFFFF'
    rowColorList = []
    for i in range(0, len(df_incompleted_mission_list)):
        if i % 2 == 0:
            rowColorList.append(rowOddColor)
        else:
            rowColorList.append(rowEvenColor)

    fig4 = go.Figure(data=[go.Table(
        columnwidth=[80, 400],
        header=dict(values=['<b>Index</b>', '<b>졸업 요건</b>', '<b>이수 현황</b>'],
                    fill_color=sub1_present_color,
                    align=['center', 'left', 'left'],
                    font=dict(color="white", size=25),
                    height=40),
        cells=dict(values=[df_incompleted_mission_list['번호'], df_incompleted_mission_list['졸업 요건'],
                           df_incompleted_mission_list['이수 현황']],
                   fill_color=[rowColorList],
                   align=['center', 'left', 'left'],
                   font_size=20,
                   height=30))],
        layout=layout4
    )
    '''
    for fig in [fig1, fig2, fig3, fig4]:
        fig.add_layout_image(
            dict(
                source="https://user-images.githubusercontent.com/63055303/103456062-fdb47680-4d35-11eb-92dc-990751fb63cc"
                       ".png",
                xref="paper", yref="paper",
                x=1, y=1.05,
                sizex=0.5, sizey=0.5,
                xanchor="right", yanchor="bottom"
            )
        )
    '''
    info_file = program_directory + '\\edited\\info\\personalinfo2.xlsx'
    df_info = read_excel(info_file, header=None)
    my_number = str(df_info.values[0][1])
    my_name = str(df_info.values[2][1])
    import getpass
    username = getpass.getuser()
    config = {'displaylogo': False, 'displayModeBar': True, 'modeBarButtonsToRemove': ['pan2d', 'lasso2d', 'select2d',
                                                                                       'zoomIn2d', 'zoomOut2d',
                                                                                       'zoom2d', ],
              'toImageButtonOptions': {'format': 'png',  # one of png, svg, jpeg, webp
                                       'filename': 'Credit_Mission_Graph_' + my_number + '_' + my_name,
                                       'height': 650,
                                       'width': 600,
                                       'scale': 1  # Multiply title/legend/axis/canvas sizes by this factor
                                       }
              }
    config2 = {'displaylogo': False, 'displayModeBar': True, 'modeBarButtonsToRemove': ['pan2d', 'lasso2d', 'select2d',
                                                                                        'zoomIn2d', 'zoomOut2d',
                                                                                        'zoom2d', 'toggleSpikelines'],
               'toImageButtonOptions': {'format': 'png',  # one of png, svg, jpeg, webp
                                        'filename': 'Major_and_Entire_Grade_' + my_number + '_' + my_name,
                                        'height': 650,
                                        'width': 1000,
                                        'scale': 1  # Multiply title/legend/axis/canvas sizes by this factor
                                        }
               }
    config3 = {'displaylogo': False, 'displayModeBar': False, 'modeBarButtonsToRemove': ['pan2d', 'lasso2d']}
    info_directory = program_directory + 'edited\\info\\'

    # Fig5 : 부전공(전공과목 제외)
    from numpy import array

    def is_completed(list1, string):
        if list1.count(string) >= 1:
            return True
        else:
            return False

    minor_list1 = ['수학', '의생명공학', '문화기술', '에너지', '지능로봇',
                   '문학', '역사학', '철학',  # 인문학
                   '법학', '정치학', '경제학', '경영학', '사회학', '심리학', '과학기술학']  # 사회과학

    df_grade_code = list(array(df_subject['교과목'].tolist()))
    # 1-1. 수학
    nec_math_token = False
    nec_math_count1 = 0  # nec_math_token1 해당 과목 개수
    nec_math_token1 = False  # 다변수 선대 미분 현대 중 택3
    nec_math_token2 = False  # 해석학 또는 복소함수학
    sel_math_token = False  # 수학 선택과목
    sel_math_credit = 0  # 수학 선택과목 학점
    math_token = False  # 수학 과목
    math_list = ["MM2006", "MM2007", "MM3012", "MM3015",  # 수학 선택과목 리스트
                 "MM4003", "MM4005", "MM4006", "MM4007",
                 "MM4008", "MM4009", "MM4010", "MM4015",
                 "MM4016", "MM4017", "MM4018", "MM4019",
                 "GS2007", "GS3012", "GS4005", "GS4008",
                 "GS3015", "GS4006", "GS4007", "GS4015",
                 "GS4016", "GS4017", "GS4018", "GS4019"]

    if is_completed(df_grade_code, "GS2003") or is_completed(df_grade_code, "GS2013"):  # 선대미분방정식 수강 여부
        nec_math_count1 += 1
        # 1: 선대미분방정식(3) + 다변수(3) + 현대대수학(3) 조합      (필수교과2 3학점 + 선택교과 6학점)
        if is_completed(df_grade_code, "GS1002") or is_completed(df_grade_code, "GS1012") or \
                is_completed(df_grade_code, "GS2001") or is_completed(df_grade_code, "GS2011") or \
                is_completed(df_grade_code, "MM2001") or is_completed(df_grade_code, "MM2011"):
            nec_math_count1 += 1

        if is_completed(df_grade_code, "GS3004") or is_completed(df_grade_code, "GS4004") or \
                is_completed(df_grade_code, "MM4004"):
            nec_math_count1 += 1

        if nec_math_count1 >= 3:
            nec_math_token1 = True

        if is_completed(df_grade_code, "MM3001") or is_completed(df_grade_code, "MM4002") or \
                is_completed(df_grade_code, "GS3001") or is_completed(df_grade_code, "GS3002") or \
                is_completed(df_grade_code, "GS4002"):
            nec_math_token2 = True
        if len(math_list) != 0:
            for code in math_list:
                if is_completed(df_grade_code, code):
                    if code == "MM3012":  # 수학적 모델링(2학점)
                        sel_math_credit += 2
                    else:
                        sel_math_credit += 3
        if sel_math_credit >= 6:
            sel_math_token = True
    else:
        # 2: 선대(3)+미분방정식(3)+다변수(3)+현대대수학(3) 조합(필수교과2 3학점 + 선택교과 3학점)
        if is_completed(df_grade_code, "GS1002") or is_completed(df_grade_code, "GS1012") or \
                is_completed(df_grade_code, "GS2001") or is_completed(df_grade_code, "GS2011") or \
                is_completed(df_grade_code, "MM2001") or is_completed(df_grade_code, "MM2011"):
            nec_math_count1 += 1

        if is_completed(df_grade_code, "GS2004") or is_completed(df_grade_code, "MM2004"):
            nec_math_count1 += 1

        if is_completed(df_grade_code, "GS2002") or is_completed(df_grade_code, "MM2002"):
            nec_math_count1 += 1

        if is_completed(df_grade_code, "GS3004") or is_completed(df_grade_code, "GS4004") or \
                is_completed(df_grade_code, "MM4004"):
            nec_math_count1 += 1

        if nec_math_count1 >= 4:
            nec_math_token1 = True

        if is_completed(df_grade_code, "MM3001") or is_completed(df_grade_code, "MM4002") or \
                is_completed(df_grade_code, "GS3001") or is_completed(df_grade_code, "GS3002") or \
                is_completed(df_grade_code, "GS4002"):
            nec_math_token2 = True
        if len(math_list) != 0:
            for code in math_list:
                if is_completed(df_grade_code, code):
                    if code == "MM3012":  # 수학적 모델링(2학점)
                        sel_math_credit += 2
                    else:
                        sel_math_credit += 3
        if sel_math_credit >= 3:
            sel_math_token = True
    if nec_math_token1 and nec_math_token2:
        nec_math_token = True
    if nec_math_token and sel_math_token:
        math_token = True

    # 1-2. 의생명공학
    medic_credit = 0  # 의생명공학 학점
    medic_credit_token = False
    medic_token = False
    medic_list = ["MD2101", "MD4101", "MD4102", "MD4301",  # 의생명공학 과목
                  "MD4302", "MD4303", "MD4501", "MD4502",
                  "MD4601", "GS4801"]
    if len(df_grade_code) != 0:
        for code in medic_list:
            if is_completed(df_grade_code, code):
                if code == "MD2101":  # 의공학 입문(2)
                    medic_credit += 2
                else:
                    medic_credit += 3
        if medic_credit >= 15:
            medic_credit_token = True
    if medic_credit_token:
        medic_token = True

    # 1-3. 문화기술
    nec_culture_token = False
    nec_culture_count = 0  # 문화기술 필수 개수(3개 이상)
    nec_culture_list = ["CT4101", "CT4201", "CT4301", "CT4302", "EC4215"]
    culture_credit_token = False
    culture_credit = 0  # 문화기술 선택 학점
    sel_culture_list = ["CT2501", "CT2502", "CT2503", "CT2504",
                        "CT2505", "CT2506", "CT4504", "CT4506",
                        "GS2707", "GS2543", "GS2501", "GS2814",
                        "GS4005", "GS4008"]
    culture_token = False
    if len(df_grade_code) != 0:
        # 문화기술 필수 과목 개수 및 학점 계산
        for code in nec_culture_list:
            if is_completed(df_grade_code, code):
                nec_culture_count += 1
                culture_credit += 3
        # 문화기술 선택 과목 학점 계산
        for code in sel_culture_list:
            if is_completed(df_grade_code, code):
                culture_credit += 3
        # 문화기술 과목 토큰 관련 식
        if nec_culture_count >= 3:
            nec_culture_token = True
        if culture_credit >= 15:
            culture_credit_token = True
        if nec_culture_token and culture_credit_token:
            culture_token = True

    # 1-4. 에너지
    energy_credit = 0  # 문화기술 선택 학점
    energy_credit_token = False
    sel_energy_list = ["ET2101", "ET4102", "ET4201", "ET4302",
                       "ET4303", "ET4304", "ET4306", "ET4501",
                       "GS2820"]
    energy_token = False

    if len(df_grade_code) != 0:
        # 에너지 선택 과목 학점 계산
        for code in sel_energy_list:
            if is_completed(df_grade_code, code):
                if code == "ET2101":  # 에너지와 미래사회(1)
                    energy_credit += 1
                else:
                    energy_credit += 3
        if energy_credit >= 15:
            energy_credit_token = True
        if energy_credit_token:
            energy_token = True

    # 1-5. 지능로봇
    nec_irobot_token = False
    nec_irobot_count = 0  # 지능로봇 필수 개수(3개 이상)
    nec_irobot_list = ["IR4201", "IR4202", "IR4203", "IR4204",
                       "IR4205", "IR4206", "GS4762"]
    irobot_credit_token = False
    irobot_credit = 0  # 지능로봇 학점
    sel_irobot_list = ["IR2201", "IR2202", "IR3202", "IR3203",
                       "IR4207", "IR4208", "IR4209", "GS2401",
                       "EC2201", "MC2103", "MC3203", "EC3214",
                       "MC4216"]
    irobot_token = False

    if len(df_grade_code) != 0:
        # 문화기술 필수 과목 개수 및 학점 계산
        for code in nec_irobot_list:
            if is_completed(df_grade_code, code):
                nec_irobot_count += 1
                irobot_credit += 3
        # 문화기술 선택 과목 학점 계산
        for code in sel_irobot_list:
            if is_completed(df_grade_code, code):
                irobot_credit += 3
        # 문화기술 과목 토큰 관련 식
        if nec_irobot_count >= 3:
            nec_irobot_token = True
        if irobot_credit >= 15:
            irobot_credit_token = True
        if nec_irobot_token and irobot_credit_token:
            irobot_token = True

    # 2-1. 문학
    nec_liter_token = False
    nec_liter_count = 0  # 문학 필수 개수(1개 이상)
    nec_liter_list = ["GS3504", "GS2507", "GS2521", "GS2526"]
    liter_credit_token = False
    liter_credit = 0  # 문학 학점
    sel_liter_list = ["GS2501", "GS2503", "GS2506", "GS2509",
                      "GS2511", "GS2512", "GS2522", "GS2523",
                      "GS2524", "GS2525", "GS2814", "GS3501",
                      "GS3502", "GS3505", "GS3802", "GS3803"]
    sel_liter_list_u17 = ["GS2502", "GS2505", "GS2510"]  # 17학번까지만 인정하는 과목
    liter_token = False

    if len(df_grade_code) != 0:
        # 문학 필수 과목 개수 및 학점 계산
        for code in nec_liter_list:
            if is_completed(df_grade_code, code):
                nec_liter_count += 1
                liter_credit += 3
        # 문학 선택 과목 학점 계산
        for code in sel_liter_list:
            if is_completed(df_grade_code, code):
                liter_credit += 3
        if int(my_number[:4]) <= 2017:  # 17학번 이전일 시
            for code in sel_liter_list_u17:
                if is_completed(df_grade_code, code):
                    liter_credit += 3
        # 문학 과목 토큰 관련 식
        if nec_liter_count >= 1:
            nec_liter_token = True
        if liter_credit >= 15:
            liter_credit_token = True
        if nec_liter_token and liter_credit_token:
            liter_token = True

    # 2-2. 역사학
    nec_history_token = False
    nec_history_count = 0  # 역사학 필수 개수(1개 이상)
    nec_history_list = ["GS2602"]
    history_credit_token = False
    history_credit = 0  # 역사학 학점
    sel_history_list = ["GS2601", "GS2603", "GS3901", "GS2612",
                        "GS2612", "GS2613", "GS2614", "GS2615",
                        "GS2616", "GS2656", "GS2618", "GS3601"]
    history_token = False

    if len(df_grade_code) != 0:
        # 역사학 필수 과목 개수 및 학점 계산
        for code in nec_history_list:
            if is_completed(df_grade_code, code):
                nec_history_count += 1
                history_credit += 3
        # 역사학 선택 과목 학점 계산
        for code in sel_history_list:
            if is_completed(df_grade_code, code):
                history_credit += 3
        # 역사학 과목 토큰 관련 식
        if nec_history_count >= 1:
            nec_history_token = True
        if history_credit >= 15:
            history_credit_token = True
        if nec_history_token and history_credit_token:
            history_token = True

    # 2-3. 철학
    nec_philos_token = False
    nec_philos_count = 0  # 철학 필수 개수(1개 이상)
    nec_philos_list = ["GS2620"]
    philos_credit_token = False
    philos_credit = 0  # 철학 학점
    sel_philos_list = ["GS2621", "GS2622", "GS2625", "GS2626",
                       "GS2627", "GS3621", "GS3622", "GS3628",
                       "GS2623", "GS3623", "GS3624", "GS2629",
                       "GS3625", "GS2630", "GS3626", "GS3662",
                       "GS3663", "GS2661", "GS3631", "GS3632",
                       "GS3633", "CT2506"]
    philos_token = False

    if len(df_grade_code) != 0:
        # 철학 필수 과목 개수 및 학점 계산
        for code in nec_philos_list:
            if is_completed(df_grade_code, code):
                nec_philos_count += 1
                philos_credit += 3
        # 철학 선택 과목 학점 계산
        for code in sel_philos_list:
            if is_completed(df_grade_code, code):
                philos_credit += 3
        # 철학 과목 토큰 관련 식
        if nec_philos_count >= 1:
            nec_philos_token = True
        if philos_credit >= 15:
            philos_credit_token = True
        if nec_philos_token and philos_credit_token:
            philos_token = True

    # 3-1. 법학
    nec_law_token = False
    nec_law_count = 0  # 법학 필수 개수(1개 이상)
    nec_law_list = ["GS2763"]
    law_credit_token = False
    law_credit = 0  # 법학 학점
    sel_law_list = ["GS2761", "GS2762", "GS2765", "GS2812",
                    "GS2812", "GS3762", "GS3763", "GS3861",
                    "GS4761", "GS3761", "GS4762"]
    law_token = False

    if len(df_grade_code) != 0:
        # 법학 필수 과목 개수 및 학점 계산
        for code in nec_law_list:
            if is_completed(df_grade_code, code):
                nec_law_count += 1
                law_credit += 3
        # 법학 선택 과목 학점 계산
        for code in sel_law_list:
            if is_completed(df_grade_code, code):
                law_credit += 3
        # 법학 과목 토큰 관련 식
        if nec_law_count >= 1:
            nec_law_token = True
        if law_credit >= 15:
            law_credit_token = True
        if nec_law_token and law_credit_token:
            law_token = True

    # 3-2. 정치학
    nec_poli_token = False
    nec_poli_count = 0  # 정치학 필수 개수(1개 이상)
    nec_poli_list = ["GS2787"]
    poli_credit_token = False
    poli_credit = 0  # 정치학 학점
    sel_poli_list = ["GS2785", "GS2786", "GS2788", "GS2781",
                     "GS2782", "GS2783", "GS2784", "GS2709"]
    poli_token = False

    if len(df_grade_code) != 0:
        # 정치학 필수 과목 개수 및 학점 계산
        for code in nec_poli_list:
            if is_completed(df_grade_code, code):
                nec_poli_count += 1
                poli_credit += 3
        # 정치학 선택 과목 학점 계산
        for code in sel_poli_list:
            if is_completed(df_grade_code, code):
                poli_credit += 3
        # 정치학 과목 토큰 관련 식
        if nec_poli_count >= 1:
            nec_poli_token = True
        if poli_credit >= 15:
            poli_credit_token = True
        if nec_poli_token and poli_credit_token:
            poli_token = True

    # 3-3. 경제학(econo: 과학기술과 경제 토큰(economy_token 과 혼동 우려))
    nec_econo_token = False
    nec_econo_count = 0  # 경제학 필수 개수(2개 이상)
    nec_econo_list = ["GS2724", "GS2731"]
    econo_credit_token = False
    econo_credit = 0  # 경제학 학점
    sel_econo_list = ["GS2725", "GS2721", "GS2726", "GS2727",
                      "GS2728", "GS2729", "GS2730", "GS2732",
                      "GS2733", "GS3721", "GS3765", "GS2736"]
    econo_token = False

    if len(df_grade_code) != 0:
        # 경제학 필수 과목 개수 및 학점 계산
        for code in nec_econo_list:
            if is_completed(df_grade_code, code):
                nec_econo_count += 1
                econo_credit += 3
        # 경제학 선택 과목 학점 계산
        for code in sel_econo_list:
            if is_completed(df_grade_code, code):
                econo_credit += 3
        # 경제학 과목 토큰 관련 식
        if nec_econo_count >= 2:
            nec_econo_token = True
        if econo_credit >= 15:
            econo_credit_token = True
        if nec_econo_token and econo_credit_token:
            econo_token = True

    # 3-4. 경영학
    nec_business_token = False
    nec_business_count = 0  # 경영학 필수 개수(1개 이상)
    nec_business_list = ["GS2750"]
    business_credit_token = False
    business_credit = 0  # 경영학 학점
    sel_business_list = ["GS2704", "GS2751", "GS3752", "GS2731",
                         "GS2752", "GS3751", "GS3752", "GS3753"]
    business_token = False

    if len(df_grade_code) != 0:
        # 경영학 필수 과목 개수 및 학점 계산
        for code in nec_business_list:
            if is_completed(df_grade_code, code):
                nec_business_count += 1
                business_credit += 3
        # 경영학 선택 과목 학점 계산
        for code in sel_business_list:
            if is_completed(df_grade_code, code):
                business_credit += 3
        # 경영학 과목 토큰 관련 식
        if nec_business_count >= 1:
            nec_business_token = True
        if business_credit >= 15:
            business_credit_token = True
        if nec_business_token and business_credit_token:
            business_token = True

    # 3-5. 사회학
    nec_social_token = False
    nec_social_count = 0  # 사회학 필수 개수(1개 이상)
    nec_social_list = ["GS2704"]
    social_credit_token = False
    social_credit = 0  # 사회학 학점
    sel_social_list = ["GS2701", "GS2702", "GS2703", "GS2705",
                       "GS2706", "GS2707", "GS2708", "GS2785",
                       "GS2786", "GS2803", "GS2831", "GS3751"]
    social_token = False

    if len(df_grade_code) != 0:
        # 사회학 필수 과목 개수 및 학점 계산
        for code in nec_social_list:
            if is_completed(df_grade_code, code):
                nec_social_count += 1
                social_credit += 3
        # 사회학 선택 과목 학점 계산
        for code in sel_social_list:
            if is_completed(df_grade_code, code):
                social_credit += 3
        # 사회학 과목 토큰 관련 식
        if nec_social_count >= 1:
            nec_social_token = True
        if social_credit >= 15:
            social_credit_token = True
        if nec_social_token and social_credit_token:
            social_token = True

    # 3-6. 심리학
    nec_psycho_token = False
    nec_psycho_count = 0  # 심리학 필수 개수(2개 이상)
    nec_psycho_list = ["GS2742", "GS2743"]
    psycho_credit_token = False
    psycho_credit = 0  # 심리학 학점
    sel_psycho_list = ["GS2747", "GS2748", "GS4741", "GS3764"]
    psycho_token = False

    if len(df_grade_code) != 0:
        # 심리학 필수 과목 개수 및 학점 계산
        for code in nec_psycho_list:
            if is_completed(df_grade_code, code):
                nec_psycho_count += 1
                psycho_credit += 3
        # 심리학 선택 과목 학점 계산
        for code in sel_psycho_list:
            if is_completed(df_grade_code, code):
                psycho_credit += 3
        # 심리학 과목 토큰 관련 식
        if nec_psycho_count >= 2:
            nec_psycho_token = True
        if psycho_credit >= 15:
            psycho_credit_token = True
        if nec_psycho_token and psycho_credit_token:
            psycho_token = True

    # 3-7. 과학기술학
    nec_sts_token = False
    nec_sts_count = 0  # 과학기술학 필수 개수(1개 이상)
    nec_sts_list = ["GS2831"]
    sts_credit_token = False
    sts_credit = 0  # 과학기술학 학점
    sel_sts_list = ["GS2803", "GS2832", "GS2833", "GS2834",
                    "GS2837", "GS3801", "GS4761", "GS3761"]
    sts_token = False

    if len(df_grade_code) != 0:
        # 과학기술학 필수 과목 개수 및 학점 계산
        for code in nec_sts_list:
            if is_completed(df_grade_code, code):
                nec_sts_count += 1
                sts_credit += 3
        # 과학기술학 선택 과목 학점 계산
        for code in sel_sts_list:
            if is_completed(df_grade_code, code):
                sts_credit += 3
        # 과학기술학 과목 토큰 관련 식
        if nec_sts_count >= 1:
            nec_sts_token = True
        if sts_credit >= 15:
            sts_credit_token = True
        if nec_sts_token and sts_credit_token:
            sts_token = True

    # 부전공(주전공 이외) 관련 표 데이터
    layout5 = go.Layout(title='<b>분야별 부전공 달성 현황<br>(7개 주전공 이외, 인문학 및 사회과학 분야: 단일분야 기준)</b>',
                        titlefont=dict(family='맑은 고딕', size=title_font_size),
                        font=dict(family='맑은 고딕', size=15), paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                        dragmode=False)
    # present_color 기반 밝은 색 제작
    lighter_red2 = int(after_source1_sub2[0]) % 256 + (255 - int(after_source1_sub2[0]) % 256) * 4 / 5
    lighter_green2 = int(after_source1_sub2[1]) % 256 + (255 - int(after_source1_sub2[1]) % 256) * 4 / 5
    lighter_blue2 = int(after_source1_sub2[2]) % 256 + (255 - int(after_source1_sub2[2]) % 256) * 4 / 5
    lighter_color2 = 'rgb(' + str(lighter_red2) + ',' + str(int(lighter_green2)) + ',' + str(lighter_blue2) + ')'

    fig5_cell_color_list = []
    true_value = True  # 항상 참 값을 의미하는 변수
    minor1_tf_list = (math_token, math_token, nec_math_token, sel_math_token), \
                     (medic_token, medic_token, true_value, medic_credit_token), \
                     (culture_token, culture_token, nec_culture_token, culture_credit_token), \
                     (energy_token, energy_token, true_value, energy_credit_token), \
                     (irobot_token, irobot_token, nec_irobot_token, irobot_credit_token), \
                     (liter_token, liter_token, nec_liter_token, liter_credit_token), \
                     (history_token, history_token, nec_history_token, history_credit_token), \
                     (philos_token, philos_token, nec_philos_token, philos_credit_token), \
                     (law_token, law_token, nec_law_token, law_credit_token), \
                     (poli_token, poli_token, nec_poli_token, poli_credit_token), \
                     (econo_token, econo_token, nec_econo_token, econo_credit_token), \
                     (business_token, business_token, nec_business_token, business_credit_token), \
                     (social_token, social_token, nec_social_token, social_credit_token), \
                     (psycho_token, psycho_token, nec_psycho_token, psycho_credit_token), \
                     (sts_token, sts_token, nec_sts_token, sts_credit_token)
    minor1_credit_list = (sel_math_credit, medic_credit, culture_credit, energy_credit,
                          irobot_credit, liter_credit, history_credit, philos_credit,
                          law_credit, poli_credit, econo_credit, business_credit,
                          social_credit, psycho_credit, sts_credit)
    for num in range(0, 4):
        row_color_list = []
        for minor in minor1_tf_list:
            if minor[num] or num == 0:  # 첫 번째 열(Index)과 조건 만족 셀에는 색칠
                row_color_list.append(lighter_color2)
            else:
                row_color_list.append("#FFFFFF")
        fig5_cell_color_list.append(row_color_list)

    # fig5 구성 데이터 리스트
    fig5_index_data = []  # index
    fig5_minor_name = minor_list1  # 부전공 교과목명
    fig5_nec_data = []  # 필수과목 이수 현황
    fig5_credit_data = []  # 취득 학점 현황

    for i in range(0, len(minor1_tf_list)):
        fig5_index_data.append(str(i + 1))

    if nec_math_token1:
        temp_math_tf1 = "이수"  # 필수교과 I 이수여부(이수, 미이수)
    else:
        temp_math_tf1 = "미이수"
    if nec_math_token2:
        temp_math_tf2 = "이수"  # 필수교과 II 이수여부(이수, 미이수)
    else:
        temp_math_tf2 = "미이수"
    fig5_nec_data.append("필수 교과 Ⅰ(다변수+선대+미방+현대): " + temp_math_tf1 + "<br>"
                                                                     "필수 교과 Ⅱ(해석학, 복소함수학 중 택1): " + temp_math_tf2)
    fig5_index_data[0] += "<br>"
    fig5_minor_name[0] += "<br>"
    # 수학 (선택)교과 학점 append 하기
    if is_completed(df_grade_code, "GS2003") or is_completed(df_grade_code, "GS2013"):
        fig5_credit_data.append("선택 교과 취득 학점: " + str(sel_math_credit) + "/6(학점)<br>")
    else:
        fig5_credit_data.append("선택 교과 취득 학점: " + str(sel_math_credit) + "/3(학점)<br>")
    # 수학 제외 부전공 과목 필수 이수 여부 및 총 학점 append 하기
    for row_num in range(1, len(minor1_tf_list)):
        if minor1_tf_list[row_num][2]:
            if row_num == 1 or row_num == 3:  # 수강 필수 과목 없는 부전공
                fig5_nec_data.append("-")
            else:
                fig5_nec_data.append("이수")
        else:
            fig5_nec_data.append("미이수")
        fig5_credit_data.append("총 취득 학점: " + str(minor1_credit_list[row_num]) + "/15(학점)")

    fig5_nec_advice = ["", "", " - CT4101(4201, 4301, 4302) 중 택3", "",
                       " - IR420x(1~6) 중 택3", " - GS3504(2507, 2521, 2526) 중 택1", ' - "한국사의 이해" 이수 필요',
                       ' - "철학의 근본 문제들" 이수 필요', ' - "현대법학의 이해" 이수 필요', ' - "한국정치론" 이수 필요',
                       ' - "거시경제학", "미시경제학" 이수 필요(2과목)', ' - "경영학원론" 이수 필요', ' - "기업과 사회(Ⅰ)" 이수 필요',
                       ' - "인간과 마음의 행동 1 및 2" 이수 필요(2과목)', ' - "과학기술학의 이해 1" 이수 필요']
    for nec_data_num in range(0, len(fig5_nec_data)):
        if fig5_nec_data[nec_data_num] == "미이수":
            fig5_nec_data[nec_data_num] += fig5_nec_advice[nec_data_num]

    # 글자 색 결정하는 상수(sub2_present_color 를 기반으로 함)
    fig5_header_luma = 0.2126 * (int(after_source1_sub2[0]) % 256) + \
                       0.7152 * (int(after_source1_sub2[1]) % 256) + \
                       0.0722 * (int(after_source1_sub2[2]) % 256)
    if fig5_header_luma >= 127.5:
        fig5_header_color = "#333333"
    else:
        fig5_header_color = "white"
    fig5 = go.Figure(data=[go.Table(
        columnwidth=[80, 100, 300],
        header=dict(values=['<b>Index</b>', '<b>부전공</b>', '<b>필수 교과목 이수 여부</b>', '<b>취득 학점</b>'],
                    fill_color=sub2_present_color,
                    align=['center', 'center', 'left', 'left'],
                    font=dict(color=fig5_header_color, size=25),
                    height=40),
        cells=dict(values=[fig5_index_data, fig5_minor_name,
                           fig5_nec_data, fig5_credit_data],
                   fill_color=fig5_cell_color_list,
                   align=['center', 'center', 'left', 'left'],
                   font_size=20,
                   height=30))],
        layout=layout5
    )

    # Fig6 : 부전공 과목(선언 교과 관련)
    minor_list2 = ['물리', '화학', '생명과학', '전기전자컴퓨터', '기계공학', '신소재공학', '지구·환경공학']

    # 1: 물리
    nec_min_physics_token = False
    nec_min_physics_count = 0  # 물리 필수 개수(3개 이상, 3-4천번대 과목)
    nec_min_physics_list = ['PS3204', 'PS3101', 'PS3103', 'PS3104',
                            'PS3105', 'PS3106', 'PS3107']
    min_physics_credit_token = False
    min_physics_credit = 0  # 물리 학점
    my_min_physics_code = list(filter(lambda x: x.startswith('PS3') or x.startswith('PS4'), df_grade_code))
    min_physics_token = False

    if len(my_min_physics_code) != 0:
        # 물리 필수 과목 개수 계산
        for code in nec_min_physics_list:
            if is_completed(my_min_physics_code, code):
                nec_min_physics_count += 1
        # 물리 과목 학점 계산
        for code_num in range(0, len(my_min_physics_code)):
            min_physics_credit += 3
        # 물리 과목 토큰 관련 식
        if nec_min_physics_count >= 3:
            nec_min_physics_token = True
        if min_physics_credit >= 15:
            min_physics_credit_token = True
        if nec_min_physics_token and min_physics_credit_token:
            min_physics_token = True

    # 2: 화학
    nec_min_chemi_token = False
    nec_min_chemi_count = 0  # 화학 필수 개수(3개 이상, 3-4천번대 과목)
    nec_min_chemi_list_u17 = ['CH2101', 'CH3204', 'CM3204', 'CM501Y',
                              'CH3104', 'CM3104', 'CH2104', 'CH2105',
                              'CH3102', 'CM3102', 'CH3106', 'CM3203',
                              'CH3107', 'CH4218', 'CM3101', 'CM4218']
    nec_min_chemi_list_s18 = list(filter(lambda x: x.startswith('CH3') or x.startswith('CH4') or x.startswith('CM'),
                                         nec_min_chemi_list_u17))
    min_chemi_credit_token = False
    min_chemi_credit = 0  # 화학 학점
    # ~17학번: 유기화학 I 이외 모든 화학 과목 추출
    my_min_chemi_code_u17 = list(filter(lambda x: x.startswith('CH') or x.startswith('CM') and x != "CH2103",
                                        df_grade_code))
    # 18학번~: 3-4번대 과목 추출
    my_min_chemi_code_s18 = list(filter(lambda x: x.startswith('CH3') or x.startswith('CH4') or x.startswith('CM'),
                                        df_grade_code))
    min_chemi_token = False

    if len(my_min_chemi_code_u17 or my_min_chemi_code_s18) != 0:
        if int(my_number[:4]) <= 2017:  # 17학번 이전일 시
            for code in nec_min_chemi_list_u17:
                # 화학 필수 과목 개수 계산
                if is_completed(my_min_chemi_code_u17, code):
                    nec_min_chemi_count += 1
            # 화학 과목 학점 계산
            for code_num in range(0, len(my_min_chemi_code_u17)):
                min_chemi_credit += 3
            # 화학 과목 학점 토큰 계산
            if min_chemi_credit >= 15:
                min_chemi_credit_token = True
        else:  # 18학번 이후일 시
            for code in nec_min_chemi_list_s18:
                # 화학 필수 과목 개수 계산
                if is_completed(my_min_chemi_code_s18, code):
                    nec_min_chemi_count += 1
            # 화학 과목 학점 계산
            for code_num in range(0, len(my_min_chemi_code_s18)):
                min_chemi_credit += 3
            # 화학 과목 학점 토큰 계산
            if min_chemi_credit >= 21:
                min_chemi_credit_token = True
        # 화학 과목 토큰 관련 식
        if nec_min_chemi_count >= 3:
            nec_min_chemi_token = True
        if nec_min_chemi_token and min_chemi_credit_token:
            min_chemi_token = True

    # 3. 생명과학
    nec_min_bio_token = False
    nec_min_bio_count = 0  # 생명과학 필수 전공 개수(전공 2개 이상, 3-4천번대 과목)
    nec_min_bio_exp_count = 0  # 생명과학 실험 필수 교과 개수(실험 1개 이상)
    nec_min_bio_list = ['BS3101', 'BS3105', 'BS3113']
    nec_min_bio_exp_list = ['BS3111', 'BS3112']
    min_bio_credit_token = False
    min_bio_credit = 0  # 물리 학점
    my_min_bio_code = list(filter(lambda x: x.startswith('BS3') or x.startswith('BS4'), df_grade_code))
    min_bio_token = False

    if len(my_min_bio_code) != 0:
        # 생명과학 필수 과목 개수 계산
        for code in nec_min_bio_list:
            if is_completed(my_min_bio_code, code):
                nec_min_bio_count += 1
        for code in nec_min_bio_exp_list:
            if is_completed(my_min_bio_code, code):
                nec_min_bio_exp_count += 1
        # 생명과학 과목 학점 계산
        for code_num in range(0, len(my_min_bio_code)):
            min_bio_credit += 3
        # 생명과학 과목 토큰 관련 식
        if nec_min_bio_count >= 2 and nec_min_bio_exp_count >= 1:
            nec_min_bio_token = True
        if min_bio_credit >= 15:
            min_bio_credit_token = True
        if nec_min_bio_token and min_bio_credit_token:
            min_bio_token = True

    # 4. 전기전자컴퓨터
    nec_min_eecs_token = False
    nec_min_eecs_list = ['EC3101', 'EC3102']
    min_eecs_credit_token = False
    min_eecs_credit3 = 0  # 전컴 학점(3천번대 이상)
    min_eecs_credit2 = 0  # 전컴 학점(2천번대)
    # ~17학번: 3-4번대 과목 추출
    my_min_eecs_code_u17 = list(filter(lambda x: x.startswith('EC3') or x.startswith('EC4')
                                       or x == "EC2206" or x == "EC2204" or x == "EC2205"
                                       or x == "MM4010"
                                       or x == "CT4201",  # 이산수학, 그래픽스
                                       df_grade_code))
    # 18학번~: 모든 전컴 과목 추출
    my_min_eecs_code_s18 = list(filter(lambda x: x.startswith('EC')
                                       or x == "MM4010"
                                       or x == "CT4201",  # 이산수학, 그래픽스
                                       df_grade_code))
    min_eecs_token = False

    if len(my_min_eecs_code_u17 or my_min_eecs_code_s18) != 0:
        for code in nec_min_eecs_list:
            # 전컴 필수 과목 개수 및 학점 계산
            if is_completed(my_min_eecs_code_u17, code):
                nec_min_eecs_token = True
        if int(my_number[:4]) <= 2017:  # 17학번 이전일 시
            # 전컴 과목 학점 계산
            for code_num in range(0, len(my_min_eecs_code_u17)):
                if my_min_eecs_code_u17[code_num] == nec_min_eecs_list[1]:  # 컴시이실
                    min_eecs_credit3 += 4
                else:
                    min_eecs_credit3 += 3
            # 전컴 과목 학점 토큰 계산
            if min_eecs_credit3 >= 15:
                min_eecs_credit_token = True
        else:  # 18학번 이후일 시
            # 전컴 과목 학점 계산
            for code_num in range(0, len(my_min_eecs_code_s18)):
                if my_min_eecs_code_s18[code_num][:3] == "EC2":  # 2천번대 과목일 시
                    min_eecs_credit2 += 3
                else:  # 3천번대 이상 과목일 시
                    if my_min_eecs_code_s18[code_num] == nec_min_eecs_list[1]:  # 컴시이실
                        min_eecs_credit3 += 4
                    else:
                        min_eecs_credit3 += 3
            # 전컴 과목 학점 토큰 계산
            if min_eecs_credit2 >= 6 and min_eecs_credit3 >= 12:
                min_eecs_credit_token = True
        if nec_min_eecs_token and min_eecs_credit_token:
            min_eecs_token = True

    # 5. 기계공학
    nec_min_mecha_token = False
    nec_min_mecha_count = 0  # 기계공학 필수 개수(3개 이상)
    nec_min_mecha_list = ['MC2100', 'MC2101', 'MC3102', 'MC2102', 'MC3105', 'MC2103', 'MC3212']
    min_mecha_credit_token = False
    min_mecha_credit = 0  # 기계공학 학점
    my_min_mecha_code = list(filter(lambda x: x.startswith('MC3') or x.startswith('MC4') or \
                                              x.startswith('MC210'), df_grade_code))
    min_mecha_token = False

    if len(my_min_mecha_code) != 0:
        # 기계공학 필수 과목 개수 계산
        for code in nec_min_mecha_list:
            if is_completed(my_min_mecha_code, code):
                nec_min_mecha_count += 1
        # 기계공학 과목 학점 계산
        for code_num in range(0, len(my_min_mecha_code)):
            min_mecha_credit += 3
        # 기계공학 과목 토큰 관련 식
        if nec_min_mecha_count >= 3:
            nec_min_mecha_token = True
        if min_mecha_credit >= 15:
            min_mecha_credit_token = True
        if nec_min_mecha_token and min_mecha_credit_token:
            min_mecha_token = True

    # 6. 신소재공학
    nec_min_mater_token = False
    nec_min_mater_count = 0  # 신소재 필수 전공 개수(~17학번)
    nec_min_mater_credit = 0  # 신소재 필수 전공 학점(18학번~)
    nec_min_mater_list = ['MA2101', 'MA3101', 'MA2102', 'MA3205', 'MA2103',
                          'MA2104', 'MA3102', 'MA3104', 'MA3105']
    min_mater_credit_token = False
    min_mater_credit = 0  # 신소재 학점(3천번대 이상)
    # ~17학번: 모든 신소재 과목 추출
    my_min_mater_code_u17 = list(filter(lambda x: x.startswith('MA3') or x.startswith('MA4') or x.startswith('MA210'),
                                        df_grade_code))
    # 18학번~: 모든 신소재 과목 추출
    my_min_mater_code_s18 = list(filter(lambda x: x.startswith('MA'), df_grade_code))
    min_mater_token = False

    if len(my_min_mater_code_u17 or my_min_mater_code_s18) != 0:
        if int(my_number[:4]) <= 2017:  # 17학번 이전일 시
            for code in nec_min_mater_list:
                # 신소재 필수 과목 개수 및 학점 계산
                if is_completed(my_min_mater_code_u17, code):
                    nec_min_mater_count += 1
            # 신소재 과목 학점 계산
            for code_num in range(0, len(my_min_mater_code_u17)):
                min_mater_credit += 3
            # 신소재 과목 학점 토큰 계산
            if nec_min_mater_count >= 3:
                nec_min_mater_token = True
            if min_mater_credit >= 15:
                min_mater_credit_token = True
        else:  # 18학번 이후일 시
            for code in nec_min_mater_list:
                # 신소재 필수 과목 개수 및 학점 계산
                if is_completed(my_min_mater_code_s18, code):
                    nec_min_mater_credit += 3
            # 신소재 과목 학점 계산
            for code_num in range(0, len(my_min_mater_code_s18)):
                if my_min_mater_code_s18[code_num][:3] == "MA2":  # 2천번대 과목일 시
                    continue
                else:  # 3천번대 이상 과목일 시
                    min_mater_credit += 3
            # 신소재 과목 학점 토큰 계산
            if nec_min_mater_credit >= 6:
                nec_min_mater_token = True
            if min_mater_credit >= 9:
                min_mater_credit_token = True
        if nec_min_mater_token and min_mater_credit_token:
            min_mater_token = True

    # 7. 지구·환경공학
    nec_min_enviro_token = False
    nec_min_enviro_count = 0  # 지환공 필수 개수(3개 이상, 3-4천번대 과목)
    nec_min_enviro_list = ['EV3101', 'EV3111', 'EV3103', 'EV4106']
    min_enviro_engine_token = False
    min_enviro_credit_token = False
    min_enviro_credit = 0  # 지환공 학점
    my_min_enviro_code = list(filter(lambda x: x.startswith('EV3') or x.startswith('EV4'), df_grade_code))
    min_enviro_token = False

    if len(my_min_enviro_code) != 0:
        # 지환공 필수 과목 개수 계산
        for code in nec_min_enviro_list:
            if is_completed(my_min_enviro_code, code):
                nec_min_enviro_count += 1
        # 지환공 과목 학점 계산
        for code_num in range(0, len(my_min_enviro_code)):
            if my_min_enviro_code[code_num] == 'EV4222':  # 대기질 연구 동향(1)
                min_enviro_credit += 1
            else:
                min_enviro_credit += 3
        # 지환공 과목 토큰 관련 식
        if nec_min_enviro_count >= 3:
            nec_min_enviro_token = True
        if is_completed(my_min_enviro_code, 'EV3101'):
            min_enviro_engine_token = True
        if min_enviro_credit >= 15:
            min_enviro_credit_token = True
        if nec_min_enviro_token and min_enviro_engine_token and min_enviro_credit_token:
            min_enviro_token = True

    # 부전공(주전공 이외) 관련 표 데이터
    layout6 = go.Layout(title='<b>분야별 부전공 달성 현황<br>(7개 주전공 관련, 특별 인정 제외 2천번대 전공 교과목 불인정)</b>',
                        titlefont=dict(family='맑은 고딕', size=title_font_size),
                        font=dict(family='맑은 고딕', size=15), paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                        dragmode=False)

    fig6_cell_color_list = []
    minor2_tf_list = (min_physics_token, min_physics_token, nec_min_physics_token, min_physics_credit_token), \
                     (min_chemi_token, min_chemi_token, nec_min_chemi_token, min_chemi_credit_token), \
                     (min_bio_token, min_bio_token, nec_min_bio_token, min_bio_credit_token), \
                     (min_eecs_token, min_eecs_token, nec_min_eecs_token, min_eecs_credit_token), \
                     (min_mecha_token, min_mecha_token, nec_min_mecha_token, min_mecha_credit_token), \
                     (min_mater_token, min_mater_token, nec_min_mater_token, min_mater_credit_token), \
                     (min_enviro_token, min_enviro_token, nec_min_enviro_token, min_enviro_credit_token)

    for num in range(0, 4):
        row_color_list = []
        for minor in minor2_tf_list:
            if minor[num] or num == 0:  # 첫 번째 열(Index)과 조건 만족 셀에는 색칠
                row_color_list.append(lighter_color2)
            else:
                row_color_list.append("#FFFFFF")
        fig6_cell_color_list.append(row_color_list)

    # fig6 구성 데이터 리스트
    fig6_index_data = []  # index
    fig6_minor_name = minor_list2  # 부전공 교과목명
    fig6_nec_data = []  # 필수과목 이수 현황
    fig6_credit_data = []  # 취득 학점 현황

    # index 내용 추가하기
    for i in range(0, len(minor2_tf_list)):
        fig6_index_data.append(str(i + 1))

    # 필수과목 이수 현황 추가하기
    # 1. 물리
    fig6_nec_data.append("전공필수 교과목: " + str(nec_min_physics_count) + "/3(과목)")
    fig6_credit_data.append("총 취득 학점: " + str(min_physics_credit) + "/15(학점)")

    # 2. 화학
    if int(my_number[:4]) <= 2017:  # 17학번 이전일 시
        fig6_index_data[1] += "<br>"
        fig6_minor_name[1] += "<br>"
        fig6_nec_data.append("전공필수 교과목: " + str(nec_min_chemi_count) + "/3(과목)<br>")
        fig6_credit_data.append("총 취득 학점: " + str(min_chemi_credit) + "/15(학점)<br>"
                                                                      "※ 2천번대 전공 교과목 인정(단, 유기화학I, 물리화학I 불인정)")
    else:
        fig6_nec_data.append("전공필수 교과목: " + str(nec_min_chemi_count) + "/3(과목)")
        fig6_credit_data.append("총 취득 학점: " + str(min_chemi_credit) + "/21(학점)")

    # 3. 생명과학
    fig6_index_data[2] += "<br>"
    fig6_minor_name[2] += "<br>"
    fig6_nec_data.append("전공필수 교과목: " + str(nec_min_bio_count) + "/2(과목)<br>"
                                                                 "전공필수 실험과목: " + str(nec_min_bio_exp_count) + "/1(과목)")
    fig6_credit_data.append("총 취득 학점: " + str(min_bio_credit) + "/15(학점)<br>")

    # 4. 전기전자컴퓨터
    fig6_index_data[3] += "<br>"
    fig6_minor_name[3] += "<br>"
    if nec_min_eecs_token:
        fig6_nec_data.append("전공필수 교과목 이수 여부: 이수<br>")
    else:
        fig6_nec_data.append("전공필수 교과목 이수 여부: 미이수<br>")

    if int(my_number[:4]) <= 2017:  # 17학번 이전일 시
        fig6_credit_data.append("총 취득 학점: " + str(min_eecs_credit3) + "/15(학점)<br>"
                                                                      "※ 알고리즘 개론, 컴퓨터 구조, 공학전자기학 교과목 이수 인정")
    else:
        fig6_credit_data.append("2천번대 취득 학점: " + str(min_eecs_credit2) + "/6(학점)<br>"
                                                                         "3천번대 이상 취득 학점: " + str(
            min_eecs_credit3) + "/12(학점)")

    # 5. 기계공학
    fig6_index_data[4] += "<br>"
    fig6_minor_name[4] += "<br>"
    fig6_nec_data.append("전공필수 교과목: " + str(nec_min_mecha_count) + "/3(과목)<br>")
    fig6_credit_data.append("총 취득 학점: " + str(min_mecha_credit) + "/15(학점)<br>"
                                                                  "※ 2천번대 전공 교과목 인정")

    # 6. 신소재공학
    if int(my_number[:4]) <= 2017:  # 17학번 이전일 시
        fig6_nec_data.append("전공필수 교과목: " + str(nec_min_mater_count) + "/3(과목)")
        fig6_credit_data.append("총 취득 학점: " + str(min_mater_credit) + "/15(학점)")
    else:
        fig6_nec_data.append("전공필수 교과목: " + str(nec_min_mater_credit) + "/6(학점)")
        fig6_credit_data.append("총 취득 학점: " + str(min_mater_credit) + "/9(학점)")

    # 7. 지환공학
    fig6_index_data[6] += "<br>"
    fig6_minor_name[6] += "<br>"
    if min_enviro_engine_token:
        fig6_nec_data.append("전공필수 교과목(실험과목 제외): " + str(nec_min_mecha_count) + "/3(과목)<br>"
                                                                                "\"환경공학\" 교과목 이수 여부: 이수")
    else:
        fig6_nec_data.append("전공필수 교과목(실험과목 제외): " + str(nec_min_mecha_count) + "/3(과목)<br>"
                                                                                "\"환경공학\" 교과목 이수 여부: 미이수")
    fig6_credit_data.append("총 취득 학점: " + str(min_enviro_credit) + "/15(학점)<br>")

    fig6 = go.Figure(data=[go.Table(
        columnwidth=[80, 100, 300],
        header=dict(values=['<b>Index</b>', '<b>부전공</b>', '<b>필수 교과목 이수 여부</b>', '<b>취득 학점</b>'],
                    fill_color=sub2_present_color,
                    align=['center', 'center', 'left', 'left'],
                    font=dict(color=fig5_header_color, size=25),
                    height=40),
        cells=dict(values=[fig6_index_data, fig6_minor_name, fig6_nec_data, fig6_credit_data],
                   fill_color=fig6_cell_color_list,
                   align=['center', 'center', 'left', 'left'],
                   font_size=20,
                   height=30))],
        layout=layout6
    )

    from plotly.offline import plot
    plot(fig1, filename=info_directory + '1. Credit_and_Mission_Graph.html', auto_open=True, config=config)
    plot(fig2, filename=info_directory + '2. Major_and_Entire_Grade.html', auto_open=True, config=config2)
    plot(fig3, filename=info_directory + '3. Completed_Mission.html', auto_open=True, config=config3)
    plot(fig4, filename=info_directory + '4. Uncompleted_Mission.html', auto_open=True, config=config3)
    plot(fig6, filename=info_directory + '5. Minor_Mission1.html', auto_open=True, config=config3)
    plot(fig5, filename=info_directory + '6. Minor_Mission2.html', auto_open=True, config=config3)

# executeAnalyse()
