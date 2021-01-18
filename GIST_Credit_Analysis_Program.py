from os import remove, getcwd
from os.path import isfile
import AnalysisProcess
# import Result  # 파일 실행 경로에 UNICODE(한글 등)이 존재 시 Result.py가 실행되지 않는 문제 발생 - 자동으로 그래프가 보여지도록 설정
import SetupWizard
import subprocess


def safeRemove(file_directory):
    if isfile(file_directory):
        remove(file_directory)


program_directory = getcwd()[:getcwd().find('Design')]
program_file_directory = getcwd()
save_token = 1

analyze_token = SetupWizard.executeWizard()
if analyze_token != 1:
    AnalysisProcess.executeAnalyse()
    # save_token = Result.executeResult()  # Result 화면에서 그래프 파일 저장 선택창에서 No 선택 시 save_token = 0
    subprocess.Popen('explorer "' + program_directory + 'edited\\info"')

file = program_directory + 'edited\\info\\personalinfo.xlsx'  # 암호화된 ID, PW 포함된 개인정보
safeRemove(file)
file = program_directory+'edited\\info\\personalinfo2.xlsx'   # 사용자 이름, 학번, 전공, 수강 학기수 포함된 개인정보
safeRemove(file)
file = program_directory+'original\\grade\\grade.xls'   # 사용자 학점 파일(원본)
safeRemove(file)
file = program_directory+'edited\\grade\\grade_edited.xlsx'   # 사용자 학점 파일(수정)
safeRemove(file)
file = program_directory+'edited\\grade\\present_subject_info.xlsx'   # 현재 수강 과목 리스트
safeRemove(file)

# 수집된 정보 파일 삭제 명령(필요 시 수정 가능)
if save_token == 0:
    file = program_directory+'edited\\info\\first_graph.html'     # 사용자의 학점, 졸업요건 포함된 그래프
    safeRemove(file)
    file = program_directory+'edited\\info\\second_graph.html'    # 사용자의 평점 통계 그래프
    safeRemove(file)
    file = program_directory+'edited\\info\\first_table.html'     # 사용자가 달성한 졸업요건 그래프
    safeRemove(file)
    file = program_directory+'edited\\info\\second_table.html'    # 사용자가 미달성한 졸업요건 그래프
    safeRemove(file)

