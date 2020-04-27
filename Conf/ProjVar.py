import os

proj_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

PageElementLocator_file_path = os.path.join(proj_path,r"Conf\PageElementLocator.ini")

test_data_file_path = os.path.join(proj_path,r"TestData\126邮箱联系人.xlsx")


test_case_info_sheet = "126账号"
test_result_sheet = "测试结果"

#sheet为：126账号
index_col_no = 0
test_case_description_col_no = 1
test_steps_sheet_name_col_no =2
is_test_executed_flag_col_no = 3
test_time_col_no =4
test_result_col_no =5

#sheet为：联系人
test_step_index_col_no = 0
test_step_description_col_no =1
keyword_col_no=2
locator_col_no =3
value_col_no =4
is_test_step_executed_flag_col_no = 5
test_step_time_col_no =6
test_step_result_col_no =7



if __name__=="__main__":
    print(proj_path)
    print(PageElementLocator_file_path)